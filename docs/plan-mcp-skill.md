# Plan: Outlook for AI agents — MCP server + portable skill

Locked decisions (2025 review):

| Decision | Choice |
| --- | --- |
| Runtime | **Node-only toolchain** — drop Bun entirely |
| Distribution | **GitHub-only, no npm** |
| Tool surface | Read-only today **+ draft writing** (the user clicks Send; nothing we expose can send mail) |
| Sign-in | **Background device-code login + `auth_status` polling** |

## 1. Where we are

Current repo is a **Bun CLI + root `SKILL.md`**:

| Piece | State |
| --- | --- |
| `scripts/outlook.ts` (223 lines) | Everything in one file: MSAL cache plugin, login/logout, Graph client, `list`/`get`/`search`/`attachments`, limit validator, Commander CLI. Functions are exported, so it is testable. |
| `SKILL.md` (repo root) | Bun commands + a hand-rolled `nohup` recipe for device-code sign-in. |
| `test/`, `e2e/` | Good coverage, but `test/outlook.test.ts` depends on Bun-only `mock.module` (2 modules, ~90 lines of fakes). |
| `package.json` | `private: true`, no `bin`, no build, Bun lockfile. |
| Bun-specific | `Bun.env` (L44), `Bun.argv` (L159), `import.meta.main` (L219), implicit `.env` loading, `bun.lock`, `bunfig.toml`, `bun:test`. |

Capabilities: list inbox (≤100), mailbox search (≤1000), read one message as text, list attachments with base64 `contentBytes`. Read-only scopes `User.Read`, `Mail.Read`, `email`, `profile`; public client, device-code flow, cache at `~/.outlook-skill/msal-cache.json` (`0700`/`0600`), authority hard-coded to `consumers` (personal accounts only).

## 2. Target architecture

**One runtime-agnostic core, two thin surfaces.** Tool definitions live in exactly one place and are consumed by both the MCP server and the CLI, so they cannot drift.

```
outlook-skill/
├── src/
│   ├── core/
│   │   ├── config.ts        # env + ~20-line .env loader (Bun auto-loaded this)
│   │   ├── auth.ts          # MSAL cache plugin (atomic write), token acquisition, logout, authStatus
│   │   ├── device-login.ts  # detached sign-in process + state file (pending | authorized | error)
│   │   ├── graph.ts         # messages, drafts, attachments — injectable deps, cursors, folder resolution
│   │   └── errors.ts        # AuthRequiredError -> isError + remediation text
│   ├── tools.ts             # SINGLE SOURCE OF TRUTH: name, title, description, annotations, schema, handler
│   ├── mcp.ts               # McpServer: registerTool over TOOLS, resources, instructions, stdout guard
│   ├── cli.ts               # Commander over core (+ login/logout/status/serve/config)
│   └── index.ts             # bin: no args -> MCP stdio; anything else -> CLI
├── skills/outlook-skill/
│   ├── SKILL.md             # moved from repo root (§7)
│   └── references/setup.md  # Entra app-registration steps
├── docs/clients.md          # per-client install matrix
├── .mcp.json                # Claude Code project-scope example
├── .vscode/mcp.json         # VS Code / Copilot example
├── test/                    # node:test against src/*.ts directly (no build needed)
└── e2e/                     # spawned stdio JSON-RPC + CLI assertions
```

`bin: { "outlook-mcp": "dist/index.js" }` with a `#!/usr/bin/env node` shebang. **Launching with no arguments serves MCP over stdio**, so the config shape every client understands works, while `outlook-mcp login` still works for humans. A thin `scripts/outlook.ts` shim re-exporting the CLI is kept for one release so anyone who copied the old command keeps working.

## 3. Node-only toolchain

Verified on this machine (Node v25.0.0): `node file.ts` and `node --test` run TypeScript directly via native type stripping. Type stripping is enabled by default from **v22.18.0 / v23.6.0** and stable from **v24.12.0 / v25.2.0**.

- **Dev and test run TS sources directly** — `node --test test/*.test.ts e2e/*.test.ts`. No ts-node, no vitest, no loader, no build step while iterating.
- **Ship compiled JS**: `tsc -p tsconfig.build.json` → `dist/`, so the runtime floor for *users* stays low (`engines.node >= 20`).
- **Erasable syntax only** — no `enum`, `namespace`, or constructor parameter properties (they need transformation). Enforce with `"erasableSyntaxOnly": true`.
- **Explicit import extensions** (`./core/auth.ts`) so Node runs the sources; `"allowImportingTsExtensions": true` plus `"rewriteRelativeImportExtensions": true` (TS ≥5.7; repo already pins 5.9.3) makes `tsc` emit correct `.js` specifiers.
- **Remove Bun**: delete `bun.lock` + `bunfig.toml`, add `package-lock.json`, replace `Bun.env`/`Bun.argv`/`import.meta.main` (entry check = compare `process.argv[1]` realpath with `fileURLToPath(import.meta.url)`).
- **The big one: retire `mock.module`.** It has no `node:test` equivalent (Node's `mock` patches object properties, not ESM resolution). Refactor `test/outlook.test.ts` to **dependency injection** instead: `createAuth({ msal, cachePath })` and `createGraph({ token, fetchImpl })`, with fakes passed in. This removes toolchain lock-in and is better design anyway — but budget it as real work, not a mechanical port.
- Scripts: `dev`, `test`, `test:e2e`, `build`, `typecheck`, `format`, `format:check` (prettier stays), `prepare: npm run build` (makes git installs self-building).
- CI (GitHub Actions): matrix Node 20 / 22 / 24 — `typecheck`, `test`, `build`, then a dist smoke (`node dist/index.js --help`, spawned-stdio e2e). Grep guard: fail if `Bun.` appears in `src/`, `test/`, `e2e/`, or `skills/`.

## 4. MCP server design

**SDK / spec:** `@modelcontextprotocol/sdk@^1.32` — negotiates protocol `2025-11-25` down to `2024-10-07`, i.e. the broadest client compatibility today. The v2 line (`@modelcontextprotocol/server`, spec `2026-07-28`) is weeks old: track it, keep `tools.ts` transport-agnostic so the swap stays small.

Tool names stay short (`list_emails`) because clients prefix them — Pi exposes `mcp__outlook__list_emails`.

**Read tools** — `readOnlyHint: true`, `destructiveHint: false`, `idempotentHint: true`, `openWorldHint: true`:

| Tool | Inputs | Notes |
| --- | --- | --- |
| `list_emails` | `folder?` (default `inbox`), `unread?`, `limit?` 1-100, `cursor?` | Adds folder resolution and paging, both missing today. |
| `search_emails` | `query`, `limit?`, `cursor?` | Graph syntax (`from:`, `subject:`, `body:`), keeps `ConsistencyLevel: eventual`. |
| `get_email` | `id` | Full plain-text body via `Prefer: outlook.body-content-type="text"`. |
| `list_attachments` | `messageId` | **Metadata only** — `id`, `name`, `contentType`, `size`, `isInline`. |
| `download_attachment` | `messageId`, `attachmentId`, `dir?`, `encoding?` | Writes to an approved directory (default `$TMPDIR/outlook-mcp`), returns the path. Base64 inline only on request, size-capped. |
| `auth_status` | — | `{ signedIn, account, authority, cachePath, lastLoginError? }` — lets an agent self-diagnose. |
| `login` / `logout` | — | `login` returns instantly (§5); `logout` is `destructiveHint: true`. |

**Draft tools** — the user's "just click Send" requirement. Verified against Microsoft Graph: `POST /me/messages` creates a draft, and `POST /me/messages/{id}/createReply` / `/createForward` create threaded draft replies/forwards (a `comment` body sets the text). All are `readOnlyHint: false`, `idempotentHint: false`, `openWorldHint: true`, so clients prompt before running them.

| Tool | Backing call | Notes |
| --- | --- | --- |
| `create_draft` | `POST /me/messages` | `to[]`, `cc[]`, `bcc[]`, `subject`, `body`, `bodyFormat: text\|html`, `inReplyTo`, `categories`. Returns `id` + `webLink`. |
| `create_reply_draft` | `POST /me/messages/{id}/createReply` | Preserves threading; draft, never sent. |
| `create_forward_draft` | `POST /me/messages/{id}/createForward` + explicit `to[]`. |
| `update_draft` | `PATCH /me/messages/{id}` | Let the agent revise after user feedback instead of creating duplicates. |
| `list_drafts` | `GET /me/mailFolders/drafts/messages` | So the agent can find what it already made. |
| `delete_draft` | `DELETE /me/messages/{id}` | `destructiveHint: true`. |
| `add_attachment` | `POST /me/messages/{id}/attachments` | **Hard 3 MB cap** (verified: single-POST attachments must be under 3 MB; 3–150 MB needs `createUploadSession`, which is backlog). Over-limit input returns a clear actionable error, not a truncated upload. |

**Things we deliberately do not implement: `/send`, `/reply`, `/forward`.** Nothing in the tool registry or `graph.ts` can transmit mail — that is the whole point of the design and is enforced by a test that asserts (a) the registry contains no send-capable tool and (b) no request path ends in `/send`, `/reply`, or `/forward`. Document this invariant in the README and SECURITY.md; it is the strongest thing we can say to a cautious user.

**Scope consequence:** creating drafts requires **`Mail.ReadWrite`** (verified, delegated, personal and work/school), replacing `Mail.Read`. Every existing user re-consents once on next sign-in; `--readonly` drops back to `Mail.Read` and hides the draft tools from `tools/list` for users who want read-only.

Cross-cutting:

- **Auth failures are the product.** Missing/expired session returns `isError: true` with the exact next step ("call `login`, show the user this URL and code, then retry"), never a stack trace.
- **Both result shapes:** compact JSON in `content[].text` (no 2-space indent — saves tokens) and `structuredContent` for typed clients. Pi truncates text results over 20 KB, so defaults are small and full bodies only come from `get_email`.
- **No base64 by default.** Today `emails attachments` inlines `contentBytes`; that is a token bomb through MCP. CLI keeps its behaviour for back-compat; MCP does not.
- **stdout purity.** stdio carries JSON-RPC; stderr carries diagnostics. Serve mode redirects `console.log` to stderr at startup, and a test asserts no non-JSON bytes reach stdout.
- **Server `instructions`** on initialize (Pi uses it as the server description; Claude surfaces it).
- **Resources** for clients that browse rather than call: `outlook://folders`, `outlook://inbox?unread=&limit=`, `outlook://messages/{id}`. Prompts: skipped in v1, support is uneven across CLI clients.

## 5. Sign-in inside MCP

A device-code flow blocks up to 15 minutes while MSAL polls; clients won't wait (Pi's per-request MCP timeout is 60 s), so a blocking `login` tool would time out and orphan the flow.

1. `login` **spawns a detached sign-in process** — the `nohup` dance from `SKILL.md`, moved into the product — writes `{ state, verificationUri, userCode, expiresAt }` to a state file, and returns immediately.
2. The agent shows the URL and code, then polls `auth_status` (cheap, idempotent).
3. The next call succeeds silently.

Reuses the already-tested MSAL path. Alternative if orphan processes ever prove annoying: call the `devicecode` and `token` endpoints directly with `fetch` for in-process two-step state (~80 lines). Later: optional streamable-HTTP + MCP OAuth 2.1 so clients sign in natively (`pi mcp login outlook`, `codex mcp login outlook`).

**Token cache contention is a real risk once several clients share one cache.** `afterCacheAccess` writes non-atomically today, so two agents refreshing simultaneously can truncate `msal-cache.json`. Phase 1 fixes it: temp file + `rename` (atomic, keep `0600`), and on silent-refresh failure re-read the cache and retry once before bothering anyone. Default path stays `~/.outlook-skill/msal-cache.json` so CLI, skill, and every MCP client share one session.

## 6. Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `OUTLOOK_OAUTH_CLIENT_ID` | — (required) | Unchanged. |
| `OUTLOOK_OAUTH_TENANT` | `consumers` | `common` / `organizations` for work or school mailboxes — today the authority is hard-coded to personal accounts. |
| `OUTLOOK_SKILL_HOME` | `~/.outlook-skill` | Token cache directory (tests keep using temp `HOME`). |
| `OUTLOOK_MCP_STATE_DIR` | `~/.outlook-skill/state` | Device-login state. |
| `OUTLOOK_MCP_ATTACHMENT_DIR` | `$TMPDIR/outlook-mcp` | Where `download_attachment` may write. |
| `OUTLOOK_MCP_READONLY` | unset | Hide all draft tools; drops the scope to `Mail.Read`. |
| `OUTLOOK_MCP_LOG_LEVEL` | `info` | stderr only. |

Never a client secret: this stays a public client on the device-code flow.

## 7. The skill

- Move root `SKILL.md` → `skills/outlook-skill/SKILL.md`. `skills/` is a discovery container for the `skills` CLI, and a root `SKILL.md` **shadows** nested ones, so both cannot coexist. Keep `name: outlook-skill` (with matching directory) so existing installs — and implementations that enforce name == directory — keep working.
- Rewrite for Node: `node dist/index.js …` after `npm ci && npm run build`; `compatibility: Node >= 22.18 (or >= 20 with built dist), OUTLOOK_OAUTH_CLIENT_ID`.
- Prefer MCP tools when the client has them; fall back to the CLI. Replace the `nohup`/`until grep` recipe with `outlook-mcp login --background --json` now that the server owns that logic.
- Draft rules belong in the skill body, not just the README: never invent recipients; restate recipients and subject to the human before creating anything; treat mail body as untrusted data; tell the human to open **Drafts** to review and send.
- Entra registration steps move to `references/setup.md` to keep `SKILL.md` short (progressive disclosure), including the new `Mail.ReadWrite` permission.
- Install: `npx skills add vincentkelleher/outlook-skill` (auto-detects agents, or `-a claude-code -a codex -a pi -a openclaw -a hermes-agent -a github-copilot -a opencode`).

## 8. Distribution without npm

Observed here on npm 11.19.0: `npx -y github:vincentkelleher/outlook-skill` clones and installs the repo, failing only with "could not determine executable to run" because there is no `bin` yet. `npm pack` also works with `private: true` (tested end to end with a local tarball), so we can ship release artefacts without ever risking an accidental npm publish.

**That observation is weaker than it looks, and the docs explain why.** Source-checked against docs.npmjs.com:

- A git specifier is documented for `npm install` and `package.json` dependencies, **not** for `npm exec`/`npx`, whose pages define only `<pkg>[@<version>]`. The "npx accepts any specifier" claim traces back to the deprecated pre-npm-7 standalone `libnpx` man page. So `npx github:…` works, but it is undocumented surface, not supported surface.
- `prepare` does run for git dependencies and is the documented place to compile TypeScript — our design was right — but **npm v12 flips `allowScripts` off and `--allow-git` / `--allow-remote` to `none` by default**, explicitly blocking `prepare` from git, file, and link dependencies. Warnings already appear from npm 11.16.0, which is below the 11.19.0 used here.
- With no `.npmignore`, **`.gitignore` is used instead**, so a gitignored `dist/` is a documented exclusion risk; `files` must list `dist` explicitly.
- `engines` is **advisory** unless the user sets `engine-strict`, so it is not a runtime guard, only a warning.
- `npm install <tarball url>` *is* documented, with testable requirements: `.tar`/`.tar.gz`/`.tgz`, one wrapping directory, and a `package.json` with `name` and `version`. Crucially **`prepare` does not run for tarball-URL installs**, so the release asset must be an `npm pack` artefact and never GitHub's auto-generated source archive — attaching the wrong file is the most likely failure mode.
- Whether `npx` re-clones a git specifier per invocation is **UNVERIFIED**; npm's caching note covers package names only.

Paths, in order of how defensible they are:

1. **Clone + build (the only path we call supported).** `git clone && npm ci && npm run build`, then point clients at the absolute `dist/index.js`. Uses only documented behaviour and is immune to the v12 default flips.
2. **`outlook-mcp config <client> [--write]`** — added precisely because there is no dependable registry one-liner: resolves its own absolute path and prints (or idempotently writes) the exact snippet for `pi`, `claude-code`, `claude-desktop`, `codex`, `vscode`, `cursor`, `gemini-cli`, `opencode`, or `generic`. Removes the most common failure mode, which is hand-editing JSON/TOML with a wrong path.
3. **Best-effort one-liners, labelled as such.** `npx -y github:vincentkelleher/outlook-skill#v1.0.0` and an `npm pack` artefact attached to the GitHub Release. Never describe these as supported. Each carries its warnings inline: may be blocked by `--allow-git` / `--allow-remote` / `allowScripts` policy; needs `git` on the PATH for the git form; pin to an immutable tag or sha rather than a floating branch that npm rebuilds every install; and a cold clone-and-install can exceed a client's MCP startup budget (Codex defaults to 10 s). CI must generate the artefact with `npm pack` and assert that it contains `dist/` and a `package.json` with `name` and `version`.

### Verified client rules (source-checked; `docs/clients.md` and the `config` command must follow these)

Four clients were verified against official documentation. `<DIST>` means the absolute path to `dist/index.js`.

| Client | Server name | Add command | Env strategy | Units that bite |
| --- | --- | --- | --- | --- |
| Pi | `outlook` | `pi mcp add outlook --env OUTLOOK_OAUTH_CLIENT_ID="$OUTLOOK_OAUTH_CLIENT_ID" --exposure direct --description "Read, search, and draft Outlook mail" -- node <DIST>` then `pi mcp list` to verify | `"${OUTLOOK_OAUTH_CLIENT_ID}"` in `env` is documented | `timeout` is **seconds** (default 60); first prompt waits 10 s for `direct` tools only |
| Claude Code | `outlook` | `claude mcp add outlook --scope user --env OUTLOOK_OAUTH_CLIENT_ID="…" --transport stdio -- node <DIST>` | `${VAR}` expansion documented for `.mcp.json`; **use a literal value at user scope** (`~/.claude.json` expansion is undocumented) | `MCP_TIMEOUT` (ms) startup; per-server `timeout` in **ms**; 30 min stdio idle timeout |
| Codex / ChatGPT / Codex IDE | `outlook` | `codex mcp add outlook --env OUTLOOK_OAUTH_CLIENT_ID="…" -- node <DIST>` — writes **global** `~/.codex/config.toml` only | **No interpolation is documented for TOML `env`.** To avoid storing the value, use `env_vars = ["OUTLOOK_OAUTH_CLIENT_ID"]` | `startup_timeout_sec` default **10**, `tool_timeout_sec` default 60 |
| VS Code / Copilot | `outlook` | `code --add-mcp "{\"name\":\"outlook\",\"type\":\"stdio\",\"command\":\"node\",\"args\":[\"<DIST>\"]}"` — writes the user profile; no scope flag, no `--` separator | `${input:id}` plus a top-level `inputs` array is the documented mechanism; `${env:NAME}` is **not** documented for `mcp.json` | No documented timeout; Workspace Trust gates workspace servers, and Agent Host **silently skips** `${input:…}` servers |

Hard rules that come out of this:

- **Always emit `"type": "stdio"` for VS Code.** Its reference table marks `type` required while its own examples omit it — emit it to satisfy the stricter reading.
- **In Claude Code, never put the server name straight after `--env`.** The CLI reads the name as another `KEY=value` pair and rejects it; separate them with `--transport stdio`.
- **A `--` separator is required** for Pi, Claude Code, and Codex; **not** for `code --add-mcp`, which takes one JSON string.
- **Do not write `${VAR}` into Codex TOML.** Forward with `env_vars`, or write a literal.
- **One name for everything: `outlook`.** Pi and Claude Code enforce `[A-Za-z0-9_-]`, Codex additionally permits `: @ / .`, and VS Code recommends camelCase with no enforced charset — so `outlook` is safe everywhere and never needs prefixing in our snippets.
- **Cold `npx` installs are now known to be risky, not merely slow**: Codex's default `startup_timeout_sec` is 10 s and Claude Code has an `MCP_TIMEOUT` startup budget, which a clone-and-install can exceed. Clone-and-build stays the documented default; the git one-liner must tell users to raise those two values.
- Tool naming is prefixed to `mcp__<server>__<tool>` by Pi and Claude Code; for Codex and VS Code **no official prefix format is documented**, so never hardcode a prefixed tool name in docs or the skill.

Still unverified (do not present as fact in docs): the absolute VS Code user-profile `mcp.json` path (docs say only "the `mcp.json` file in your user profile folder"), whether `${env:NAME}` expands in `mcp.json`, whether Codex accepts SSE, and any Codex/VS Code tool-call naming. A second research pass is checking Cursor, Gemini CLI, OpenCode, Claude Desktop, and whether the `npx` git/tarball install claims are documented behaviour or convention.

## 9. Security posture

Public client, no secret, `0700`/`0600` cache with atomic writes, no telemetry, mail content only reaches the model the user is already talking to. Adding drafts widens the surface, so:

- **Cannot send.** No `/send`, `/reply`, `/forward` anywhere — enforced by test (§4).
- **A draft is not a sent message.** Mail content reaching the model is already the baseline of a read-only mail server, and a draft moves nothing the human has not confirmed. Guardrails stay behavioural and cheap: write tools carry annotations so clients prompt first, SKILL.md states that mail bodies are data and that a human confirms recipients and subject before a draft is created, and every draft result echoes back what was stored.
- **Recipients are never inferred.** Tools take explicit addresses; no tool reads the mailbox to guess who to mail.
- **Draft readback.** `create_draft` returns the stored subject, body, and recipients so the human can audit what was actually written.
- **Attachment bounds.** 3 MB cap on the JSON upload path, path-escape and symlink checks, size caps, writes confined to allowlisted directories.
- `logout` and `delete_draft` are marked `destructiveHint: true`.
- Anything that eventually sends mail needs its own opt-in (`Mail.Send`), a confirm parameter, and a separate discussion — out of scope here.

## 10. Delivery phases

**Phase 0 — scaffold (½ day).** This document, branch, `package.json` (bin/exports/engines/scripts), `tsconfig.json` + `tsconfig.build.json`, CI skeleton, prettier config.

**Phase 1 — Node-only toolchain + core extraction (1.5 days).** Split `scripts/outlook.ts` into `src/core/*`; env/argv injection and entry check; `.env` loader; tenant config; atomic cache write + refresh retry; folder + cursor support; **the `mock.module` → dependency-injection test rewrite**; remove Bun lockfiles. Acceptance: `node --test` green, `npm run typecheck` green, `node dist/index.js emails list --help` runs, existing CLI e2e assertions unchanged, CI `Bun.` grep clean.

**Phase 2 — MCP server (1 day).** `tools.ts` registry + `mcp.ts` (read tools, auth tools, resources, instructions, stdout guard), `isError` remediation, structured output, `bin`. Acceptance: `npx @modelcontextprotocol/inspector --cli node dist/index.js --method tools/list` lists every tool; in-memory `Client` tests cover happy path, auth-required path, argument validation; spawned-stdio test asserts the `initialize` handshake **and** byte-exact stdout purity; a 30 KB body stays inside the text budget.

**Phase 3 — drafts (1 day).** Create/reply/forward/update/list/delete + `add_attachment` with the 3 MB guard; `--readonly`; scope bump to `Mail.ReadWrite`; **no-send invariant tests**. Acceptance: a draft created by a test fake appears with the right recipients/subject, an over-limit or disallowed recipient fails with a clear message and no partial write, and the invariant tests pass.

**Phase 4 — distribution (1 day).** `config <client> [--write]`; README rewritten MCP-first with the skill section kept; `skills/` move + `npx skills add .` verified; `docs/clients.md` verified for real on Pi, Claude Code, Codex, and VS Code; CI job that runs `npm pack` and asserts the artefact contains `dist/` plus a `package.json` with `name` and `version`, so a wrong release asset cannot ship; SECURITY.md, CHANGELOG.md, migration note for the old `scripts/outlook.ts` path.

Total ≈ 5 working days.

**Backlog (Phase 5).** Upload sessions for 3–150 MB attachments; mark read/unread, move, folder management; streamable-HTTP + MCP OAuth; MCP SDK v2 once clients adopt spec `2026-07-28`; `.mcpb` bundle for Claude Desktop; `search` result counts; in-process device-code state machine (drop the detached process).

## 11. Risks

| Risk | Mitigation |
| --- | ---|
| Retiring `mock.module` balloons into a rewrite | Budgeted as its own Phase 1 deliverable; DI shape designed before touching tests. |
| Node type stripping rejects a construct (`enum`, `namespace`, parameter properties) | `erasableSyntaxOnly: true` in tsconfig + CI on Node 20/22/24. |
| Blocking login exceeds client MCP timeouts | Detached sign-in process, instant return, `auth_status` polling (§5). |
| stdout contamination corrupts JSON-RPC | stdout→stderr redirect in serve mode + byte-level e2e test. |
| Concurrent clients corrupt the shared token cache | Atomic temp+rename write, re-read-and-retry on refresh failure. |
| Read + draft lets hostile mail steer draft content | Human must click Send before anything leaves the mailbox, client confirmation annotations, SKILL.md rules, draft readback in every result (§9). |
| Scope bump to `Mail.ReadWrite` surprises existing users | Documented in README + CHANGELOG, one-time re-consent, `--readonly` escape hatch. |
| Base64 attachments exhaust context | Metadata-only by default, explicit capped download. |
| 3 MB Graph attachment limit hit silently | Hard input validation with an actionable error; upload sessions in backlog. |
| `webLink` is unreliable for personal outlook.com accounts (reported on Microsoft Q&A) | Return it opportunistically; SKILL.md tells the human to open Drafts instead. |
| GitHub-only means no registry semver/resolution | Tag releases, pin `#v1.0.0` in configs, ship release tarballs. |
| Cold `npx github:` install exceeds MCP startup patience | Recommend clone+build as the default path; release tarball as the fast option. |
| Pi defaults MCP tools to `codemode` (not declared to the model) | Document `"exposure": "direct"` for a server this small. |
| Existing skill users break when root `SKILL.md` moves | Keep `name: outlook-skill`, shim `scripts/outlook.ts` for one release, migration note. |
| Graph quirks (`$search` needs `ConsistencyLevel`, `$top` ≤ 1000, `$search` + `$filter` conflicts) | Encode limits in schemas, document, unit-test request shapes as the suite does today. |
