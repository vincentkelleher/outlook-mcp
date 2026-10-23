# Changelog

All notable changes to this project. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## 0.1.0 — not yet tagged

The conversion from a Bun command-line script into a Node MCP server that an
agent can drive. Everything published before this was the Bun CLI.

### Added

- **MCP server.** Launch with no arguments (or `serve`) and it speaks Model
  Context Protocol over stdio: 15 tools, `outlook://folders`,
  `outlook://inbox` and `outlook://messages/{id}` resources, server
  `instructions`, and a stdout guard so nothing but JSON-RPC reaches the
  transport.
- **Streamable HTTP transport.** `serve --http` serves the same registry on
  `/mcp` with the SDK's session handling, and answers `GET /health` with version,
  mode and session count — never mail. The bind defaults to `127.0.0.1:3939`
  (`OUTLOOK_MCP_HOST`, `OUTLOOK_MCP_PORT`, `--port 0` for a free port); every
  request is admitted on its `Host` and `Origin`, because the transport
  authenticates nobody; bodies are capped at 4 MiB, sessions at 32, and a session
  with nothing connected retires after `OUTLOOK_MCP_SESSION_IDLE_MINUTES`.
- **A Docker image**, published to `ghcr.io/vincentkelleher/outlook-skill` on tag
  for `linux/amd64` and `linux/arm64`: multi-stage, `node:24-alpine`, no dev
  dependencies, running as `node`, dumb-init as PID 1 so SIGTERM lands and the
  detached sign-in child is reaped, and `/data` as the only state — session, sign-in
  state, attachments. `docker build -t outlook-skill .` builds the same file CI
  publishes from.
- **`config docker`**, printing the remote entry (`{"type":"http",
"url":"http://127.0.0.1:3939/mcp"}`) with the run command, the
  `docker exec … outlook login` sign-in step and the loopback warning around it.
  `--write` refuses it: no client file owns a URL for "Docker".
- **Read tools**: `list_emails` (with folder resolution and cursor paging),
  `search_emails`, `get_email`, `list_attachments` (metadata only),
  `download_attachment` (confined to an allowed directory, symlink-checked, base64
  inline only on request and only up to 64 KiB), `auth_status`, `login`,
  `logout`.
- **Draft tools**: `create_draft`, `create_reply_draft`, `create_forward_draft`,
  `update_draft`, `list_drafts`, `delete_draft`, `add_attachment` (hard 3 MB
  cap). Every write reads back what Microsoft stored and says `sent: false`.
- **The no-send invariant.** Nothing in the tool registry or the Graph layer can
  transmit mail; `test/no-send.test.ts` asserts both halves of that.
- **Background sign-in.** `login` detaches the device-code flow and returns the
  verification URL and user code at once; agents poll `auth_status`. The CLI
  exposes the same behaviour as `login --background --json`.
- **`OUTLOOK_MCP_READONLY`**, an environment switch that removes every draft tool
  and drops the requested scope to `Mail.Read`. Deliberately not a command-line
  flag, so the permissions and the advertised tools cannot disagree.
- **`config <client> [--write]`**, printing or idempotently installing the
  snippet for Pi, Claude Code, Claude Desktop, Codex, VS Code, Cursor, Gemini
  CLI, OpenCode and generic stdio clients, each in its own format.
- **Leveled logs on stderr.** `OUTLOOK_LOG_LEVEL` (`debug`, `info` — the default
  — `warn`, `error`, `silent`) and `OUTLOOK_LOG_FORMAT` (`text`, `json`) over
  pino: one line per command, tool call, resource read, sign-in and HTTP session,
  with the Microsoft Graph calls, token-cache reads and state-file transitions at
  `debug`, each timed. stdout is never touched, in either format. Mail stays out
  of it — no token, body or attachment content at any level (redacted before
  serialisation, so the human format cannot leak what the JSON one hides), sender
  addresses masked and ids shortened until somebody asks for `debug`. Scoped
  loggers (`cli`, `mcp`, `http`, `auth`, `graph`, `login`, `storage`, `client`) and
  a session id on every line from an HTTP session say who did what.
- **Atomic token-cache writes**, and one re-read-and-retry on a silent-refresh
  failure, so several clients can share `~/.outlook-skill/msal-cache.json`.
- **`bin`: `outlook-mcp`**, `files`, `engines`, and a `prepare` build, for
  installs from GitHub. No npm publish, ever: this project is distributed from
  the repository alone.
- **Documentation**: `SECURITY.md`, this file, `docs/clients.md` with a
  per-client install matrix, `skills/outlook-skill/references/setup.md` for the
  Entra ID registration, and a rewritten README.
- **CI for the image** in `.github/workflows/docker.yml`: build, run, and check
  the health route, the `Host`/`Origin` refusals, and that the layer carries no
  `.env` and no session, on every pull request; multi-arch push to GHCR on the
  default branch (`latest`, `sha-…`) and on `v*` tags (semver).
- **CI** on Node 22 and 24 through format, typecheck, unit tests on the sources,
  build, and end-to-end tests against the built entry; a Node 20 job that only
  smoke-tests the build, since native type stripping does not exist there; a
  guard that fails on any `Bun.` identifier; and an `npm pack` step asserting the
  release artefact carries `dist/`.

### Changed

- **`serve` is a real command now.** It used to be intercepted before the CLI
  saw it, which meant `serve --help` started a server instead of printing help.
  No arguments still serves stdio exactly as before, and so does a bare `serve`.
- **Node only.** Bun is gone: no `Bun.` identifier, no `bun.lock`, no
  `bunfig.toml`. Sources run directly under Node's native type stripping
  (Node 22.18+ or 23.6+), and `npm run build` emits `dist/` for Node 20.
- **The skill moved** from `SKILL.md` in the repository root to
  `skills/outlook-skill/SKILL.md`, because a root `SKILL.md` shadows nested ones
  for the skills installer. The `name` is unchanged, so an existing install keeps
  working — re-run `npx skills add` to pick up the new copy.
- **`scripts/outlook.ts` was split** into `src/core/{config,auth,device-login,graph,errors}.ts`
  with `src/tools.ts` as the single tool registry, and deleted. There is no
  compatibility shim.
- **Sign-in now requests `Mail.ReadWrite` instead of `Mail.Read`,** because a
  draft is a write. **Existing users are asked to consent once on next sign-in.**
  `OUTLOOK_MCP_READONLY=1` keeps the old `Mail.Read` scope and its read-only
  surface.
- **Tests no longer depend on a test runner's module mocking**: `createAuth` and
  `createGraph` take their dependencies, so the fakes are ordinary objects and the
  suite is `node:test` on the sources.

### Removed

- `scripts/outlook.ts`, the root `SKILL.md`, `bun.lock`, `bunfig.toml`, and the
  unused module-level `accessToken`/`authStatus` convenience wrappers.

### Fixed

Findings from review of this branch, all of them shipped before tagging.

- **A sign-in that could never finish.** A session cached under the old
  `Mail.Read` scope made `auth_status` report `signedIn: true` while every draft
  tool failed for want of `Mail.ReadWrite`, and `login` replied "already signed
  in" — so an agent looped. `auth_status` now confirms a token is actually
  obtainable for the scopes in force, reports `needsConsent` when it is not, and
  `login` starts a device-code flow in that state.
- **A sign-in state that could not expire.** `expiresAt` was written with a
  pending device code and never read, so a helper process that died left every
  client telling the user to enter a dead code. Expired pending states are now
  ignored, `auth_status` reports a pending code, a second `login` reuses the live
  flow instead of overwriting it, and a spawn that fails is reported as failed
  rather than as slow.
- **A cursor that was trusted as a URL.** `cursor` went straight into the Graph
  request, so a hostile value could aim it, bearer token included, at any host.
  Scheme, host and path are now checked first; a genuine `nextLink` still works.
- **An attachment name that escaped its directory.** `dir` was confined, symlink
  included, but the file name came from the sender's file and was joined
  unchecked, so `../` wrote outside the attachment directory. Names are reduced
  to a safe basename and the resolved path is re-checked before the write.
- **`OUTLOOK_MCP_LOG_LEVEL` is gone**, and there is a logger now: the knob lives
  on as `OUTLOOK_LOG_LEVEL`, read by every surface, with the dead name neither
  documented nor honoured (a test pins that, so a variable nobody reads cannot be
  advertised as a control again).
- Smaller: leftover plaintext cache copies are swept, the sign-in state file is
  `0600` in a `0700` directory, `logout` clears the sign-in state too, page
  limits and tool-name lists have one source instead of four, `list_drafts` no
  longer advertises itself as a write, and unused exports are private.

### Known limitations

- Attachments of 3–150 MB need a Graph upload session that this version does not
  implement, so `add_attachment` refuses them rather than truncating.
- Sign-in still runs as a detached process; an in-process state machine is on the
  backlog.
- Streamable HTTP and native MCP OAuth are not implemented, so every client
  launches a local stdio server.
