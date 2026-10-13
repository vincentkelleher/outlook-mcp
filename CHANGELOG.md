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
- **Atomic token-cache writes**, and one re-read-and-retry on a silent-refresh
  failure, so several clients can share `~/.outlook-skill/msal-cache.json`.
- **`bin`: `outlook-mcp`**, `files`, `engines`, and a `prepare` build, for
  installs from GitHub. No npm publish, ever: this project is distributed from
  the repository alone.
- **Documentation**: `SECURITY.md`, this file, `docs/clients.md` with a
  per-client install matrix, `skills/outlook-skill/references/setup.md` for the
  Entra ID registration, and a rewritten README.
- **CI** on Node 22 and 24 through format, typecheck, unit tests on the sources,
  build, and end-to-end tests against the built entry; a Node 20 job that only
  smoke-tests the build, since native type stripping does not exist there; a
  guard that fails on any `Bun.` identifier; and an `npm pack` step asserting the
  release artefact carries `dist/`.

### Changed

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

### Known limitations

- Attachments of 3–150 MB need a Graph upload session that this version does not
  implement, so `add_attachment` refuses them rather than truncating.
- Sign-in still runs as a detached process; an in-process state machine is on the
  backlog.
- Streamable HTTP and native MCP OAuth are not implemented, so every client
  launches a local stdio server.
