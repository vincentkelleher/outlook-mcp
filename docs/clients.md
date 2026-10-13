# Installing on each client

`<DIST>` is the absolute path to the built entry,
`<checkout>/outlook-skill/dist/index.js`. The server name is `outlook`
everywhere: Pi and Claude Code require `[A-Za-z0-9_-]`, Codex additionally allows
`: @ / .`, and VS Code recommends camelCase without enforcing it, so one name
works for all of them.

## The supported path: clone and build

```bash
git clone https://github.com/vincentkelleher/outlook-skill
cd outlook-skill
npm ci
npm run build
```

Then point the client at `<DIST>`. This is the only route this repository calls
supported: it uses documented behaviour only, and it is immune to the npm v12
default changes described below.

### The shortcut: `config`

Hand-editing JSON and TOML with a wrong path is the most common failure here, so
the server prints the snippet itself, with its own absolute path resolved:

```bash
node dist/index.js config codex
node dist/index.js config pi
node dist/index.js config claude-code --write
```

Clients: `pi`, `claude-code`, `claude-desktop`, `codex`, `vscode`, `cursor`,
`gemini-cli`, `opencode`, `generic`. `--write` merges the entry into that
client's user configuration file: it creates the file when it is missing, keeps
every unrelated key, is idempotent, and refuses with a clear message when an
existing file does not parse — it never overwrites a configuration file it cannot
parse. When a client's file location is not one we can state with confidence,
`--write` refuses and tells you to paste instead.

If `dist/index.js` is missing, `config` warns and falls back to the TypeScript
sources, which need Node 22.18 or newer.

## The one-liners, and why they are labelled best-effort

Two shortcuts exist and neither is documented by npm as supported:

```bash
# git specifier: needs git on the PATH
npx -y github:vincentkelleher/outlook-skill#v0.1.0

# release tarball produced by `npm pack` in CI
npm install https://github.com/vincentkelleher/outlook-skill/releases/download/v0.1.0/outlook-skill-0.1.0.tgz
```

Checked against docs.npmjs.com:

- A git specifier is documented for `npm install` and for `package.json`
  dependencies, **not** for `npm exec` / `npx`, whose pages define only
  `<pkg>[@<version>]`. The `npx github:…` form works, but it is convention, not
  documented surface.
- **npm v12 flips `allowScripts` off and `--allow-git` / `--allow-remote` to
  `none` by default**, which blocks the `prepare` script that builds `dist/`.
  Warnings about this already appear from npm 11.16.0. A git or remote install
  that silently skips `prepare` has no `dist/` to run.
- Whether `npx` re-clones a git specifier per invocation is **unverified**: npm's
  caching note covers package names only.
- `npm install <tarball url>` **is** documented, with requirements: `.tar`,
  `.tar.gz` or `.tgz`, one wrapping directory, and a `package.json` with `name`
  and `version`. **`prepare` does not run for a tarball install**, so the release
  asset must be an `npm pack` artefact. GitHub's auto-generated source archive
  contains no build output and will not build one: attaching the wrong file is
  the most likely failure mode, which is why CI packs the artefact and asserts it
  contains `dist/`.
- Without `.npmignore`, npm falls back to `.gitignore`, so the gitignored `dist/`
  is a documented exclusion risk; `package.json` lists `dist` in `files` to
  override that.
- `engines` is advisory unless the user turns on `engine-strict`. It is a warning,
  not a guard.
- Pin to an immutable tag or sha (`#v0.1.0`), never a floating branch: npm
  rebuilds a git dependency for every install.

And a cold clone-and-install can outlast a client's patience: Codex gives a
server 10 seconds to start, Claude Code has an `MCP_TIMEOUT` startup budget. If
you use a one-liner, raise both.

## Matrix

| Client               | Add command                                                                                                                                                                             | Environment variable                                                                    | Units that bite                                                                               |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| Pi                   | `pi mcp add outlook --env OUTLOOK_OAUTH_CLIENT_ID="$OUTLOOK_OAUTH_CLIENT_ID" --exposure direct --description "Read, search, and draft Outlook mail" -- node <DIST>`, then `pi mcp list` | `"${OUTLOOK_OAUTH_CLIENT_ID}"` in `env` is documented                                   | `timeout` is **seconds** (default 60); the first prompt waits 10 s, for `direct` tools only   |
| Claude Code          | `claude mcp add outlook --scope user --env OUTLOOK_OAUTH_CLIENT_ID="…" --transport stdio -- node <DIST>`                                                                                | `${VAR}` documented in `.mcp.json`; write a literal at user scope                       | `MCP_TIMEOUT` (ms) at startup; per-server `timeout` in **ms**; 30 min stdio idle              |
| Claude Desktop       | no add command; edit its config file and restart the app                                                                                                                                | literal value                                                                           | the app reads its config at launch only                                                       |
| Codex                | `codex mcp add outlook --env OUTLOOK_OAUTH_CLIENT_ID="…" -- node <DIST>` (writes global `~/.codex/config.toml` only)                                                                    | **no interpolation documented for TOML `env`**: forward with `env_vars`                 | `startup_timeout_sec` default **10**, `tool_timeout_sec` default 60                           |
| VS Code / Copilot    | `code --add-mcp "{\"name\":\"outlook\",\"type\":\"stdio\",\"command\":\"node\",\"args\":[\"<DIST>\"]}"`                                                                                 | `${input:id}` plus a top-level `inputs` array is documented; write an env value instead | no documented timeout; Workspace Trust gates workspace servers, Agent Host skips `${input:…}` |
| Cursor               | no add command; edit its `mcp.json`                                                                                                                                                     | literal value (unverified)                                                              | unverified                                                                                    |
| Gemini CLI           | `gemini mcp add outlook node <DIST> --scope user --env OUTLOOK_OAUTH_CLIENT_ID="…"` (unverified)                                                                                        | literal value (unverified)                                                              | unverified                                                                                    |
| OpenCode             | no add command; edit its config file                                                                                                                                                    | literal value (unverified)                                                              | unverified                                                                                    |
| Generic stdio client | launch `node <DIST>`                                                                                                                                                                    | up to you; the snippet uses `${OUTLOOK_OAUTH_CLIENT_ID}`                                | your own                                                                                      |

Pi and Claude Code prefix tool names to `mcp__outlook__<tool>`. For Codex and VS
Code **no official prefix format is documented**, so never hardcode a prefixed
name.

## Pi

```bash
pi mcp add outlook \
  --env OUTLOOK_OAUTH_CLIENT_ID="$OUTLOOK_OAUTH_CLIENT_ID" \
  --exposure direct \
  --description "Read, search, and draft Outlook mail" \
  -- node <DIST>
pi mcp list
```

**Add the description and `--exposure direct`.** Pi defaults MCP tools to
codemode, where they are not declared to the model and the agent has to find them
indirectly; a server this small is better declared, and the description is what
tells the model why it exists.

Or as a configuration entry, where `${OUTLOOK_OAUTH_CLIENT_ID}` expansion is
documented:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["<DIST>"],
      "env": { "OUTLOOK_OAUTH_CLIENT_ID": "${OUTLOOK_OAUTH_CLIENT_ID}" },
      "description": "Read, search, and draft Outlook mail",
      "exposure": "direct"
    }
  }
}
```

`timeout` counts **seconds** and defaults to 60. Sign-in is already
non-blocking, so that is enough for every call here.

## Claude Code

```bash
claude mcp add outlook \
  --scope user \
  --env OUTLOOK_OAUTH_CLIENT_ID="your-application-client-id" \
  --transport stdio \
  -- node <DIST>
```

**Never put the server name straight after `--env`.** The CLI reads the next
token as another `KEY=value` pair and rejects the name; keep `--transport stdio`
between them.

For a project-scoped `.mcp.json`, `${VAR}` expansion is documented:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["<DIST>"],
      "env": { "OUTLOOK_OAUTH_CLIENT_ID": "${OUTLOOK_OAUTH_CLIENT_ID}" }
    }
  }
}
```

At user scope (`~/.claude.json`) expansion is **not** documented, so `config
claude-code` prints, and `--write` merges, an entry holding the literal client id
the command was started with. `MCP_TIMEOUT` is milliseconds, and it is the
startup budget to raise if you install from a one-liner.

## Claude Desktop

_Unverified in this repository: the file location below was not checked against
Anthropic's documentation, so treat it as a strong convention and confirm it in
the app's own settings._

Add the `outlook` entry under `mcpServers` in `claude_desktop_config.json`
(`~/Library/Application Support/Claude/` on macOS, `%APPDATA%\Claude\` on
Windows, `~/.config/Claude/` on Linux) and restart the app:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["<DIST>"],
      "env": { "OUTLOOK_OAUTH_CLIENT_ID": "your-application-client-id" }
    }
  }
}
```

The app reads this file when it starts. `node dist/index.js config
claude-desktop --write` merges it for you, keeping anything else in the file.

## Codex

```bash
codex mcp add outlook --env OUTLOOK_OAUTH_CLIENT_ID="your-application-client-id" -- node <DIST>
```

That writes the value into the global `~/.codex/config.toml`; there is no
project scope. **Codex documents no interpolation for TOML `env`**, so to keep the
client id out of the file, forward it instead:

```toml
[mcp_servers.outlook]
command = "node"
args = ["<DIST>"]
env_vars = ["OUTLOOK_OAUTH_CLIENT_ID"]
startup_timeout_sec = 30
```

Export `OUTLOOK_OAUTH_CLIENT_ID` before starting Codex, or write the literal
value into an `env` table. `startup_timeout_sec` defaults to **10**, which a cold
one-liner install can exceed, so `config codex --write` sets it to 30.
`tool_timeout_sec` stays at its 60 second default: sign-in never blocks and a
read or a draft finishes well inside it.

## VS Code and Copilot

```bash
code --add-mcp "{\"name\":\"outlook\",\"type\":\"stdio\",\"command\":\"node\",\"args\":[\"<DIST>\"],\"env\":{\"OUTLOOK_OAUTH_CLIENT_ID\":\"your-application-client-id\"}}"
```

One JSON string, no `--` separator and no scope flag: it writes the user profile.
**Always write `"type": "stdio"`.** VS Code's reference table marks `type`
required while its own examples omit it, so emit it to satisfy the stricter
reading. Or edit the `mcp.json` in your user profile (`MCP: Open User Config`) —
_unverified: the absolute path of that file is not published, and the location
`config vscode --write` uses is a convention_:

```json
{
  "servers": {
    "outlook": {
      "type": "stdio",
      "command": "node",
      "args": ["<DIST>"],
      "env": { "OUTLOOK_OAUTH_CLIENT_ID": "your-application-client-id" }
    }
  }
}
```

Note the top-level key is `servers`, not `mcpServers`. `${input:id}` with a
top-level `inputs` array is the documented way to prompt for a value, but
Workspace Trust gates workspace servers and Agent Host **silently skips**
`${input:…}` servers — so this entry asks for nothing and puts the value in
`env`. Whether `${env:NAME}` expands here is unverified, which is why the snippet
does not use it.

## Cursor

_Unverified in this repository: neither the file location nor the entry shape
below was checked against Cursor's documentation._

`~/.cursor/mcp.json` for the whole user, or `.cursor/mcp.json` in a workspace:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["<DIST>"],
      "env": { "OUTLOOK_OAUTH_CLIENT_ID": "your-application-client-id" }
    }
  }
}
```

## Gemini CLI

_Unverified in this repository: the flag names below were not checked against
Gemini CLI's documentation. Read `gemini mcp --help` before trusting them._

```bash
gemini mcp add outlook node <DIST> --scope user --env OUTLOOK_OAUTH_CLIENT_ID="your-application-client-id"
```

User-scope servers live in `~/.gemini/settings.json` under `mcpServers`, which
`node dist/index.js config gemini-cli --write` merges.

## OpenCode

_Unverified in this repository: the key names below were not checked against
OpenCode's documentation._

OpenCode's `mcp` entries carry a local command as a **list**, not as a command
plus args:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "outlook": {
      "type": "local",
      "command": ["node", "<DIST>"],
      "environment": {
        "OUTLOOK_OAUTH_CLIENT_ID": "your-application-client-id"
      },
      "enabled": true
    }
  }
}
```

## Any other client

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["<DIST>"],
      "env": { "OUTLOOK_OAUTH_CLIENT_ID": "${OUTLOOK_OAUTH_CLIENT_ID}" }
    }
  }
}
```

With no arguments the server speaks MCP over stdio; `serve` says so explicitly.
Everything else on stdout is a diagnostic, and `node <entry> --help` is the CLI.
If your harness cannot expand `${VAR}`, put the literal client id there.

## What was verified, and what was not

Verified against each client's own official documentation, and recorded entry by
entry in `src/core/clients.ts`, where the absence of an `unverified` note is the
record: **Pi**, **Claude Code**, **Codex**, and the **VS Code / Copilot** shape —
the add commands, the environment-variable strategy, the units, and the failure
modes in the matrix above.

Also checked against npm's documentation: the git-specifier and tarball
behaviour, `prepare`, `files` versus `.gitignore`, and `engines`.

**Not verified**, and marked as such wherever it appears: Claude Desktop's file
location, Cursor's location and shape, Gemini CLI's flag names, OpenCode's key
names, the absolute path of the VS Code user-profile `mcp.json`, whether
`${env:NAME}` expands in an MCP config, whether `npx` re-clones a git specifier,
and any Codex or VS Code tool-name prefix.
