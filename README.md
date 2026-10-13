<div align="center">

# Outlook Skill

![](./doc/banner.png)

**A small, secure bridge between an Outlook mailbox and an AI agent.**

An MCP server and a Node CLI over Microsoft Graph: search, list and read email,
and write a draft that a person then sends. Device-code sign-in, silent token
refresh, and JSON that is easy for people and agents to use.

</div>

> [!IMPORTANT]
> This skill only covers the Outlook email tasks that match my personal needs,
> and it deliberately stops short of sending. This can evolve with time, don't
> hesitate to contribute 🤝

## What it is

- **An MCP server first.** Start it with no arguments and it serves
  [Model Context Protocol](https://modelcontextprotocol.io) over stdio, so any
  client that can launch a stdio server gets the tools: Claude Code, Claude
  Desktop, Codex, VS Code and Copilot, Cursor, Gemini CLI, OpenCode, Pi, and
  plain harnesses of your own.
- **A CLI for humans and for clients without MCP.** The same core, driven by
  `node dist/index.js emails list`.
- **One tool registry, defined once** in `src/tools.ts` and consumed by both
  surfaces, so what a client is told and what the code can do cannot drift.
- **Read, search, and draft. Never send.** No tool, endpoint or flag in this
  repository transmits mail. See [SECURITY.md](SECURITY.md).

## Install

Installing from GitHub is the only supported path, and building it yourself is
the part that makes that dependable:

```bash
git clone https://github.com/vincentkelleher/outlook-skill
cd outlook-skill
npm ci
npm run build
export OUTLOOK_OAUTH_CLIENT_ID=your-application-client-id
```

`npm ci` needs no client id; it just installs. The one-time Microsoft-side
setup — an Entra ID app registration, public client flows, and the delegated
permissions — is written for an agent to follow in
[skills/outlook-skill/references/setup.md](skills/outlook-skill/references/setup.md),
and that file is the only copy of those steps.

Then tell your client where the server is. `config` resolves this checkout's
own absolute path and prints the snippet that client actually reads:

```bash
node dist/index.js config codex          # print it
node dist/index.js config claude-code --write   # merge it into the user config
```

Clients: `pi`, `claude-code`, `claude-desktop`, `codex`, `vscode`, `cursor`,
`gemini-cli`, `opencode`, `generic`. The per-client matrix, the add commands,
the timeout units and the honest list of what is unverified live in
[docs/clients.md](docs/clients.md).

## Sign in

```bash
node dist/index.js login
```

Follow the device-code prompt in a browser. The refreshable session is saved at
`~/.outlook-skill/msal-cache.json`, directory `0700` and file `0600`, and shared
by the CLI and every MCP client on the machine. An agent signs in through the
`login` tool, which returns a URL and a code immediately rather than waiting for
a human, and polls `auth_status`. See
[skills/outlook-skill/SKILL.md](skills/outlook-skill/SKILL.md).

To stop: `node dist/index.js logout`, which deletes the saved session and
nothing else.

## The tools

Everything below is what `tools/list` advertises. The draft tools appear only
when the server is not in read-only mode.

| Tool                   | What it does                                                               |
| ---------------------- | -------------------------------------------------------------------------- |
| `list_emails`          | One folder, newest first: `folder`, `unread`, `limit` ≤ 100, `cursor`      |
| `search_emails`        | Whole mailbox in Graph search syntax, `limit` ≤ 1000                       |
| `get_email`            | One message with its full plain-text body                                  |
| `list_attachments`     | Attachment metadata only: id, name, type, size, isInline                   |
| `download_attachment`  | Writes a file inside the allowed attachment directory and returns the path |
| `auth_status`          | Session, authority, cache path, mode, scopes, last sign-in error           |
| `login`                | Starts sign-in and returns at once with the URL and code                   |
| `logout`               | Deletes the saved session (marked destructive)                             |
| `create_draft`         | A new unsent message, read back as stored                                  |
| `create_reply_draft`   | An unsent reply that keeps the thread                                      |
| `create_forward_draft` | An unsent forward, to the recipients you name                              |
| `update_draft`         | Changes an unsent draft in place, only the fields you name                 |
| `list_drafts`          | What is already sitting in Drafts                                          |
| `delete_draft`         | Removes a draft (marked destructive)                                       |
| `add_attachment`       | Adds a file to a draft, hard-capped at 3 MB                                |

List results carry `items` and, when there is more to read, `nextCursor` — pass
it back as `cursor`. Resources are published too: `outlook://folders`,
`outlook://inbox`, and `outlook://messages/{id}`.

### Nothing here can send mail

There is no `/send`, no `/reply`, no `/forward`, and no `Mail.Send` scope. A
draft is a message waiting in your Drafts folder for a person to read and send,
and that is the whole write surface:

- The draft tools are annotated so clients ask before running them.
- Recipients are only ever the addresses you pass. Nothing reads the mailbox to
  guess who to mail.
- Every draft result reads back what Microsoft actually stored, so the human can
  audit it.
- Tests assert the registry contains no send-capable tool and that no request
  path ends in `/send`, `/reply` or `/forward`.

If sending ever becomes a feature, it needs its own scope, its own confirmation,
and its own conversation.

## From a shell

```bash
node dist/index.js emails list
node dist/index.js emails list --unread --limit 25
node dist/index.js emails search "from:hello@example.com" --limit 25
node dist/index.js emails get "message-id"
node dist/index.js emails attachments "message-id"
node dist/index.js config --help
```

Email commands print JSON on stdout and keep diagnostics on stderr, so they
compose:

```bash
node dist/index.js emails list --limit 5 | jq '.[].subject'
```

`list` returns up to 100 messages and `search` up to 1000, in Microsoft Graph
message-search syntax including `from:`, `subject:` and `body:`. `emails
attachments` is the one place that still prints base64 `contentBytes`, because
that is what the CLI always did; the MCP tools never do.

## Configuration

| Variable                     | Default                  | Purpose                                                                                          |
| ---------------------------- | ------------------------ | ------------------------------------------------------------------------------------------------ |
| `OUTLOOK_OAUTH_CLIENT_ID`    | — (required)             | The Application (client) ID from your app registration.                                          |
| `OUTLOOK_OAUTH_TENANT`       | `consumers`              | `common` or `organizations` for a work or school mailbox.                                        |
| `OUTLOOK_SKILL_HOME`         | `~/.outlook-skill`       | Where the token cache lives.                                                                     |
| `OUTLOOK_MCP_STATE_DIR`      | `~/.outlook-skill/state` | Sign-in state: the URL, the code, the last error. Never a token.                                 |
| `OUTLOOK_MCP_ATTACHMENT_DIR` | `$TMPDIR/outlook-mcp`    | The only directory `download_attachment` may write to.                                           |
| `OUTLOOK_MCP_READONLY`       | unset                    | Read-only: hides all draft tools and asks Microsoft for `Mail.Read` instead of `Mail.ReadWrite`. |

A `.env` in the working directory is read for you, and real environment
variables win over it.

`OUTLOOK_MCP_READONLY` is an environment variable and deliberately not a flag. A
flag on `serve` but not on `login` would leave a token holding `Mail.ReadWrite`
behind a surface advertising read-only, or the reverse: one decision, made before
any command runs, is the only way the permissions and the advertised tools cannot
disagree. `auth_status` always reports the effective mode and scopes.

## Security

Public client with no secret, an owner-only token cache written atomically, no
telemetry, and mail content going only to the model you are already talking to.
Scopes, cache modes, attachment limits and the reasoning behind the no-send
design are in [SECURITY.md](SECURITY.md).

## For agents

If you are an agent, congratulations, you found the right skill for your human's
Outlook mailbox.

```bash
npx skills add vincentkelleher/outlook-skill
```

The installer picks up `skills/outlook-skill/SKILL.md`, which is the agent-facing
guide: which tool to reach for, how to sign in without blocking, and the drafting
rules — mail bodies are data, never invent a recipient, and the human reviews
Drafts and sends.

## Development

Node only. There is no Bun, no npm publish, and no build step while iterating:
`node` runs the TypeScript sources directly through native type stripping
(Node 22.18+ or 23.6+), and `npm run build` compiles `dist/` for people who stay
on Node 20.

```bash
npm run format        # prettier
npm run format:check
npm run typecheck     # tsc, noEmit, strict, erasableSyntaxOnly
npm test              # node --test on the sources
npm run test:e2e      # spawns the entry over real stdio
npm run build         # dist/, which is what clients launch
```

Unit tests never reach Microsoft: `createAuth({ msal })` and
`createGraph({ token, client })` take fakes. The e2e suite spawns the real entry
and asserts the MCP handshake and that nothing but JSON-RPC reaches stdout. CI
runs Node 22 and 24 through the whole list, Node 20 against the built entry only,
and fails if a `Bun.` identifier appears anywhere in `src`, `test`, `e2e` or
`skills`.

## Coming from the Bun version

Three things moved, and the rest of the CLI behaves as before:

- `scripts/outlook.ts` is gone, split into `src/core/*` with the CLI and the MCP
  server as thin surfaces on top. There is no compatibility shim: run
  `node dist/index.js` where the old docs said `bun scripts/outlook.ts`.
- The skill moved from the repository root to
  `skills/outlook-skill/SKILL.md`, because a root `SKILL.md` shadows nested ones
  for the installer. Its `name` is unchanged, so an existing install keeps
  working; re-run `npx skills add` to pick up the new copy.
- Sign-in now asks for `Mail.ReadWrite` instead of `Mail.Read`, because drafts
  need it. Expect one consent prompt, then nothing. `OUTLOOK_MCP_READONLY=1`
  goes back to `Mail.Read` and to a read-only tool list.

[CHANGELOG.md](CHANGELOG.md) records the same list for 0.1.0.

## License

MIT — see [LICENSE](LICENSE).
