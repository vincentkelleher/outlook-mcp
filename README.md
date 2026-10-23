<div align="center">

# Outlook MCP

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
- **Or in a container, over HTTP.** `serve --http` speaks the same protocol over
  a Streamable HTTP endpoint, which is what the published image runs: one
  long-lived container, a URL any client can reach. See
  [Run it in a container](#run-it-in-a-container).
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
git clone https://github.com/vincentkelleher/outlook-mcp
cd outlook-mcp
npm ci
npm run build
export OUTLOOK_OAUTH_CLIENT_ID=your-application-client-id
```

`npm ci` needs no client id; it just installs. The one-time Microsoft-side
setup — an Entra ID app registration, public client flows, and the delegated
permissions — is written for an agent to follow in
[skills/outlook-mcp/references/setup.md](skills/outlook-mcp/references/setup.md),
and that file is the only copy of those steps.

Then tell your client where the server is. `config` resolves this checkout's
own absolute path and prints the snippet that client actually reads:

```bash
node dist/index.js config codex          # print it
node dist/index.js config claude-code --write   # merge it into the user config
```

Clients: `pi`, `claude-code`, `claude-desktop`, `codex`, `vscode`, `cursor`,
`gemini-cli`, `opencode`, `docker`, `generic`. The per-client matrix, the add commands,
the timeout units and the honest list of what is unverified live in
[docs/clients.md](docs/clients.md).

## Run it in a container

The image runs the server over Streamable HTTP on port 3939 and keeps nothing on
its own filesystem: the session, the sign-in state and downloaded attachments all
live in `/data`, which is a volume.

```bash
docker run -d --name outlook-mcp \
  -p 127.0.0.1:3939:3939 \
  -v outlook-mcp-home:/data \
  -e OUTLOOK_OAUTH_CLIENT_ID=your-application-client-id \
  ghcr.io/vincentkelleher/outlook-mcp:0.1.0

docker logs outlook-mcp
# Outlook MCP serving on http://127.0.0.1:3939/mcp
```

Then sign in through it. The device-code flow needs no port of its own, so a
browser on the host is enough and nothing has to be exposed for it:

```bash
docker exec outlook-mcp outlook login
# To sign in: https://microsoft.com/devicelogin  Enter the code ABC-XYZ123
```

And point a client at the URL — `node dist/index.js config docker` (or
`docker exec outlook-mcp outlook config docker`) prints it:

```json
{
  "mcpServers": {
    "outlook": { "type": "http", "url": "http://127.0.0.1:3939/mcp" }
  }
}
```

Four things about that command line are important:

- **`-p 127.0.0.1:3939:3939`, not `-p 3939:3939`.** The server inside the
  container binds the container's wildcard interface, because a loopback bind
  there is unreachable through a published port. The host side of the mapping is
  where the port stays local. Plain `-p` publishes it to your network, where
  `Host` and `Origin` checks will still refuse strangers — but refuse them from
  a port you did not mean to open.
- **`-v outlook-mcp-home:/data`, or you sign in again every time.** The token
  cache lives there, with the sign-in state. Two containers sharing a volume
  share one session, which is how `login` and `serve` find each other.
- **`OUTLOOK_OAUTH_CLIENT_ID` goes to the container.** The client entry carries a
  URL and no credentials; the app registration is the container's business.
- **Nothing on that port authenticates.** It is meant for your own machine.
  `GET /health` answers without touching Microsoft, so `docker ps` can report on
  it; `docker stop` takes it down in well under a second.

Build it yourself with `docker build -t outlook-mcp .` — the same Dockerfile CI
publishes from. To use the image over stdio instead, drop `--http`:
`docker run -i --rm -v outlook-mcp-home:/data -e OUTLOOK_OAUTH_CLIENT_ID …
outlook-mcp serve`, which is the shape of it for a client that launches stdio
servers and will not take a URL.

## Sign in

```bash
node dist/index.js login
```

Follow the device-code prompt in a browser. The refreshable session is saved at
`~/.outlook-mcp/msal-cache.json`, directory `0700` and file `0600`, and shared
by the CLI and every MCP client on the machine. An agent signs in through the
`login` tool, which returns a URL and a code immediately rather than waiting for
a human, and polls `auth_status`. See
[skills/outlook-mcp/SKILL.md](skills/outlook-mcp/SKILL.md).

To stop: `node dist/index.js logout`, which deletes the saved session and
nothing else.

Coming from the name this project had before: the directory and its variable were
`.outlook-skill` and `OUTLOOK_SKILL_HOME`, and nothing reads that location any
more. `mv ~/.outlook-skill ~/.outlook-mcp` keeps the session; without it you sign
in again.

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

| Variable                           | Default                  | Purpose                                                                                          |
| ---------------------------------- | ------------------------ | ------------------------------------------------------------------------------------------------ |
| `OUTLOOK_OAUTH_CLIENT_ID`          | — (required)             | The Application (client) ID from your app registration.                                          |
| `OUTLOOK_OAUTH_TENANT`             | `consumers`              | `common` or `organizations` for a work or school mailbox.                                        |
| `OUTLOOK_MCP_HOME`                 | `~/.outlook-mcp`         | Where the token cache lives.                                                                     |
| `OUTLOOK_MCP_STATE_DIR`            | `~/.outlook-mcp/state`   | Sign-in state: the URL, the code, the last error. Never a token.                                 |
| `OUTLOOK_MCP_ATTACHMENT_DIR`       | `$TMPDIR/outlook-mcp`    | The only directory `download_attachment` may write to.                                           |
| `OUTLOOK_MCP_READONLY`             | unset                    | Read-only: hides all draft tools and asks Microsoft for `Mail.Read` instead of `Mail.ReadWrite`. |
| `OUTLOOK_MCP_HOST`                 | `127.0.0.1`              | `serve --http` only: the interface to bind. The image sets `0.0.0.0`; see the note below.        |
| `OUTLOOK_MCP_PORT`                 | `3939`                   | `serve --http` only: the port to bind. `--port` overrides it, and accepts `0` for a free one.    |
| `OUTLOOK_MCP_ALLOWED_HOSTS`        | unset                    | Extra `Host` names to serve, comma-separated, for a server reached by a name of your own.        |
| `OUTLOOK_MCP_ALLOWED_ORIGINS`      | unset                    | Extra browser origins, comma-separated, beyond the same origin.                                  |
| `OUTLOOK_MCP_SESSION_IDLE_MINUTES` | `1440`                   | `serve --http` only: how long an HTTP session with nothing connected survives.                   |
| `OUTLOOK_MCP_IMAGE`                | the published GHCR image | Which image `config docker` names, if you mirror it elsewhere.                                   |
| `OUTLOOK_LOG_LEVEL`                | `info`                   | `debug`, `info`, `warn`, `error` or `silent`. Every log line goes to stderr.                     |
| `OUTLOOK_LOG_FORMAT`               | `text`                   | `text` for a terminal, `json` for a log collector.                                               |

A `.env` in the working directory is read for you, and real environment
variables win over it.

The four HTTP variables that decide who may connect — `OUTLOOK_MCP_HOST`,
`OUTLOOK_MCP_ALLOWED_HOSTS`, `OUTLOOK_MCP_ALLOWED_ORIGINS` and the bind itself —
are checked on every request, because the transport authenticates nobody. The
`Host` must be a loopback address or one you named, and a browser `Origin` must
be the origin the request was addressed as; a request carrying neither header is
a native MCP client, which is the ordinary case. Without that rule, a page any
browser visits could re-resolve `localhost` and read your mail through a port you
left open, which is what [SECURITY.md](SECURITY.md) is about. So: a container on
`0.0.0.0` behind `-p 127.0.0.1:3939:3939` is local, and a name that is not the
machine's own is refused from either side.

`OUTLOOK_MCP_READONLY` is an environment variable and deliberately not a flag. A
flag on `serve` but not on `login` would leave a token holding `Mail.ReadWrite`
behind a surface advertising read-only, or the reverse: one decision, made before
any command runs, is the only way the permissions and the advertised tools cannot
disagree. `auth_status` always reports the effective mode and scopes.

### Reading the logs

Every command, tool call, sign-in and session writes a line to stderr — never to
stdout, which is JSON-RPC on stdio and JSON for `jq` in a shell:

```bash
OUTLOOK_LOG_LEVEL=debug node dist/index.js emails search "from:ana" --limit 5
node dist/index.js serve --http 2> ~/outlook-mcp.log     # info by default
tail -f ~/outlook-mcp.log
```

A line says when, how loud, which module, what happened, and the scalars that
explain it. This is an MCP session over HTTP at the default level:

```text
2026-01-14T09:12:04.334Z info 41802 http: session opened sessionId=6f1c0b0e-2d59-4a1e-9a4f-3b7f0c2d1e55 sessions=1 readonly=false
2026-01-14T09:12:04.356Z info 41802 mcp: tool call sessionId=6f1c0b0e-2d59-4a1e-9a4f-3b7f0c2d1e55 tool=search_emails queryChars=31 limit=5
2026-01-14T09:12:04.357Z info 41802 auth: session check confirmed a session account=v***@example.com ms=188
2026-01-14T09:12:05.087Z info 41802 mcp: tool call completed sessionId=6f1c0b0e-2d59-4a1e-9a4f-3b7f0c2d1e55 tool=search_emails ms=731 bytes=1642 items=5
```

and the same call at `debug`, where the Microsoft Graph request is visible:

```text
2026-01-14T09:12:04.361Z debug 41802 graph: graph request method=GET path=/me/messages ms=724 items=5
```

A sign-in that expired explains itself, and the recovery after it is two lines
that read as one story:

```text
2026-01-14T09:12:41.203Z warn 41802 auth: silent token refresh failed account=v***@example.com problem=expired ms=412 reason="AADSTS700082: ..."
2026-01-14T09:12:41.688Z info 41802 login: Microsoft issued a device code verificationUri=https://microsoft.com/devicelogin userCode=HKDF*** expiresIn=900
```

- **`info`** (the default): one line per command or tool call, per sign-in, per
  HTTP session, plus every warning and error.
- **`debug`**: adds the Microsoft Graph calls with their verb, path, duration and
  item counts, the token-cache reads and writes, and the message ids, file paths,
  search queries and full addresses a call named. That is the level for
  reproducing one call.
- **`warn`** and **`error`**: only what needs reading. A missing session is a
  warning, because a tool call fixes it; anything else is an error, because
  nobody has read it yet.
- **`silent`**: nothing.

Mail itself never travels in a log line at any level: no body, no attachment
content, no token, ever — those keys are struck out before a line is written.
Sender addresses are masked, a message id is shortened to eight characters and a
search query appears as its length until you ask for `debug`, because stderr is
not owner-only: a container log, a CI artefact and a pasted terminal are all
places mail can end up without anyone deciding that.

`OUTLOOK_LOG_FORMAT=json` writes one pino record per line for a collector
(`level`, `time`, `pid`, `scope`, `msg`, then the fields); the same redaction
applies, because the text format is rendered from that record.

## Security

Public client with no secret, an owner-only token cache written atomically, no
telemetry, and mail content going only to the model you are already talking to.
Scopes, cache modes, attachment limits and the reasoning behind the no-send
design are in [SECURITY.md](SECURITY.md).

## For agents

If you are an agent, congratulations, you found the right skill for your human's
Outlook mailbox.

```bash
npx skills add vincentkelleher/outlook-mcp
```

The installer picks up `skills/outlook-mcp/SKILL.md`, which is the agent-facing
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
runs Node 22 and 24 through the whole list, and packs and inspects the release
tarball.

## License

MIT — see [LICENSE](LICENSE).
