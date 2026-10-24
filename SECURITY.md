# Security notes

What this software touches, what it can do with it, and where it writes. Written
for someone deciding whether to run it against a real mailbox.

## It is a public client with no secret

The Microsoft Entra application is a **public client** using the device-code
flow: no client secret, no certificate, no redirect URI, nothing that can be
leaked from this repository, your shell history, or a config file. The
`OUTLOOK_OAUTH_CLIENT_ID` is an identifier, not a credential, and the app
registration must have **Allow public client flows** enabled for sign-in to work
at all.

The only outbound API destinations are the Microsoft identity endpoint and
Microsoft Graph: this code has no analytics, no update check, and no third-party
API call. MCP returns mail content to the connected client.

## Local MCP transports

Stdio is the default and opens no listening socket. `serve --http` explicitly
enables Streamable HTTP at `http://127.0.0.1:3939/mcp` (`OUTLOOK_MCP_HOST`,
`OUTLOOK_MCP_PORT`, `--host`, `--port`, with `--port 0` choosing a free port).
Every request is admitted on its `Host` and `Origin` before anything is read —
loopback by default, with other names accepted through
`OUTLOOK_MCP_ALLOWED_HOSTS` — protecting against DNS rebinding and browser
requests from non-local origins. Unknown paths return 404.

**HTTP clients are not authenticated.** These guards are not authorization:
local processes and allowed localhost browser origins can use the cached mailbox
session, including its writes — flagging mail and drafts — unless read-only mode
is enabled. Only run it
on a trusted single-user machine. Do not expose it via proxies, tunnels or port
forwarding; remote hosting needs verified client authentication and TLS first.
Each session gets its own MCP server instance — at most 32 open at a time, and
idle ones are retired after `OUTLOOK_MCP_SESSION_IDLE_MINUTES` — but every
session shares the configured Outlook account, cache and sign-in state.
SIGINT/SIGTERM close the sessions and the listener.

## Scopes, and the one it will never ask for

Requested at sign-in (`src/core/config.ts`, `scopesFor`):

| Scope              | Why                                                 |
| ------------------ | --------------------------------------------------- |
| `User.Read`        | Identifying the signed-in account                   |
| `Mail.ReadWrite`   | Reading mail **and** editing it: drafts, read flags |
| `email`, `profile` | Standard OpenID scopes                              |

With `OUTLOOK_MCP_READONLY=1` the second becomes `Mail.Read` and no mailbox write
is advertised at all — neither `mark_read`/`mark_unread` nor a draft tool — so the
permissions and the tool surface cannot disagree.

**`Mail.Send` is never requested, and never will be by this design.** Creating a
draft is a write to your mailbox, not a transmission. See below.

The default authority is `consumers` (personal accounts);
`OUTLOOK_OAUTH_TENANT` switches it to `common` or `organizations`.

## Read plus draft, never send

No tool, code path, or flag can transmit mail:

- There is no `/send`, `/reply` or `/forward` call anywhere in `src/`.
- The tool registry contains nothing that offers to send.
- Both claims are asserted by tests (`test/no-send.test.ts`), including a check
  that the scan itself distinguishes Graph's `createReply`/`createForward` draft
  actions from the sending ones.

The write surface is six draft operations plus one attachment upload —
`create_draft`, `create_reply_draft`, `create_forward_draft`, `update_draft`,
`delete_draft`, `add_attachment`, with `list_drafts` to read them — and two flag
operations, `mark_read` and `mark_unread`, which send one `PATCH` of `isRead` per
id and touch no message content. Guardrails around it:

- Every draft tool is annotated `readOnlyHint: false`, `idempotentHint: false`,
  so conforming clients ask the human first. `logout` and `delete_draft` are
  marked `destructiveHint: true`. The flag tools are `readOnlyHint: false` too,
  with `idempotentHint: true`: flagging a message that is already read changes
  nothing, which is what makes a retry safe.
- **Recipients are never inferred.** Tools accept explicit addresses and no tool
  reads the mailbox to guess one.
- **Draft readback.** Every write returns what Microsoft actually stored —
  subject, recipients, body, format — so a person can audit it instead of
  trusting what was asked for. Body readback is capped at 4000 characters and
  the cap is stated in the result.
- A draft result says `sent: false` and points at the Drafts folder.
- Mail content is untrusted data. A message body is never an instruction, and an
  email asking for a reply is not permission to draft one. That is stated in the
  server `instructions`, in the tool descriptions, and in
  `skills/outlook-mcp/SKILL.md`.

## Where it writes, and with which permissions

| Path                                   | Modes                   | Content                                                           |
| -------------------------------------- | ----------------------- | ----------------------------------------------------------------- |
| `~/.outlook-mcp/` (`OUTLOOK_MCP_HOME`) | `0700`                  | the cache directory                                               |
| `~/.outlook-mcp/msal-cache.json`       | `0600`                  | the refreshable token cache                                       |
| `~/.outlook-mcp/state/login.json`      | dir `0700`, file `0600` | sign-in state: URL, user code, expiry, last error. Never a token. |
| `$TMPDIR/outlook-mcp`                  | `0700`                  | downloaded attachments                                            |

The cache is the sensitive file and it is written atomically: a temp file in the
same directory at `0600`, then a `rename`. Several agents can share one cache,
and a non-atomic write there can truncate it mid-refresh. A temp file is a
plaintext copy of the cache, so a process killed mid-write leaves one behind: the
next cache read sweeps temp files older than an hour, since nothing can still be
writing at that age. On a silent-refresh failure the cache is re-read and
refreshed once more before anyone is asked to sign in again.

`~/.outlook-mcp/` is shared by the CLI and every MCP client on the machine, so
one sign-in is enough — and anyone with write access to that directory is
effectively signed in as you. Keep it on your own account. The sign-in state file
carries a short-lived redeemable code rather than a token, so both it and its
directory are owner-only, and a finished sign-in rewrites it to a bare status, so
the code is on disk only while a flow is live.

## Attachment limits

Two directions, both bounded:

- **Download** (`download_attachment`): writes only inside
  `OUTLOOK_MCP_ATTACHMENT_DIR` (default `$TMPDIR/outlook-mcp`), mode `0600`. A
  requested `dir` is resolved inside that root and checked twice — as a path and
  as a real path — so neither `../` nor a symlink escapes it, and a Graph
  filename is reduced to its final segment before it reaches the disk. Base64
  content is returned inline only when explicitly requested and only up to 64
  KiB; above that the result says so and gives the file path instead, which keeps
  a large attachment out of a model's context by default.
- **Upload** (`add_attachment`): a **hard 3 MB cap** on the decoded file, checked
  before any request, with an actionable error and no partial write. Graph needs
  an upload session between 3 and 150 MB and this server has none, so the limit
  is the limit rather than a silent truncation.

## Session handling

Sign-in is a detached background process because the device-code flow can wait
minutes for a person; it writes the verification URL and user code to the state
file and returns. Codes are short-lived and are not tokens. `logout` deletes the
token cache and that state file together, so a failure recorded by an old session
cannot be read as a fresh one by the next. The MCP server keeps stdout clean for
JSON-RPC and sends diagnostics to stderr.

## What the logs say, and what they never say

Every command, tool call, sign-in and HTTP session writes one line to stderr —
never to stdout, which is a protocol on stdio and a `jq` input in a shell.
`OUTLOOK_LOG_LEVEL` is `debug`, `info` (the default), `warn`, `error` or
`silent`; `OUTLOOK_LOG_FORMAT` chooses a line for a terminal or one JSON record
per line for a collector.

A log line is a copy of a mailbox, and stderr is not owner-only the way
`~/.outlook-mcp/msal-cache.json` is: a container log, a CI artefact and a
pasted terminal are all places mail can end up without anyone deciding that. So:

- **Never written, at any level:** an access or refresh token, an `Authorization`
  header, a message body, or attachment content. Those keys are struck out of a
  record before it is serialised, and the human format is rendered from that same
  serialised record — so there is no path that bypasses it. A test asserts it.
- **`debug` only, because somebody asked for it:** message and attachment ids, a
  Microsoft Graph path, a search query, a recipient list, a sender address in
  full, a saved attachment's path, and an error stack.
- **The default level:** what happened and how much of it — the tool, the folder,
  a count of items, bytes, milliseconds, a session id, a masked address
  (`v***@example.com`), an id shortened to eight characters, a query as its
  length.

The sign-in state file carries the verification URL and user code, which it must
for a polling client to find them; a log line masks that code until `debug`, and
no state file or log line carries a token.

## The HTTP port, and what it costs

`serve --http` — which is what the Docker image runs — puts the same server behind
a listening socket. That is the one part of this project where something other
than a process you started can reach a mailbox, so it is bounded deliberately:

- **It authenticates nobody.** Anything that can open the socket can read the
  mailbox and write drafts, exactly as the stdio server can. It is a local
  transport, not a hosted one: no tokens, no users, no TLS, and nothing in this
  repository adds them. Do not put it behind a proxy, and do not publish it.
- **The default bind is loopback**, `127.0.0.1:3939`, and the container's wildcard
  bind is answered on the host side by `-p 127.0.0.1:3939:3939`, which keeps the
  port off the network. Plain `-p 3939:3939` puts it on your LAN.
- **Every request is admitted by `Host` and `Origin`.** The `Host` must be a
  loopback address or a name in `OUTLOOK_MCP_ALLOWED_HOSTS`; a browser `Origin`
  must be the same origin the request was addressed as, or one in
  `OUTLOOK_MCP_ALLOWED_ORIGINS`. That is the DNS-rebinding defence: without it, a
  page you merely visited could re-resolve `localhost` to your own machine and
  read your mail through a port you left open for your agent. A request with no
  `Origin` is a native MCP client, and is served on the strength of its `Host`
  alone.
- **The response body says nothing useful to a stranger** — a refusal names the
  variable to set, nothing about the mailbox — and `/health` reports only the
  version, the mode and a session count, never mail and never an account.
- **Bodies are capped at 4 MiB** and sessions at 32, with an idle session retired
  after `OUTLOOK_MCP_SESSION_IDLE_MINUTES` (default a day) once nothing is
  connected to it, so a client that disconnects without saying so cannot pile up
  server instances for the life of the container.

What running it in a container does change is where the secrets live: the token
cache moves to `/data`, so the volume holds a refreshable session. Give that
volume to no other container, do not bake `.env` into an image (the entry reads
one from its working directory, and `docker history` keeps layers), and remember
that a container running as root leaves files owned by root in it.

Stdio remains the default transport, and the default for a reason: a process a
client launches, on pipes, with no socket at all, is a smaller surface than a
port. If you do not need the container, do not run it.

## Reporting a problem

Open a [private security advisory](https://github.com/vincentkelleher/outlook-mcp/security/advisories/new)
rather than a public issue. This is a personal project with no formal SLA; please
give it a few days before disclosing publicly.
