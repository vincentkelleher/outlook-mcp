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

The only network destinations are the Microsoft identity endpoint and Microsoft
Graph: this code has no analytics, no update check, and no third-party call.
Mail content reaches the model you are already talking to and nowhere else.

## Scopes, and the one it will never ask for

Requested at sign-in (`src/core/config.ts`, `scopesFor`):

| Scope              | Why                                             |
| ------------------ | ----------------------------------------------- |
| `User.Read`        | Identifying the signed-in account               |
| `Mail.ReadWrite`   | Reading mail **and** creating or editing drafts |
| `email`, `profile` | Standard OpenID scopes                          |

With `OUTLOOK_MCP_READONLY=1` the second becomes `Mail.Read` and no draft tool is
advertised at all, so the permissions and the tool surface cannot disagree.

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

The write surface is exactly six draft operations plus one attachment upload:
`create_draft`, `create_reply_draft`, `create_forward_draft`, `update_draft`,
`delete_draft`, `add_attachment`, and `list_drafts` to read them. Guardrails
around it:

- Every draft tool is annotated `readOnlyHint: false`, `idempotentHint: false`,
  so conforming clients ask the human first. `logout` and `delete_draft` are
  marked `destructiveHint: true`.
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
  `skills/outlook-skill/SKILL.md`.

## Where it writes, and with which permissions

| Path                                       | Modes                   | Content                                                           |
| ------------------------------------------ | ----------------------- | ----------------------------------------------------------------- |
| `~/.outlook-skill/` (`OUTLOOK_SKILL_HOME`) | `0700`                  | the cache directory                                               |
| `~/.outlook-skill/msal-cache.json`         | `0600`                  | the refreshable token cache                                       |
| `~/.outlook-skill/state/login.json`        | dir `0700`, file `0600` | sign-in state: URL, user code, expiry, last error. Never a token. |
| `$TMPDIR/outlook-mcp`                      | `0700`                  | downloaded attachments                                            |

The cache is the sensitive file and it is written atomically: a temp file in the
same directory at `0600`, then a `rename`. Several agents can share one cache,
and a non-atomic write there can truncate it mid-refresh. A temp file is a
plaintext copy of the cache, so a process killed mid-write leaves one behind: the
next cache read sweeps temp files older than an hour, since nothing can still be
writing at that age. On a silent-refresh failure the cache is re-read and
refreshed once more before anyone is asked to sign in again.

`~/.outlook-skill/` is shared by the CLI and every MCP client on the machine, so
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
JSON-RPC and sends diagnostics to stderr, and nothing logs mail content.

## Reporting a problem

Open a [private security advisory](https://github.com/vincentkelleher/outlook-skill/security/advisories/new)
rather than a public issue. This is a personal project with no formal SLA; please
give it a few days before disclosing publicly.
