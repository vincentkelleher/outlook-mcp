# When something is wrong

Keyed on the sentence you are looking at, because that is what you have. The
message is the remediation: this server says what failed and what to run, so
match it here rather than guessing.

To see more, set `OUTLOOK_LOG_LEVEL=debug` on the failing process and read its
stderr — one line per step, the Graph call, its verb, path, duration and item
count. `OUTLOOK_LOG_FORMAT=json` is for a log collector. Nothing that looks like
mail ever travels in a line: no token, no body, no attachment content, at any
level.

## Sign-in and session

| What you see                                       | What it means                                                                                                                                         | What to do                                                                                                                                     |
| -------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| `No Outlook session. Run the login command first.` | Nothing is cached on this machine.                                                                                                                    | Start a sign-in (`login` tool, or `node dist/index.js login --background --json`), show the URL and code, poll `auth status` until `signedIn`. |
| `Outlook session expired (AADSTS700082…)`          | The refresh token is dead. `reason` is `expired`.                                                                                                     | Sign in again. There is no other fix, and no need to touch the app registration.                                                               |
| `Outlook session needs re-authorisation …`         | The cached grant is narrower than what this process asks for — the app gained `Mail.ReadWrite`, or read-only was flipped. `reason` is `needsConsent`. | Sign in again to re-consent. Check the permission list in [setup.md](setup.md) first, or the app cannot grant what is asked.                   |
| `Could not confirm the Outlook session …`          | The check itself failed: network, DNS, a Microsoft incident. `reason` is `unavailable` and it says so outright.                                       | Retry later. **Do not start a device-code flow for this** — it needs the same network, and Microsoft rate-limits these.                        |
| `A sign-in is already running as process N …`      | Another client holds the live flow. Two codes are worse than none: the user gets one to type.                                                         | Poll `auth status` (or `auth_status`) — it carries the live `verificationUri` and `userCode`. `node dist/index.js logout` clears the state.    |
| `AADSTS700016`, `unauthorized_client` at sign-in   | Not a public client, or the client id is wrong.                                                                                                       | **Authentication → Advanced settings → Allow public client flows → Yes**, then check `OUTLOOK_OAUTH_CLIENT_ID` is the Application (client) ID. |
| `AADSTS65001`, `interaction_required`              | Delegated permissions are not consented. In a work or school tenant, user consent is often blocked.                                                   | An **admin** has to grant the API permissions on the registration. Nothing in this repository can do it.                                       |
| The wrong mailbox answers                          | A different account is cached.                                                                                                                        | `node dist/index.js logout`, sign in as the account you want.                                                                                  |
| Drafts and flags fail although `signedIn` is true  | The registration carries `Mail.Read` only.                                                                                                            | Add `Mail.ReadWrite` to the app, then sign in again — see “Switching an existing registration” in [setup.md](setup.md).                        |
| Draft tools absent from `tools/list`               | `OUTLOOK_MCP_READONLY` is set on this process. That is one decision, made before any command runs.                                                    | `auth status` reports `readonly` and the `scopes` behind it. Clear the variable and restart the server; you cannot switch it per command.      |

## Client configuration

| What you see                                                         | What it means                                                                                           | What to do                                                                                                                                     |
| -------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| `Cannot find module …/dist/index.js`, or no `dist/` after an install | `prepare` did not run — npm v12 flips `allowScripts` off and `--allow-git` to `none` by default.        | `cd <checkout> && npm ci && npm run build`. Prefer the clone-and-build route; see `docs/clients.md`.                                           |
| The client says the server timed out on startup                      | A cold clone-and-install outlasts the client's budget: Codex waits 10 s, Claude Code has `MCP_TIMEOUT`. | Build once, then register the built path. Raise `startup_timeout_sec` / `MCP_TIMEOUT` — `config codex --write` sets 30.                        |
| `config --write` refused                                             | The client's file exists and does not parse, or no location for that client is stated with confidence.  | Fix the file, or paste the printed snippet. It never overwrites something it cannot parse.                                                     |
| `OUTLOOK_OAUTH_CLIENT_ID` missing from the client's environment      | The client does not expand `${VAR}` (Codex documents no interpolation in TOML `env`).                   | Write the literal id, or forward it — `env_vars = ["OUTLOOK_OAUTH_CLIENT_ID"]` for Codex. `config <client>` prints the form that client reads. |

## HTTP and the container

| What you see                                     | What it means                                                                                                                 | What to do                                                                                                                               |
| ------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `404`                                            | Wrong path. The MCP endpoint is `/mcp`; `/health` answers with the version, the mode and a session count.                     | Point the client at `http://127.0.0.1:3939/mcp`. `config docker` prints that entry.                                                      |
| A refusal naming `OUTLOOK_MCP_ALLOWED_HOSTS`     | The `Host` header is a name the server does not recognise — DNS rebinding defence, not a bug.                                 | Set `OUTLOOK_MCP_ALLOWED_HOSTS` on the server to the name you reach it by. Loopback always passes.                                       |
| A refusal naming `OUTLOOK_MCP_ALLOWED_ORIGINS`   | A browser sent an `Origin` that is not the origin the request was addressed as.                                               | Use a native MCP client, or name that origin. This guard exists so a page you merely visited cannot read your mail through a local port. |
| Asked to sign in every time the container starts | The token cache lives in `/data`, and `/data` was not a volume.                                                               | `-v outlook-mcp-home:/data`. Two containers sharing the volume share one session.                                                        |
| `docker ps` reports unhealthy                    | The health check calls `/health`, which only a `serve --http` container answers. A container running `login` has no listener. | Only true of it. Check `docker logs outlook-mcp`.                                                                                        |
| Plain `-p 3939:3939`                             | The port is published to your network. `Host` and `Origin` still refuse strangers, but that is not a wall.                    | Use `-p 127.0.0.1:3939:3939`. Nothing on that port authenticates anybody.                                                                |

## Attachments and drafts

| What you see                                             | What it means                                                                                                                                                     | What to do                                                                                                                      |
| -------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| A `resource_link` that `resources/read` will not resolve | Handles belong to the session that issued them and die with it.                                                                                                   | Call `download_attachment` again in the current session.                                                                        |
| An inline attachment refused as too big                  | `encoding: "base64"` is capped at 64 KiB and never falls back to a path silently.                                                                                 | Drop `encoding` and read the resource link, or use `delivery: "server_file"` when the server's filesystem is really reachable.  |
| A download refused over the limit                        | Decoded size defaults to 20 MiB (`OUTLOOK_MCP_MAX_ATTACHMENT_BYTES`).                                                                                             | Raise it on the server, or don't fetch that file. Metadata from `list_attachments` tells you the size first.                    |
| An upload refused over the ceiling                       | `OUTLOOK_MCP_MAX_UPLOAD_BYTES`, 35 MiB by default, 150 MiB at most — Graph's own ceiling. The default is the mailbox's, so a file that uploads can still be sent. | Raise it to what the mailbox takes, or compress. A refused upload costs nothing: the draft is unchanged.                        |
| `… is not a draft`                                       | The target message was sent. Graph would add or remove a file on a sent message; this server refuses.                                                             | Nothing to fix. Sent mail keeps the files it went out with.                                                                     |
| No `webLink` worth opening                               | It is unreliable for personal accounts.                                                                                                                           | Tell the person the draft is in Outlook's Drafts folder and read back the stored draft, which every draft tool already returns. |

## The CLI guards that cost nothing

- `--limit must be an integer from 1 to N` — nothing was sent. `list` and
  `search` cap at 100 and 1000; `emails read`/`unread` and the flag tools at 100
  ids.
- `At most 100 ids per command` — split the batch. Nothing was sent.
- `--port must be an integer from 0 to 65535` — `0` is allowed by the flag, not
  by `OUTLOOK_MCP_PORT`, because a service has to name the port people reach.

If none of this fits, run the failing command once with
`OUTLOOK_LOG_LEVEL=debug`, find the last line before the failure, and quote that
line — it names the step, and usually Microsoft's own reason.
