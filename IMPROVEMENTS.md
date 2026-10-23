# IMPROVEMENTS

Findings from the review of **2026-10-23**, against `feat/outlook-mcp` at `3478aba`
**plus the uncommitted working tree** (the Streamable HTTP transport: `src/core/http.ts`,
`serve --http`, the `docker` client entry, `Dockerfile`, `.github/workflows/docker.yml`).

Gates at review time: `tsc -p tsconfig.json` clean · `npm test` 120/120 ·
`npm run test:e2e` 12/12 · Node 25 locally, 22 + 24 in CI.

**Verification legend**

- **reproduced** — re-derived by running code or reading the exact line; treat as fact.
- **traced** — reachable path read from source, not executed.
- **single-source** — one reviewer's read, not independently confirmed.

Suggested commit grouping is at the bottom. Items are ordered by what to fix first.

---

## 1. Fix before merge

### I-1 — README claims Host/Origin checks protect a published port. They do not. · HIGH · reproduced

> `README.md:112` — "Plain `-p` publishes it to your network, where `Host` and `Origin`
> checks will still refuse strangers"

`admitRequest` (`src/core/http.ts:96-140`) admits on the `Host` **header string**
(`isLoopbackHost`, `http.ts:16`). Nothing ever compares the socket's local or remote
address. A non-browser client on the network sets its own `Host`.

```
input: { host: "localhost:3939", allowedHosts: [], allowedOrigins: [] }  →  ADMITTED
```

`SECURITY.md:126-128` is the truthful version: "Anything that can open the socket can
read the mailbox and write drafts."

**Fix:** delete that clause; point at SECURITY.md. State plainly that the checks stop
_browsers_ only, and that the bind address is the only thing stopping a raw client.

### I-2 — `sameOrigin` skips the port comparison when either side omits one · HIGH · reproduced

`src/core/http.ts:70-77`:

```ts
return !from.port || !named.port || from.port === named.port;
```

Browsers omit the default port from `Origin`, and `new URL("http://127.0.0.1").port` is
`""` → `undefined`. **The scheme is never compared either.** Executed against the real
module:

```
ADMITTED        <- page on http://127.0.0.1:80  → mcp on :3939
ADMITTED        <- page on http://localhost     → mcp on localhost:3939
ADMITTED        <- page on https://127.0.0.1    → mcp on :3939   (wrong scheme too)
refused(403)    <- page on http://127.0.0.1:9999 (correct)
```

Any page served from `http://localhost` — a dev server, a router admin UI, a local
preview — can drive the mailbox cross-origin. No CORS headers are needed for the
_write_ path (a `text/plain`-body POST is a simple request; `readJsonBody` never checks
content type, see I-13). Reachable damage without being able to read the response:
`logout` (deletes the session), `create_*_draft`, and `login`, which hands a live device
code to whoever asks (I-4).

**Fix:** compare _effective_ ports (`url.protocol` → 80/443 when `url.port === ""`), and
require the `http:` scheme for this server. Add the three cases above to
`test/http.test.ts` — none are covered today.

### I-3 — The image binds a wildcard interface with no token anywhere · HIGH (design) · reproduced

`Dockerfile` sets `OUTLOOK_MCP_HOST=0.0.0.0` + `EXPOSE 3939` (necessary: a loopback bind
inside a container is unreachable through a published port). So exposure is decided
entirely by the user typing `-p 127.0.0.1:3939:3939`. Combined with I-1, a `-p 3939:3939`
mistake is a full mailbox leak that the README says is covered.

**Fix (highest leverage in this file):** an optional `OUTLOOK_MCP_TOKEN`, compared with
`timingSafeEqual` on the `Authorization` header, **required whenever the resolved bind is
not a loopback address** (refuse to start otherwise, with a message naming the variable).
Have `config docker` print the `-e` and the client `headers`. One control closes I-1,
I-2, I-3 and the rebinding class. Document that it is a bearer secret: plaintext until
someone puts a TLS proxy in front.

### I-4 — `login` / `logout` are reachable on the unauthenticated socket · MEDIUM · traced

`login` returns a live device code and verification URI to any socket peer
(`src/tools.ts:400-441`; `auth_status` also publishes `verificationUri`/`userCode`,
`src/tools.ts:355-360`). `logout` deletes the session (`src/tools.ts:443-460`).
`login` is a _phishing amplifier_: an unauthenticated caller can plant a device code in
the user's own server and talk the user into completing it.

**Fix:** gate `login`/`logout` off the HTTP surface, or let I-3's token cover the whole
port.

### I-5 — `list_attachments` downloads every attachment body in parallel, then throws it away · HIGH · reproduced

`src/core/graph.ts:275`:

```ts
return Promise.all(collected.map(({ id }) => getAttachment(messageId, id)));
```

`getAttachment` (`graph.ts:246`) is a per-attachment GET — the only call that returns
`contentBytes` — and it has no `.select(...)`. `metadataOf` (`src/tools.ts:124-132`) then
discards every byte. The tool description promises "File content is never returned here"
(`src/tools.ts:316-317`), and `src/cli.ts:236` prints the whole result.

One message with 25 × 10 MB attachments ≈ 250 MB decoded / ~333 MB as base64 strings held
at once, with unbounded parallel requests, for a call advertised as metadata-only.

**Fix:** `.select("id,name,size,contentType,isInline")` on the listing (keep the full
fetch for `download_attachment`), and bound the fan-out (~4) instead of bare `Promise.all`.

### I-6 — `search_emails` quotes the whole query, disabling the KQL operators it teaches · HIGH · reproduced

`src/core/graph.ts:232`:

```ts
.search(`"${query.replace(/[\\"]/g, "\\$&")}"`)
```

`src/tools.ts:255-275` instructs callers to use `from:`, `subject:`, `received:`,
`isRead:false`, `hasAttachments:true`. Wrapped in quotes those all become phrase terms,
and the documented `subject:"quarterly report"` becomes
`"subject:\"quarterly report\""`. Every documented narrow query silently degrades to a
full-text phrase match — a wrong _answer_, which is the worst kind for an agent.

**Fix:** pass through unquoted when the query carries no bare syntax, or validate balanced
quotes and pass through — the existing `\` and `"` escaping already covers injection.
Note there is currently no test pinning the quoting, so this is a free fix.

### I-7 — No outbound deadline; one stalled Graph call can pin a session forever · MEDIUM · reproduced (absence)

`sdkClient` (`src/core/graph.ts:126-136`) calls `Client.init({ authProvider })` with no
`fetchOptions`, no `AbortSignal`, no `middlewareOptions`. Nothing else sets a timer.

Path: a hung Graph request keeps `await transport.handleRequest` open → `watch()` leaves
`session.connected > 0` (`src/mcp.ts:245`) → the idle sweep's `connected > 0` test
(`src/mcp.ts:253`) never reaps it → at `maxSessions` 32, `initialize` answers 503 and the
server is dead for everyone.

**Fix:** `fetchOptions.signal = AbortSignal.timeout(ms)` (or a timeout middleware) in
`sdkClient`, plus treat connected-but-silent as idle.

### I-8 — `CHANGELOG.md` says the shipped transport is not implemented · HIGH (docs) · reproduced

`CHANGELOG.md:137` (Known limitations): "Streamable HTTP and native MCP OAuth are not
implemented, so every client launches a local stdio server" — contradicted by
`CHANGELOG.md:19-25` in the same file, by `README.md:27`, `SECURITY.md:120`, and by
`startHttp`. The release note denies the feature the release ships.

**Fix:** keep "native MCP OAuth is not implemented", drop the Streamable HTTP half and
the "every client launches a local stdio server" consequence.

### I-9 — The no-send guard has named gaps · MEDIUM (guard, not live vuln) · reproduced

Nothing sends today — verified: `graph.ts:296-320` issues only `/me/messages` POST,
`createReply`, `createForward`, attachments POST, PATCH, DELETE. But the _guard_ is weaker
than the pitch. Running `test/no-send.test.ts`'s own regexes:

```
sendEndpoint  /me/messages/1/microsoft.graph.send → true
sendEndpoint  /me/sendMail                        → false   ← Graph's real send action
sendEndpoint  /me/messages/1/sendNow              → false

flagged=false  dispatch_message
flagged=false  deliver_post
flagged=false  reply_draft     ← draftOnlyName = /_draft$/ exempts it
flagged=false  forward_draft
```

- `/me/sendMail` matches nothing and is absent from the pinned spellings
  (`test/no-send.test.ts:196-202`, which lists `/send`, `/reply`, `/replyAll`, `/forward`).
- `draftOnlyName = /_draft$/` **exempts** any `_draft`-suffixed name, so a tool that
  really replies and calls itself `reply_draft` passes the name check.
- The mutating-call census reads **only `src/core/graph.ts`** — a new `src/core/send.ts`
  adds no `.post(` there.
- `sourceFiles` collects `.ts` under `src` only (`test/no-send.test.ts:50-56,160`);
  `dist/**.js` — what `bin`, the Docker CMD and the Node-20 job actually execute — is
  never scanned. The `files.length >= 7` floor is a count, not a directory list.

**Fix:** add `sendMail|sendNow|deliver` to both the regex and the pinned list; drop the
`_draft` exemption in favour of the explicit forbidden-name list; scan every `.ts` in the
repo and the built `dist` in CI; assert the scanned _directory set_, not a file count.

---

## 2. Also real, lower priority

### Robustness of the HTTP layer

| #    | Finding                                                                                                                                                            | Evidence                                                         | Fix                                                          |
| ---- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------- | ------------------------------------------------------------ |
| I-10 | `maxSessions` is a TOCTOU: the cap is checked _before_ the async `onsessioninitialized` registration, so N parallel `initialize` POSTs all pass                    | `src/mcp.ts:428` vs `:419-421` — traced                          | reserve a slot synchronously before `await server.connect()` |
| I-11 | 413 returns without draining the body or setting `connection: close`; correctness rests on Node auto-destroying the socket                                         | `src/core/http.ts` `readJsonBody` 413 branch — traced            | `res.setHeader("connection","close")` on the 413 path        |
| I-12 | `void handle(req, res)` — anything thrown _before_ the `try` (`admitRequest`, `new URL(req.url…)`, `respondJson`) is an unhandled rejection that hangs the request | `src/mcp.ts` request entry — traced                              | wrap the whole handler body in try/catch                     |
| I-13 | Content-type and `Accept` are never checked, so an admitted peer gets 4 MiB buffered + `JSON.parse`d before the SDK's 415                                          | `readJsonBody`; SDK rejects only inside `handleRequest` — traced | reject non-`application/json` in `handle` before reading     |
| I-14 | No `requestTimeout` / `headersTimeout` / `keepAliveTimeout` or socket cap; slow-loris rests on Node defaults                                                       | `src/mcp.ts:443-447` sets only `listen` — traced                 | set them explicitly and document the per-connection cost     |
| I-15 | `serve()` calls `redirectConsoleToStderr()`; `startHttp` does not. A future in-process `console.log` lands in `docker logs` in HTTP mode                           | `src/mcp.ts:~225` vs `startHttp` — traced                        | call it in both                                              |

### Core correctness

| #    | Finding                                                                                                                                                                                                                                                                                                                                                                                                                                      | Evidence                                                                                          | Fix                                                                                                                                    |
| ---- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| I-16 | No single-flight on token refresh over a whole-file read-modify-write cache. Two concurrent tool calls with an expired AT each `acquireTokenSilent`; each deserialises the pre-refresh cache and the later rename discards the earlier rotation. The HTTP transport hands one `deps` to every session (`src/mcp.ts:~430`), and `graph.ts:~288` multiplies auth calls per page. Rename gives per-writer atomicity, not lost-update protection | `src/core/auth.ts:100-200` — traced; **undetermined whether msal-node 5.5 serialises internally** | `let inflight` keyed by account+scope in `createAuth`, cleared in `finally`; memoise the token to its expiry. Cheap either way — do it |
| I-17 | Graph-side error taxonomy does not exist. Only MSAL message _text_ is pattern-matched; 401/404/429 flatten to a string, `Retry-After` is dropped, a Graph 401 `invalid_token` never becomes `AuthRequiredError` so it loses the sign-in remediation every other auth failure gets                                                                                                                                                            | `src/core/errors.ts` (whole file), `graph.ts` catches nothing, `mcp.ts:85-90,101-111` — traced    | one adapter reading `statusCode`/`code`/retry-after → typed errors, mapped once                                                        |
| I-18 | `clientId` / `tenant` are validated differently from the port: whitespace-only `OUTLOOK_OAUTH_CLIENT_ID` is truthy and reaches MSAL as `"  "`; a blank tenant yields `https://login.microsoftonline.com//`. Both fail at network time, not config time                                                                                                                                                                                       | `src/core/config.ts:152-154` vs `numberFor` `:44-58` — traced                                     | `(env.X ?? "").trim() \|\| undefined` for both                                                                                         |
| I-19 | `page()` returns `items: response.value` with no `?? []`, so a response without `value` yields `items: undefined`, contradicting the advertised `{ items, nextCursor }`                                                                                                                                                                                                                                                                      | `src/core/graph.ts:195` — traced                                                                  | `?? []`                                                                                                                                |
| I-20 | `addAttachment` decodes the whole base64 into a Buffer (~3 MB) just to measure it                                                                                                                                                                                                                                                                                                                                                            | `src/core/graph.ts:~330` — traced                                                                 | derive the size from string length                                                                                                     |
| I-21 | Attachment dir mode is not re-asserted on a pre-existing directory: `mkdirSync(target, {recursive:true, mode:0o700})` with no follow-up `chmodSync`, so an existing `/tmp/outlook-mcp` at 0755 (or another user's) is used as-is. Contrast `auth.ts:117-119` and `device-login.ts:57-62`, which both chmod after mkdir. Default base sits under `$TMPDIR` (`config.ts:159-160`)                                                              | `src/tools.ts:198` — reproduced (asymmetry)                                                       | `chmodSync(base, 0o700)` after mkdir; move the default out of `$TMPDIR`                                                                |
| I-22 | `beforeCacheAccess` `readFileSync(cachePath)` follows a symlink and never checks mode/owner at read time; same for the device-code state file                                                                                                                                                                                                                                                                                                | `src/core/auth.ts:~112`, `device-login.ts:66-70` — traced                                         | `lstat` and require a regular file, mode ≤ 0600                                                                                        |
| I-23 | `inFlight` requires `expiresAt !== undefined` while `expired()` returns false when it is absent, so a state file without `expiresAt` protects nothing and a second flow can start over a live code                                                                                                                                                                                                                                           | `src/core/device-login.ts:116,78` — traced; **undetermined whether msal always sets `expiresIn`** | treat absent `expiresAt` as not-in-flight _and_ refuse to launch                                                                       |
| I-24 | The lease is rewritten to the child pid _after_ `spawn`; a child that reads state first sees a live lease owned by another pid and refuses to sign in                                                                                                                                                                                                                                                                                        | `src/core/device-login.ts:186-194` vs `:255-265` — traced, timing-dependent (~100 ms boot race)   | write the lease before spawn                                                                                                           |
| I-25 | `to`/`cc`/`bcc`/`categories` are `.min(1)` with no `.max`, and the SDK's element ceiling is off by default                                                                                                                                                                                                                                                                                                                                   | `src/tools.ts:430,470`; `@modelcontextprotocol/sdk/.../server/mcp.js:44-49` — traced              | add a bound                                                                                                                            |
| I-26 | `readonly` vocabulary is `"" \| 0 \| false` with no trim: `off`, `no`, `1` and `"false "` are all read-only. Safe direction, but guessable                                                                                                                                                                                                                                                                                                   | `src/core/config.ts:161` — traced                                                                 | accept an explicit set, reject the rest loudly                                                                                         |
| I-27 | `inside(path, target)` is unreachable-false after `safeFileName` already strips separators — dead code                                                                                                                                                                                                                                                                                                                                       | `src/tools.ts:212` vs `:138-145` — traced                                                         | delete, or keep and note why                                                                                                           |

### Tooling and config

| #    | Finding                                                                                                                                                                                    | Evidence                                               | Fix                                                  |
| ---- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------ | ---------------------------------------------------- |
| I-28 | `config docker` re-derives the port with `Number(...) \|\| defaultHttpPort`, so `OUTLOOK_MCP_PORT=393a` prints `-p 127.0.0.1:3939:3939` while `serve` refuses to start                     | `src/core/clients.ts:~344` vs `config.ts:165` — traced | route through `loadConfig`/`numberFor`               |
| I-29 | `dockerRun` interpolates the literal client id (`clients.ts:127-134`) into a shell command while its own note says "nothing written here carries one" (`:284`) — it lands in shell history | `src/core/clients.ts:127-134,284` — reproduced         | emit `${OUTLOOK_OAUTH_CLIENT_ID}` or the placeholder |
| I-30 | `--host` is unvalidated (only `--port` has `parsePort`), so a typo surfaces as `ENOTFOUND` from `listen`                                                                                   | `src/cli.ts:85` vs `:52` — traced                      | validate or let `listen` error be caught and phrased |
| I-31 | `main()` runs `loadConfig` before dispatch, so an unrelated invalid `OUTLOOK_MCP_PORT` breaks `outlook emails list`                                                                        | `src/cli.ts:76` — traced                               | parse HTTP config lazily                             |
| I-32 | `/health` discloses version, readonly mode and session count to any admitted peer (documented in SECURITY.md, so report-only)                                                              | `src/mcp.ts:355-365` — traced                          | gate behind the token from I-3                       |

### Tests

Coverage gaps, ranked. All **traced**; items 1-3 also have zero coverage at any level.

1. **Concurrency + the 503 path** (`T-1`) — no test issues two concurrent requests, let
   alone 33. `maxSessions` is described in its own comment as the only bound on an
   unauthenticated port.
2. **Oversized body → 413** (`T-2`) — no test; `grep` for `413` / `maxRequestBytes` /
   `readJsonBody` across `test/` returns nothing.
3. **Idle-session sweep** (`T-3`) — the janitor (`src/mcp.ts:330-347`) is untested;
   `OUTLOOK_MCP_SESSION_IDLE_MINUTES` is only tested as a number parse
   (`test/config.test.ts:110-124`).
4. **Config → admission wiring** (`T-4`) — `admitRequest` is fed hand-built arrays
   (`test/http.test.ts:177-181,226-232`); no test starts a server with
   `OUTLOOK_MCP_ALLOWED_HOSTS`/`_ORIGINS` in the env. The only end-to-end Origin check is
   `docker.yml:70-74`.
5. **Bind address / URL construction** (`T-5`) — the `0.0.0.0`/`::` → `127.0.0.1` rewrite
   (`src/mcp.ts:470-478`) and an IPv6 `::1` bind are untested; `[::1]` appears only as a
   string in `test/http.test.ts:154`.
6. **Batch/array bodies and valid-JSON-not-JSON-RPC** (`T-6`) — `isInitialize`
   (`src/mcp.ts:294-304`) deliberately supports batches; untested.
7. **Shutdown** (`T-7`) — `handle.close()` is used only as teardown; the `closed` promise
   is never asserted, and `closeAllConnections()` with an open SSE `GET` is untested.

Tautological tests — **reproduced** by reading the definitions:

- `test/clients.test.ts:93-98`, _"every client that starts the server launches the built
  entry with node"_ — `launchedLocally` is _defined_ (`:84-86`) as clients whose snippet
  `includes("node")`, then the loop asserts `includes("node")`. Provably unfailable.
  Escapes: a snippet whose `command` becomes `npx tsx` / `bun` while "node" survives in a
  note or path segment. (The `deepEqual` at `:90-92` is _not_ tautological.)
- `test/no-send.test.ts:126-129` — compares `tools.map(name)` against
  `[...readToolNames, ...draftToolNames]`, which `src/tools.ts:752-753` derives from the
  same two arrays. An identity: passes if both arrays are emptied.
- `test/no-send.test.ts:179-208`, _"the scan really distinguishes the create actions from
  the sending ones"_ — two hand-written lists probed against the regex the guard already
  uses, limited to the five spellings it was written for. It certifies itself, and it is
  the test that makes the guard look validated (see I-9).
- Honourable mention: `test/mcp.test.ts:~520-524` asserts `get_email` returns exactly what
  the stub was coded to return, with no production transform in between.

Hygiene:

- `test/http.test.ts` `raw()` (`:110-145`) and `healthOf()` (`:472-480`) have no timeout,
  and `node:test`'s default test timeout is Infinity — a regression that accepts the
  socket and never answers hangs `npm test` until the CI job limit with nothing naming the
  test. **Medium.**
- `test/clients.test.ts` `home()` (`:25-27`) mkdtemps a directory per `target()` call with
  no `t.after` cleanup — dozens of permanent dirs per run. Low.
- `e2e/http.e2e.test.ts:76-100` polls stderr every 25 ms with an `unref()`d 20 s
  `givenUp`; on a loaded runner the exit-promise branch can win first, so only the
  diagnostic differs between runs. Low.
- Correct, do not re-investigate: ephemeral ports everywhere (`test/http.test.ts:101`,
  `e2e/http.e2e.test.ts:32` with the URL parsed from stderr), `OUTLOOK_MCP_READONLY: ""`
  explicitly cleared, no `process.env` mutation in any test, no fixed sleeps waiting on
  real time.

### CI matrix

- **Node 20 never starts the MCP server** (`C-1`) — `ci.yml:38-57` is build + four CLI
  invocations. No stdio `initialize`, no `serve --http` bind, no unit tests (they import
  `.ts`, which Node 20 cannot strip). So `engines: ">=20"` is exercised only as argument
  parsing, and a Node-20-only failure in `src/mcp.ts` / `http.ts` / `graph.ts` ships.
  **Smallest fix:** run `npm run test:e2e` with `OUTLOOK_E2E_BIN=dist/index.js` in that
  job — the e2e suite already runs the built entry.
- **The container is never MCP-handshaken** (`C-2`) — `docker.yml:36-96` is good
  (`/health`, Host/Origin 403 through a real published port, `--help`, `config docker`,
  no `.env`, empty `/data`), but `/health` is answered _before_ the transport
  (`src/mcp.ts:376-386`), so an image whose `/mcp` cannot serve still passes.
- **The tarball is never installed or run** (`C-3`) — `ci.yml:71-108` checks _contents_
  and manifest fields. A tarball whose `bin` throws on import passes. **Smallest fix:**
  `npm i -g "$tarball" --omit=dev --ignore-scripts && outlook-mcp --help`.
- `docker.yml:124-131` publishes `linux/arm64` that no job ever executes (acknowledged at
  `:32-34`).
- No linter anywhere — `format:check` is prettier only.

### Documentation contract

| #    | Claim                                                                                                                  | Doc                                                                                         | Code                                                                                                                                                                                                                          | Consequence                                                                                                   |
| ---- | ---------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| D-1  | "`--write` refuses and tells you to paste" when a location is not stated with confidence                               | `docs/clients.md:39-40`                                                                     | refusal happens only when a definition has **no `path`** (`clients.ts` `installClient` ~455-465 → `pi`, `docker`); `vscode`, `cursor`, `opencode` all define `path` **and** carry an `unverified` note and are written anyway | a user's config files are edited at guessed locations while they believe the tool declined                    |
| D-2  | "`type: "http"` … is the form Claude Code, VS Code and Pi document"                                                    | `docs/clients.md:338`                                                                       | the provenance rule (`clients.ts:11-18`) plus `docs/clients.md:378`, which lists the docker remote shape as _not_ verified                                                                                                    | the doc asserts documentation it says elsewhere it never read                                                 |
| D-3  | State file is `~/.outlook-skill/state/login.json`                                                                      | `SECURITY.md:74`                                                                            | `loginStatePath` writes `login-<MailScope>.json` (`config.ts:113-119`), e.g. `login-Mail.ReadWrite.json`                                                                                                                      | anyone scripting backup/cleanup/audit targets a filename that does not exist and concludes nothing is on disk |
| D-4  | `serve --host <host>`                                                                                                  | _documented nowhere_ — zero `--host` hits in any `*.md`                                     | `src/cli.ts:~80`, defaults to `config.httpHost`                                                                                                                                                                               | a wildcard-bind user sees half the flag surface; README documents `--port` at `:223`                          |
| D-5  | Quickstart uses `ghcr.io/vincentkelleher/outlook-skill:0.1.0`                                                          | `README.md:80`, `docs/clients.md:317`, `setup.md:94`, and `defaultImage` (`clients.ts:26`)  | `CHANGELOG.md:7` — "0.1.0 — not yet tagged"; `docker.yml` pushes on `v*` or `latest`/`sha-` only                                                                                                                              | copy-paste quickstart fails with `manifest unknown`. Fix the docs to `:latest` or tag the release             |
| D-6  | Release asset `.../download/v0.1.0/outlook-skill-0.1.0.tgz`                                                            | `docs/clients.md:52-56`                                                                     | same root cause as D-5                                                                                                                                                                                                        | 404 on a fresh install                                                                                        |
| D-7  | `README.md:3` embeds `./doc/banner.png`                                                                                | not in `package.json files` (`:26-34`), `.dockerignore`d                                    | —                                                                                                                                                                                                                             | a tarball reader gets a broken image; the docs it _does_ promise are correctly packed and CI asserts it       |
| D-8  | `OUTLOOK_MCP_PORT` "`serve --http` only"                                                                               | `README.md:223`                                                                             | also read by `config docker` (`clients.ts:344`), where it is **not** validated                                                                                                                                                | see I-28                                                                                                      |
| D-9  | `OUTLOOK_MCP_IMAGE`                                                                                                    | `README.md:227` only                                                                        | `clients.ts:352`                                                                                                                                                                                                              | the Docker section of `docs/clients.md:311-345`, which prints the run command, never mentions it              |
| D-10 | Valid ranges for the numeric env vars                                                                                  | —                                                                                           | `numberFor` throws (`config.ts:44-58`, `:60-62`)                                                                                                                                                                              | `OUTLOOK_MCP_PORT=393a` aborts startup with no documented range; `README.md:223,226` present the vars bare    |
| D-11 | `SKILL.md:31-34` lists config targets as "Claude Code, Claude Desktop, Codex, VS Code, Cursor, Gemini CLI or OpenCode" | `clientIds` also has `docker` and `generic` (`clients.ts:31-42`), and `serve --http` exists | —                                                                                                                                                                                                                             | an agent following the skill cannot know the HTTP image exists                                                |

---

## 3. Verified correct — do not re-investigate

- **DNS rebinding is blocked by ordering.** The Host rule runs _before_ the same-origin
  rule (`http.ts admitRequest`), so a rebound `Host: evil.example` + `Origin:
http://evil.example` is 403 even though `sameOrigin` would match. Covered by
  `test/http.test.ts` ("a Host header naming another machine is refused before the body")
  and by `docker.yml:65-79`, which asserts 403 for both headers against the real container
  behind a real published port.
- `Origin: null` (opaque / file / sandboxed) is rejected — `originAuthority` returns
  undefined → 403. Repeated or absent `Host` → 400.
- **No CORS surface.** Nothing sets `Access-Control-Allow-*`; `OPTIONS` is not `POST`, so
  it hits the session lookup and 404s with no ACAO. Browser preflight cannot succeed.
- **A cursor cannot be injected into another request or host.** `assertGraphCursor`
  (`graph.ts:~106`) demands `https:` + `hostname === graph.microsoft.com` + empty
  port/user/pass + a pathname matching `graphPagePath`, and rejects before any call;
  `test/graph.test.ts:347-380` asserts a hostile list yields the error **and**
  `assert.deepEqual(requests, [])`. With a cursor present, `folder`/`top`/`unread` are
  ignored, so options cannot be smuggled. (A forged cursor may still carry arbitrary query
  options on its own `/me/messages` GET — a read inside the caller's existing grant, not an
  escalation. A stale-but-well-formed cursor is accepted by design.)
- **No odata/path injection anywhere else.** Folder names pass `folderPattern` +
  `encodeURIComponent`; ids are encoded; the only `$filter` in the codebase is the constant
  `"isRead eq false"`.
- **Validation precedes every network call, and limits are bounded.** The SDK zod-parses
  before invoking the handler; `limit` is `.int().min(1).max(…)` on every listing tool;
  the constants are exported once from `graph.ts` so prose and ceiling cannot disagree;
  the CLI validates too; `add_attachment`'s size check runs before the request.
- **Attachment containment is not a prefix-only test** (`tools.ts:185-229`):
  `resolve(base, dir)` → prefix check → `realpathSync` on both sides → `safeFileName`
  strips `/`, `\`, control chars including NUL, and leading dots; no URL decoding anywhere
  so percent-encoding is inert; a symlink already at the destination is rejected by
  `lstatSync` and the write uses `O_NOFOLLOW` with mode 0600.
- **Token-cache write is race-free per writer**: temp file with `mode: 0o600` → `chmodSync`
  → `renameSync`, dir `mkdir 0700` + `chmod 0700`, stale-temp sweep. State dir likewise,
  state file carries URL/code/error and never a token, and claim is `wx`-atomic.
- **Config precedence is single-sourced and correct.** `loadEnvFile` seeds from
  `process.env` and skips keys already set, so `.env` cannot override a real variable;
  both entry points route through it; commander defaults are the config values.
- **Read-only is one coherent switch.** `scopesFor` drops to `Mail.Read`, `toolsFor`
  removes every draft tool, the MCP instructions are _generated_ from
  `readToolNames`/`draftToolNames` rather than hand-copied, `auth_status` reports
  `readonly` + `scopes`, and login state is namespaced per mail scope so the two modes
  cannot block each other. No reachable disagreement between advertised tools and mode.
- **Sign-in cannot clobber a valid session.** `runLogin` → `accessToken({interactive:true})`
  tries silent first; the tool short-circuits on `signedIn` and starts nothing when the
  session could not be checked; a second foreground login against a live flow bails with an
  explicit message; `claimLaunch` uses `wx`. Failed flows never touch the token cache.
- **MSAL error wording has exactly one decider** (`errors.ts:30-55`), `needsConsent`
  requires Microsoft to have said so, and `cause` is preserved. The gap is Graph-side only
  (I-17).
- **No token exists at all** — confirmed by reading all of `http.ts` and `mcp.ts`: the only
  credential is `mcp-session-id`, generated by `randomUUID()` (CSPRNG, so no
  `==`-vs-`timingSafeEqual` question). `SECURITY.md` states the model plainly.
- **HTTP exposes nothing stdio does not** — both use the same `createServer(deps)`;
  read-only hides the draft tools identically over HTTP, asserted in `test/http.test.ts`.
- **No stack traces to clients**; `respondJson` sets `cache-control: no-store`; the janitor
  is `unref()`d and `close()` clears the timer and calls `closeAllConnections()`.
- **Secrets in images:** the Dockerfile bakes no client id (CI injects `00000000-…`), no
  `.env`, no session; `.dockerignore` excludes `.env`, `.env.*`, `.git`, `.github`, `test`,
  `e2e`; CI asserts `test ! -e /app/.env` and that `/data` is empty; `USER node`, `/data`
  0700 node-owned. `.env` is untracked.
- **The never-send invariant holds in the shipped code** (see I-9 for the _guard_'s gaps):
  `graph.ts:296-320` issues only draft create, `createReply`, `createForward`, attachments
  POST, PATCH, DELETE. No `send`/`reply`/`forward` action exists in the file.
- **`docs/plan-mcp-skill.md` is correctly handled** — a "historical record … this one is
  out of date" banner, excluded from `files`, `.dockerignore`d, and CI asserts its absence
  from the tarball. Stale-but-labelled; no need to delete.
- **The Node floor is coherent**: `engines: ">=20"`, `SKILL.md:4` "Node 22.18 or newer to
  run the sources, or Node 20 with a built dist", CI 22/24 full + Node 20 built-entry-only.
  (C-1 above is about _coverage_ of that claim, not its wording.)
- **`private: true` + "installing from GitHub is the only supported path" + `npx skills
add` do not conflict** — `private` only blocks npm publish, which `README.md:283`
  explicitly disclaims.
- **All 12 env vars, both directions** — every `OUTLOOK_*` name in `README.md:216-227` is
  read, character-for-character, at `config.ts:152-176` and `clients.ts:344,352`; no
  documented var is unread; `OUTLOOK_MCP_LOG_LEVEL` was genuinely removed. Defaults
  `127.0.0.1` / `3939` / `1440` match the exported constants.
- **Tool contract** — 15 tools in README's table = 15 registered (8 core + 7 draft).
- **No `--readonly` flag exists**, exactly as `README.md:243` claims.
- **`config docker`** prints exactly `{"type":"http","url":"http://127.0.0.1:3939/mcp"}`
  and `--write` refuses it; `codex` really sets `startup_timeout_sec = 30`; `vscode` uses
  root key `servers` + `"type":"stdio"`; `opencode` uses `mcp` + command-as-list.
- **SECURITY.md's new HTTP section matches code** — 4 MiB cap, 32 sessions, idle sweep with
  `unref`, `/health` returning only version/mode/session count, no TLS, refusal text naming
  the variable to set.
- **Strongest tests, keep as-is:** `e2e/mcp.e2e.test.ts:117-171` (byte-level stdout
  accounting on real stdio pipes — exactly three frames, no non-JSON line, NUL check);
  `test/graph.test.ts:311-345` (11 hostile cursor spellings closed by
  `deepEqual(requests, [])` — proves no request was built); `test/mcp.test.ts:598-700`
  (attachment traversal/backslash/dot-only names and the pre-existing symlink, asserting
  real on-disk contents); `test/http.test.ts:287-330, 332-379, 429-445` (a real SDK
  `Client` over `StreamableHTTPClientTransport`, DELETE retiring a session, reused id →
  404, `EADDRINUSE` failing the start not the first request); `docker.yml:65-79`;
  `test/no-send.test.ts:106-124` (the hand-written seven-name pin — deliberately the copy
  that cannot move with the code).

---

## 4. Residual risk — open questions nobody answered

- **Does msal-node 5.5.0 serialise custom-cache access per `PublicClientApplication`?**
  Decides whether I-16 is a genuine lost update or merely N redundant cache reads plus our
  own unsynchronised retry. Cross-process concurrency (two client processes sharing
  `msal-cache.json`) has the same shape and is untested either way.
- **Does `@microsoft/microsoft-graph-client` 3.0.7 put `RetryHandler` in the default
  middleware chain** for `Client.init({authProvider})`, and with what ceiling? A grep found
  a `RetryHandler` constructed at `es/src/HTTPClientFactory.js:46`, but that path returns
  ENOENT in this tree (`src/`/`lib/` builds) and the no-handlers branch was not confirmed.
  This decides whether a throttled call blocks without limit.
- **Does msal's `deviceCodeCallback` response always carry `expiresIn`?** Decides I-23.
- **Effective Node timeouts** (`requestTimeout` / `headersTimeout` / `keepAliveTimeout`)
  are unset in this repo and were never measured.
- **Can any MSAL/Graph `Error.message` that reaches a tool result or the 500 body contain
  a bearer fragment?** Confirmed "no stack traces, no logger configured" — not "no token in
  any upstream message string".
- Whether Graph requires `$count=true` alongside `ConsistencyLevel: eventual` for `$search`
  on messages (`graph.ts:~240` sets the header only).
- **`dist/**.js` was never reviewed and is never scanned** — it is what `bin`, the Docker
  CMD and the Node-20 CI job actually execute. Committed history was out of scope: the
  review covered the working tree.
- Real on-disk modes of `~/.outlook-skill`, `state/`, and `$TMPDIR/outlook-mcp` on a target
  machine (umask-dependent, not measured).
- Whether `npx skills add` resolves `skills/outlook-skill/SKILL.md` as `README.md:290`
  claims (external installer behaviour, untested).
- `README.md:120-121` "`docker stop` takes it down in well under a second" — SIGTERM
  handlers exist, shutdown latency was never measured.
- Line numbers from the two revived review lanes carry ±3 slack; symbol names were given
  for disambiguation. Every item marked **reproduced** above was re-checked directly.

---

## 5. Suggested order of work

**Commit A — documentation and one-line fixes, no behaviour change.**
I-1, I-8, D-1..D-10, I-29. Low risk, and I-1 in particular should not survive a release.

**Commit B — admission correctness.**
I-2 (effective ports + scheme) with the three reproduced cases added to
`test/http.test.ts`, plus T-4 (allow-list driven from the env, not hand-built arrays).

**Commit C — the bearer token (I-3, I-4, I-32).**
New behaviour, needs docs and `config docker` changes. This is the change that makes the
HTTP story defensible; I-1's wording gets to be much shorter afterwards.

**Commit D — Graph call hygiene: one outbound policy in `sdkClient`.**
I-7 (deadline) and I-17 (typed errors, `Retry-After`) are the same seam and should land
together. I-16 (single-flight) fits here or alone.

**Commit E — content budget.**
I-5 (metadata-only listing + bounded fan-out), I-19, I-20, I-21, I-22.

**Commit F — search semantics (I-6).**
Its own commit: it changes what users get back, and deserves a CHANGELOG entry of its own.

**Commit G — the guard (I-9) and the test suite.**
Guard spellings + drop the `_draft` exemption + scan scope; fix or delete the three
tautological tests; add T-1, T-2, T-3; C-1 (Node 20 e2e), C-2 (container handshake),
C-3 (install the tarball); timeouts on the raw test helpers.

I-10..I-15, I-12, I-13, I-14, I-25..I-28 are independent small fixes that can ride along
with whichever commit is nearest their seam.
