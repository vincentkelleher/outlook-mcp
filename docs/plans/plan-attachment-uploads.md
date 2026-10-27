# Plan: put files on a draft — uploads, sources, and the way back

Status: **agreed and built.** The code, `README.md` and `SECURITY.md` are
authoritative; where this document disagrees with them, they are right. It is
kept for the reasoning behind the choices. Decisions locked with the maintainer:

| Decision        | Choice                                                                                                    |
| --------------- | --------------------------------------------------------------------------------------------------------- |
| Byte sources    | **All three**: `contentBase64`, `serverPath`, `fromAttachment` — one `add_attachment`, exactly one source |
| Size ceiling    | New **`OUTLOOK_MCP_MAX_UPLOAD_BYTES`**, default **35 MiB**, never above Graph's **150 MB**                |
| Undo            | **Yes** — new `remove_attachment`, marked destructive                                                     |
| Safety addition | A write to a message Graph says is **not a draft** is refused (`isDraft: false`)                          |
| Deliverable     | Proposal + implementation in one change                                                                   |

## 1. Where we are

`add_attachment` already exists: `POST /me/messages/{id}/attachments` with
`contentBase64`, and a **hard refusal at 3 MB** (`singleRequestAttachmentBytes`)
whose error says "Graph needs an upload session from 3 to 150 MB, which this
server does not implement". So the feature is not missing, it is capped, and the
cap is the documented gap.

Three things follow from reading the code rather than the README:

1. **The cap is the wrong limit to lift for agents.** A 20 MB file is 27 MB of
   base64 that a model would have to emit token by token. Raising the ceiling
   alone only serves a client that builds arguments by hand. The useful change
   is a way to attach bytes that never enter a context window.
2. **A mistake cannot be undone.** There is no `DELETE` on an attachment, so a
   wrong file stays on the draft, and re-adding lands a second `quote.pdf`
   beside it (`list_drafts` shows `hasAttachments: true` and nothing more).
3. **Nothing checks that the target is a draft.** `POST …/attachments` works on
   a sent message too. Today that only bites someone who points a message id at
   `add_attachment` by accident; a `remove_attachment` would let the same slip
   delete a file from mail that has already gone out. That is worth closing in
   the same change.

## 2. The Graph contract

Source: [Attach large files to Outlook messages](https://learn.microsoft.com/en-us/graph/outlook-large-attachments),
[attachment: createUploadSession](https://learn.microsoft.com/en-us/graph/api/attachment-createuploadsession?view=graph-rest-1.0).

| File size     | How                                                                             |
| ------------- | ------------------------------------------------------------------------------- |
| under 3 MB    | one `POST /me/messages/{id}/attachments` with `contentBytes`                    |
| 3 MB – 150 MB | `POST …/attachments/createUploadSession`, then sequential `PUT`s of byte ranges |
| over 150 MB   | nothing — Graph itself refuses, so this server refuses first                    |

The small file goes up as one `fileAttachment`, and the field holding its bytes
is called `contentBytes` on the wire — not `contentBase64`, which is what the
tool's own argument is called. Graph reads the body for `contentBytes`, finds
none, and answers `400 ErrorRequiredPropertyMissing`, "Required property is
missing.", with nothing attached:

```http
POST /v1.0/me/messages/{id}/attachments
{ "@odata.type": "#microsoft.graph.fileAttachment", "name": "note.txt",
  "contentType": "text/plain", "contentBytes": "aGVsbG8=" }
```

Session mechanics, in the words of the document that defines them:

```http
POST /v1.0/me/messages/{id}/attachments/createUploadSession
{ "AttachmentItem": { "attachmentType": "file", "name": "report.pdf",
                      "contentType": "application/pdf", "size": 3483322 } }

201 → { "uploadUrl": "https://outlook.office.com/…/AttachmentSessions('…')?authtoken=eyJ…",
        "expirationDateTime": "…", "nextExpectedRanges": ["0-"] }

PUT {uploadUrl}                       Content-Type: application/octet-stream
    Content-Range: bytes 0-2097151/3483322      body: the raw bytes of that range
200 → { "nextExpectedRanges": ["2097152"] }     (more to come)
201 → Location: …/Attachments('AAMkAD…=')       (done; the id is in the URL)
```

Four properties of that contract shape the implementation:

- **`uploadUrl` is a credential.** It is pre-authenticated and carries a bearer
  token in its query string. `Do not specify an Authorization request header`,
  the docs say — so the chunk `PUT` cannot go through the SDK builder, which
  adds one to every call. It goes through `fetch`, and `uploadUrl` / `authtoken`
  join `neverLoggedKeys` so neither can reach a log line at any level.
- **Ranges are sequential.** Out of order is an error, so the pump never
  parallelises, and it follows the server's `nextExpectedRanges` rather than its
  own arithmetic where the two disagree — which is also what makes a resumed
  chunk safe instead of a guaranteed 416.
- **Ranges are small by preference.** "Keep each byte range less than 4 MB", so
  `uploadChunkBytes` is **3.75 MiB** (`320 KiB × 12`): under 4 MB and a multiple
  of 320 KiB, the alignment every Graph upload session asks for.
- **The mailbox has its own ceiling, below Graph's.** Outlook.com and Microsoft
  365 refuse an oversized message at **Send** — around 20 MB and 35 MB
  respectively — so a 140 MB upload succeeds and then strands a draft nobody can
  send. The default cap is 35 MiB for exactly that reason: the failure lands
  here, before the bytes, with the setting that changes it named.

## 3. The tool

One tool, three sources, exactly one of them. A discriminated union spelled as a
`strictObject` with a refine, because the registry is one flat list of zod
schemas shared by both surfaces.

```jsonc
// add_attachment
{
  "messageId": "draft-1", // must be a draft
  "name": "report.pdf", // what the recipient sees
  "contentType": "application/pdf", // optional; see the defaults
  "contentBase64": "…", // A. small, and in the caller's hands already
  "serverPath": "downloads/quote.pdf", // B. inside OUTLOOK_MCP_ATTACHMENT_DIR
  "fromAttachment": {
    // C. a copy, server-side
    "messageId": "mail-1",
    "attachmentId": "att-1",
  },
}
```

```jsonc
// every success reads back the draft, and the draft's whole attachment list
{
  "id": "draft-1",
  "subject": "Q3",
  "to": ["ana@example.com"],
  "sent": false,
  "note": "This is a draft: nothing has been sent…",
  "attachment": {
    "name": "report.pdf",
    "contentType": "application/pdf",
    "bytes": 3483322,
    "transport": "upload_session",
  },
  "attachments": [{ "id": "AAMkAD…", "name": "quote.pdf", "size": 41219 }],
}
```

`attachments` is a fresh metadata read of the draft, so the id needed by
`remove_attachment` is always in the answer to `add_attachment`. It is a `GET`,
and a failure to list is never allowed to turn a successful write into an error:
the field is dropped and the note says so.

| Source           | Bytes reach the model? | Bound by                                              | Transport                              |
| ---------------- | ---------------------- | ----------------------------------------------------- | -------------------------------------- |
| `contentBase64`  | **yes**                | `OUTLOOK_MCP_MAX_UPLOAD_BYTES`                        | under 3 MB POST, otherwise the session |
| `serverPath`     | no                     | `OUTLOOK_MCP_MAX_UPLOAD_BYTES`, file size from `stat` | streamed, 3.75 MiB at a time           |
| `fromAttachment` | no                     | `min(maxAttachmentBytes, maxUploadBytes)`             | fetched once, then the session         |

- **`serverPath` is a read, and a read is a new surface.** It is confined by the
  same guard that confines downloads, in `src/core/files.ts`: resolved inside
  `OUTLOOK_MCP_ATTACHMENT_DIR`, absolute paths and `..` refused, the real path
  rechecked, a symlink at the destination refused, `O_NOFOLLOW` on the open, and
  only a regular file accepted. So the workflow it opens is exactly
  `download_attachment delivery=server_file` → `add_attachment serverPath`, and
  nothing else: a model cannot point it at `~/.ssh/id_rsa` and mail it out.
- **`fromAttachment` inherits the download limit.** Copying fetches the source
  through the existing, tested `attachmentContent`, which is memory-bounded by
  `OUTLOOK_MCP_MAX_ATTACHMENT_BYTES` (20 MiB). A 30 MB copy is therefore refused
  today. Streaming the raw `GET …/attachments/{id}/$value` range-by-range would
  lift that, and is the obvious follow-up; it is not in this change, and the
  error names the setting rather than hiding behind a generic failure. The
  metadata read that settles the size is passed into `attachmentContent` as
  `prefetched`, so a copy is two Graph calls — metadata, then content — not three.
- **`name` and `contentType` default from the source** when they are not given —
  the leaf of `serverPath`, or the received attachment's own name and type,
  either one passed through `safeFileName`. An explicit value always wins.
- **`remove_attachment`** takes `messageId` + `attachmentId`, refuses a target
  Graph says is not a draft, deletes, and reads back what is left.
  `destructiveHint: true`, because on a draft it destroys a file the human may
  have put there.

Both tools sit in `draftTools`, so **read-only mode advertises neither** and the
scope needed is the one already requested: `Mail.ReadWrite`, which the upload
session documentation confirms is the permission for a message session. No new
scope, no new consent prompt.

## 4. Where the code goes

```
src/core/files.ts      NEW — every byte this server reads or writes on disk:
                        containment, O_NOFOLLOW write, O_NOFOLLOW read
src/core/uploads.ts    NEW — the three sources → one { name, contentType, size, read() }
                        contract; no HTTP, no Graph types beyond UploadSource
src/core/graph.ts      upload session (POST) + chunk (fetch PUT) + attachment DELETE
                        + isDraft guard; owns every Graph call and the pump loop
src/core/config.ts     maxUploadBytes / OUTLOOK_MCP_MAX_UPLOAD_BYTES
src/core/log.ts        uploadUrl, authtoken never logged; serverPath is content
src/tools.ts           add_attachment reworked, remove_attachment added
```

`UploadSource` is declared in `graph.ts` (it is the contract of the Graph call)
and imported as a type by `uploads.ts`, so the dependency runs one way:
`tools → uploads → graph`, the same shape `attachments.ts` already has.

The pump loop lives in `graph.ts` beside the two calls it drives, not in
`uploads.ts`, for one reason: `test/no-send.test.ts` counts the mutating call
sites in `src/core/graph.ts`, and a write that moved to a module the guard does
not read is a write nobody reviewed.

## 5. The guard, updated

`no-send.test.ts` pins the write surface from three sides, and this change moves
all three. Each move is the test doing its job — a new write has to be argued
for, in the test, before it can exist.

| Assertion                         | Before                       | After                                                                      |
| --------------------------------- | ---------------------------- | -------------------------------------------------------------------------- |
| draft tools                       | 7 names, written out by hand | 8 — `remove_attachment` added                                              |
| mutating call sites in `graph.ts` | 7: `delete, patch×2, post×4` | 9: `delete×2, patch×2, post×5`                                             |
| writes that regex cannot see      | —                            | exactly **one** `fetch` `PUT`, and it must carry no `Authorization` header |
| prose of the instructions         | `… and add_attachment`       | `… add_attachment and remove_attachment`                                   |

The `PUT` is the interesting line. `.put(` never appears — the chunk goes out as
`fetch(url, { method: "PUT" })` — so the guard gains a second scan for
`method: "PUT"` and asserts both its count and the absence of an Authorization
header on it. Without that addition the change would quietly halve the coverage
of the project's central safety test, which is the opposite of why it exists.

Also unchanged, and worth restating because `fromAttachment` touches it:
nothing here calls `/send`, `/reply`, `/replyAll` or `/forward`.
`…/attachments/createUploadSession` is not a transmitting action, and
`createForward` already carries the original's attachments on its own.

## 6. Failure behaviour

| What happened                                      | What the caller gets                                                                  |
| -------------------------------------------------- | ------------------------------------------------------------------------------------- |
| file above `maxUploadBytes`                        | refused before any request, size + limit + env name + Graph's 150 MB, draft unchanged |
| `serverPath` outside the attachment directory      | refused, with the real path that was refused, and a `warn`                            |
| `serverPath` is a symlink, a directory, or missing | refused, each with its own reason                                                     |
| zero-byte file                                     | refused — `bytes 0--1/0` is not a range                                               |
| target is not a draft (`isDraft: false`)           | refused, no write attempted                                                           |
| a `PUT` returns neither 200 nor 201                | error with the status and a short body excerpt, `authtoken=` struck out               |
| `nextExpectedRanges` does not advance              | error rather than a loop that never ends                                              |
| session expired mid-upload                         | error saying the session is gone and the add must be retried                          |

An interrupted session leaves **nothing** attached: Graph materialises the
attachment when the last byte lands, so a failed add is a no-op on the draft
rather than half a file. `transport` in the echo says which route the bytes took,
which is the first thing worth knowing when a draft behaves oddly.

## 7. Build order (as it went)

1. `files.ts` — extract the download guard untouched, add the read half, keep the
   existing tests green as they move.
2. `config.ts` + `graph.ts` — ceiling, constants, session, pump, delete, draft
   guard; `graph.test.ts` with an injected `fetch`.
3. `uploads.ts` — the three sources.
4. `tools.ts` — schema, `remove_attachment`, echoes; then `log.ts` redaction.
5. `no-send.test.ts` — updated counts, the new `PUT` scan.
6. Docs: README (tools table, "Putting files on a draft", env table),
   SECURITY.md (attachment limits, the new read surface, the guard recount),
   `skills/outlook-mcp/SKILL.md`.

Not in this change, deliberately: streaming `$value` copies; attaching several
files in one call (each file has its own size and its own failure, and the
per-file answer is what a caller needs); chunked arguments as an upload
protocol (MCP has no binary argument, so N chunks cost the same context as one
big one); anything that sends.
