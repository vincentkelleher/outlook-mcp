import assert from "node:assert/strict";
import { test } from "node:test";
import {
  createGraph,
  defaultMaxUploadBytes,
  defaultPageSize,
  maxFlagReasonChars,
  maxPageSize,
  messagesPath,
  singleRequestAttachmentBytes,
  uploadChunkBytes,
} from "../src/core/graph.ts";
import type {
  FetchLike,
  GraphClient,
  GraphRequest,
  UploadSource,
} from "../src/core/graph.ts";
import { AuthRequiredError, noSessionMessage } from "../src/core/errors.ts";
import { GraphError } from "@microsoft/microsoft-graph-client";
import { captureLogs } from "./support/log-capture.ts";

const messageFields =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments";
const fullMessageFields = `${messageFields},toRecipients,ccRecipients,body`;
const draftFields = `${messageFields},isDraft,webLink,toRecipients`;

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

type Recorded = {
  path: string;
  verb: "get" | "post" | "patch" | "delete";
  headers: string[][];
  select?: string[];
  orderby?: string;
  top?: number;
  filter?: string;
  search?: string;
  body?: unknown;
};

/**
 * The rule Graph applies to `POST …/attachments`: a `fileAttachment` with no
 * `contentBytes` has no content as far as the service is concerned, and the
 * answer is the bare "Required property is missing." (`ErrorRequiredPropertyMissing`)
 * with nothing attached to the draft.
 *
 * This is here because a fake that accepted any body let the field be spelled
 * `contentBase64` — the name of the tool's own argument, not of the wire field —
 * and every test pass while no real draft ever took a file. `@odata.type` is
 * checked with it: Graph needs the discriminator to know which attachment type
 * the body describes.
 */
function storedAttachment(body: unknown): { id: string } {
  const attachment = (body ?? {}) as Record<string, unknown>;
  const missing = ["@odata.type", "name", "contentBytes"].filter(
    (property) =>
      attachment[property] === undefined || attachment[property] === "",
  );
  if (missing.length > 0)
    throw new Error(
      `{"error":{"code":"ErrorRequiredPropertyMissing","message":"Required property is missing. ${missing.join(", ")}"}}`,
    );
  return { id: "attachment-1" };
}

function responder(path: string, body?: unknown): unknown {
  if (path === "/next-attachments") return { value: [{ id: "two" }] };
  // A POST to the collection adds one attachment; a GET lists them.
  if (path.endsWith("/attachments") && body !== undefined)
    return storedAttachment(body);
  if (path.endsWith("/attachments"))
    return {
      value: [{ id: "one/two" }],
      "@odata.nextLink": "/next-attachments",
    };
  if (path.includes("/attachments/"))
    return { id: "attachment", contentBytes: "ZmlsZQ==" };
  if (path.includes("/messages/")) return { id: "message" };
  return { value: ["message"] };
}

function fakeClient(
  respond: (path: string, body?: unknown) => unknown = responder,
) {
  const requests: Recorded[] = [];
  const client: GraphClient = {
    api(path) {
      const request: Recorded = { path, verb: "get", headers: [] };
      requests.push(request);
      const chain: GraphRequest = {
        header(name, value) {
          request.headers.push([name, value]);
          return chain;
        },
        select(fields) {
          request.select = fields;
          return chain;
        },
        orderby(value) {
          request.orderby = value;
          return chain;
        },
        top(value) {
          request.top = value;
          return chain;
        },
        filter(value) {
          request.filter = value;
          return chain;
        },
        search(value) {
          request.search = value;
          return chain;
        },
        async get<T>() {
          return respond(path) as T;
        },
        async post<T>(body: unknown) {
          request.verb = "post";
          request.body = body;
          return respond(path, body) as T;
        },
        async patch<T>(body: unknown) {
          request.verb = "patch";
          request.body = body;
          return respond(path, body) as T;
        },
        async delete() {
          request.verb = "delete";
        },
      };
      return chain;
    },
  };
  return { client, requests, factory: () => client };
}

function graphFor(
  respond?: (path: string, body?: unknown) => unknown,
  options: { fetch?: FetchLike; maxUploadBytes?: number } = {},
) {
  const fake = fakeClient(respond);
  return {
    graph: createGraph({
      token: async () => "token",
      client: fake.factory,
      fetch: options.fetch,
      maxUploadBytes: options.maxUploadBytes,
    }),
    requests: fake.requests,
  };
}

/** One file of `bytes` bytes, as an upload source. */
function fileSource(bytes: number, name = "big.bin"): UploadSource {
  const buffer = Buffer.alloc(bytes, 0x61);
  return {
    name,
    contentType: "application/octet-stream",
    size: bytes,
    read: async (offset, length) => buffer.subarray(offset, offset + length),
  };
}

type ChunkAnswer = {
  status: number;
  body?: string;
  location?: string;
};

/**
 * A stub of Graph's upload session: it answers the session `POST` with a
 * pre-authenticated URL of its own, then answers the chunk `PUT`s in order from
 * `answers`, ending in the 201 that carries the attachment id.
 *
 * The token in that URL is the one a test wants kept out of the places it should
 * not go, so it is a recognisable string rather than a plausible one.
 */
function uploadSession(
  options: {
    uploadUrl?: string;
    expiresAt?: string;
    answers?: ChunkAnswer[];
  } = {},
) {
  const url =
    options.uploadUrl ??
    "https://outlook.office.com/api/v2.0/Users('u')/Messages('m')/AttachmentSessions('s')?authtoken=secret-session-token";
  const puts: {
    url: string;
    headers: Record<string, string>;
    body: Buffer;
  }[] = [];
  const respond = (path: string): unknown =>
    path.endsWith("/attachments/createUploadSession")
      ? {
          uploadUrl: url,
          expirationDateTime: options.expiresAt ?? "2999-01-01T00:00:00Z",
          nextExpectedRanges: ["0-"],
        }
      : { id: "draft-1", isDraft: true };
  const fetch: FetchLike = async (putUrl, init) => {
    puts.push({ url: putUrl, headers: init.headers, body: init.body });
    // Unscripted, the session behaves as Graph does: every range but the last is
    // acknowledged with the next one expected, and the range that ends the file
    // comes back 201 with the attachment's URL.
    const answer: ChunkAnswer =
      options.answers?.[puts.length - 1] ??
      ((): ChunkAnswer => {
        // exec()[0] is the whole match, so the ranges start at index 1: byte
        // ranges here are `start-end/total`, and this answers with `end + 1`.
        const [, , end, total] =
          /^bytes (\d+)-(\d+)\/(\d+)$/.exec(
            init.headers["Content-Range"] ?? "",
          ) ?? [];
        return Number(end) + 1 === Number(total)
          ? {
              status: 201,
              location:
                "https://outlook.office.com/api/v2.0/Users('u')/Messages('m')/Attachments('att%2D1')",
            }
          : {
              status: 200,
              body: JSON.stringify({
                nextExpectedRanges: [`${Number(end) + 1}`],
              }),
            };
      })();
    return {
      status: answer.status,
      header: (name) =>
        name.toLowerCase() === "location" ? (answer.location ?? null) : null,
      body: async () => answer.body ?? "",
    };
  };
  return { url, respond, fetch, puts };
}

test("list messages requests the documented fields, newest first", async () => {
  const { graph, requests } = graphFor();
  const page = await graph.listMessages({ unread: true, top: 2 });

  assert.deepEqual(page.items, ["message"]);
  assert.equal(page.nextCursor, undefined);
  assert.equal(requests.length, 1);
  const request = requests[0];
  assert.equal(request.path, "/me/mailFolders/inbox/messages");
  assert.deepEqual(request.select, messageFields.split(","));
  assert.equal(request.orderby, "receivedDateTime desc");
  assert.equal(request.top, 2);
  assert.equal(request.filter, "isRead eq false");
  assert.deepEqual(request.headers, []);
});

test("list messages omits the unread filter unless asked", async () => {
  const { graph, requests } = graphFor();
  await graph.listMessages();

  const request = requests[0];
  assert.equal(request.filter, undefined);
  assert.equal(request.top, defaultPageSize);
});

test("a page that names no size asks for the one every caller describes", async () => {
  const { graph, requests } = graphFor();

  await graph.listMessages();
  await graph.searchMessages({ query: "invoice" });
  await graph.listDrafts();
  // A folder list takes no size from a caller, so it asks for the most one page
  // may carry - the same maximum every list tool accepts.
  await graph.listFolders();

  assert.deepEqual(
    requests.map((request) => request.top),
    [defaultPageSize, defaultPageSize, defaultPageSize, maxPageSize],
  );
});

test("folders resolve to Graph paths and unsafe names are refused", async () => {
  assert.equal(messagesPath("inbox"), "/me/mailFolders/inbox/messages");
  assert.equal(messagesPath(), "/me/mailFolders/inbox/messages");
  assert.equal(messagesPath("all"), "/me/messages");
  assert.equal(
    messagesPath("Project Stuff"),
    "/me/mailFolders/Project%20Stuff/messages",
  );
  for (const folder of ["", "  ", "bad/name", "..%2F", "a?b", "x'y"])
    assert.throws(() => messagesPath(folder), /Invalid Outlook folder name/);

  const { graph, requests } = graphFor();
  await graph.listMessages({ folder: "Archive" });
  assert.equal(requests[0].path, "/me/mailFolders/Archive/messages");
  await assert.rejects(graph.listMessages({ folder: "bad/name" }));
  assert.equal(requests.length, 1);
});

test("get message asks for a plain text body and full recipients", async () => {
  const { graph, requests } = graphFor();
  assert.deepEqual(await graph.getMessage("a/b"), { id: "message" });

  const request = requests[0];
  assert.equal(request.path, "/me/messages/a%2Fb");
  assert.deepEqual(request.headers, [
    ["Prefer", 'outlook.body-content-type="text"'],
  ]);
  assert.deepEqual(request.select, fullMessageFields.split(","));
});

test("attachments follow the nextLink then fetch each id", async () => {
  const { graph, requests } = graphFor();
  const attachments = await graph.listAttachments("a/b");

  assert.deepEqual(attachments, [
    { id: "attachment", contentBytes: "ZmlsZQ==" },
    { id: "attachment", contentBytes: "ZmlsZQ==" },
  ]);
  assert.deepEqual(
    requests.map((request) => request.path),
    [
      "/me/messages/a%2Fb/attachments",
      "/next-attachments",
      "/me/messages/a%2Fb/attachments/one%2Ftwo",
      "/me/messages/a%2Fb/attachments/two",
    ],
  );
});

test("attachment metadata requests exclude content and avoid per-item body fetches", async () => {
  const { graph, requests } = graphFor();
  assert.deepEqual(await graph.listAttachments("a/b", true), [
    { id: "one/two" },
    { id: "two" },
  ]);
  assert.equal(requests.length, 2);
  for (const request of requests)
    assert.deepEqual(request.select, [
      "id",
      "name",
      "contentType",
      "size",
      "isInline",
    ]);
  await graph.getAttachment("a/b", "c/d", true);
  assert.deepEqual(requests[2].select, [
    "id",
    "name",
    "contentType",
    "size",
    "isInline",
  ]);
});

test("a single attachment encodes both ids", async () => {
  const { graph, requests } = graphFor();
  await graph.getAttachment("a/b", "c/d");

  assert.equal(requests[0].path, "/me/messages/a%2Fb/attachments/c%2Fd");
});

test("search escapes the phrase and requests eventual consistency", async () => {
  const { graph, requests } = graphFor();
  const page = await graph.searchMessages({ query: 'a"b\\c', top: 3 });

  assert.deepEqual(page.items, ["message"]);
  const request = requests[0];
  assert.equal(request.path, "/me/messages");
  assert.deepEqual(request.headers, [["ConsistencyLevel", "eventual"]]);
  assert.equal(request.search, '"a\\"b\\\\c"');
  assert.equal(request.top, 3);
  assert.deepEqual(request.select, messageFields.split(","));
});

test("a nextLink becomes an opaque cursor that replays the query", async () => {
  // The shape Graph really issues: an absolute https link, versioned, with an
  // opaque token that must reach it byte for byte.
  const cursor =
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skiptoken=HQE,AAAAIiVFAAA%3d&$d=2y";
  const { graph, requests } = graphFor((path) =>
    path === cursor
      ? { value: ["second"] }
      : { value: ["first"], "@odata.nextLink": cursor },
  );

  const first = await graph.listMessages({ top: 2 });
  assert.deepEqual(first.items, ["first"]);
  assert.equal(first.nextCursor, cursor);

  const second = await graph.listMessages({ folder: "all", top: 50, cursor });
  assert.deepEqual(second.items, ["second"]);
  assert.equal(second.nextCursor, undefined);
  assert.equal(requests.length, 2);
  const replay = requests[1];
  assert.equal(replay.path, cursor);
  assert.equal(replay.select, undefined);
  assert.equal(replay.orderby, undefined);
  assert.equal(replay.top, undefined);
  assert.equal(replay.filter, undefined);
  assert.equal(replay.search, undefined);
  assert.deepEqual(replay.headers, []);

  const searchCursor =
    "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=search";
  const searched = await graph.searchMessages({
    query: "ignored",
    top: 50,
    cursor: searchCursor,
  });
  assert.deepEqual(searched.items, ["first"]);
  assert.equal(requests[2].path, searchCursor);
  assert.deepEqual(requests[2].headers, []);
});

const savedDraft = () => ({ id: "draft-1", isDraft: true });

test("create_draft posts a new message, which Graph stores as a draft", async () => {
  const { graph, requests } = graphFor(savedDraft);

  const created = await graph.createDraft({
    to: ["ana@example.com"],
    cc: ["bo@example.com"],
    subject: "Hello",
    body: "<p>Hi</p>",
    bodyFormat: "html",
    categories: ["work"],
  });

  assert.deepEqual(created, { id: "draft-1", isDraft: true });
  assert.deepEqual(requests.length, 1);
  const request = requests[0];
  assert.equal(request.path, "/me/messages");
  assert.equal(request.verb, "post");
  assert.deepEqual(request.body, {
    subject: "Hello",
    body: { contentType: "html", content: "<p>Hi</p>" },
    toRecipients: [{ emailAddress: { address: "ana@example.com" } }],
    ccRecipients: [{ emailAddress: { address: "bo@example.com" } }],
    categories: ["work"],
  });
});

test("a draft request leaves out what was not given, so a patch clears nothing", async () => {
  const { graph, requests } = graphFor(savedDraft);

  await graph.createDraft({ to: ["ana@example.com"] });
  await graph.updateDraft("a/b", { subject: "Changed" });

  assert.deepEqual(requests[0].body, {
    toRecipients: [{ emailAddress: { address: "ana@example.com" } }],
  });
  assert.equal(requests[1].path, "/me/messages/a%2Fb");
  assert.equal(requests[1].verb, "patch");
  assert.deepEqual(requests[1].body, { subject: "Changed" });
});

test("a reply and a forward draft go to the create actions with the comment", async () => {
  const { graph, requests } = graphFor(savedDraft);

  await graph.createReplyDraft("a/b", "Thanks for this.");
  await graph.createForwardDraft("a/b", ["bo@example.com"], "Passing on.");

  assert.deepEqual(requests[0].path, "/me/messages/a%2Fb/createReply");
  assert.equal(requests[0].verb, "post");
  assert.deepEqual(requests[0].body, { comment: "Thanks for this." });
  assert.deepEqual(requests[1].path, "/me/messages/a%2Fb/createForward");
  assert.deepEqual(requests[1].body, {
    comment: "Passing on.",
    toRecipients: [{ emailAddress: { address: "bo@example.com" } }],
  });
});

test("list_drafts reads the drafts folder newest first", async () => {
  const { graph, requests } = graphFor();

  await graph.listDrafts({ top: 4 });

  const request = requests[0];
  assert.equal(request.path, "/me/mailFolders/drafts/messages");
  assert.deepEqual(request.select, draftFields.split(","));
  assert.equal(request.orderby, "receivedDateTime desc");
  assert.equal(request.top, 4);
});

test("list_drafts replays a cursor instead of rebuilding the query", async () => {
  const cursor =
    "https://graph.microsoft.com/v1.0/me/mailFolders/drafts/messages?$skiptoken=drafts";
  const { graph, requests } = graphFor((path) =>
    path === cursor
      ? { value: [{ id: "draft-2" }] }
      : { value: [{ id: "draft-1" }], "@odata.nextLink": cursor },
  );

  const first = await graph.listDrafts();
  assert.equal(first.nextCursor, cursor);
  const second = await graph.listDrafts({ top: 50, cursor });

  assert.deepEqual(second.items, [{ id: "draft-2" }]);
  assert.equal(requests[1].path, cursor);
  assert.equal(requests[1].select, undefined);
  assert.equal(requests[1].top, undefined);
});

test("a cursor that is not a Graph page link never becomes a request", async () => {
  // A cursor arrives from the model, and an absolute URL given to the SDK is
  // fetched as it stands, bearer token included. So nothing here may reach
  // Graph until the link is one Graph could have issued for these pages.
  const hostile = [
    "https://evil.example/v1.0/me/messages",
    "http://graph.microsoft.com/v1.0/me/messages",
    "https://graph.microsoft.com:8443/v1.0/me/messages",
    "https://graph.microsoft.com@evil.example/v1.0/me/messages",
    "https://evil.example/#https://graph.microsoft.com/v1.0/me/messages",
    "https://graph.microsoft.com/v1.0/users/anyone/messages",
    "https://graph.microsoft.com/v1.0/me/messages/../../tenantAdmin",
    "https://graph.microsoft.com\\evil.example/v1.0/me/messages",
    "/me/messages?$skiptoken=abc",
    "me/messages",
    "javascript:fetch('//evil.example')",
  ];

  for (const cursor of hostile) {
    const { graph, requests } = graphFor();
    const reads = [
      () => graph.listMessages({ cursor }),
      () => graph.searchMessages({ query: "invoice", cursor }),
      () => graph.listDrafts({ cursor }),
    ];
    for (const read of reads) {
      await assert.rejects(
        read(),
        /cursor must be the nextCursor from an earlier result/,
        `${cursor || "an empty cursor"} was accepted`,
      );
    }
    assert.deepEqual(requests, [], `${cursor} reached Graph`);
  }
});

test("delete_draft deletes that one message", async () => {
  const { graph, requests } = graphFor();

  await graph.deleteDraft("a/b");

  assert.equal(requests[0].path, "/me/messages/a%2Fb");
  assert.equal(requests[0].verb, "delete");
});

test("marking read patches isRead on each message named, once each", async () => {
  const { graph, requests } = graphFor();

  const outcome = await graph.setMessageRead(["a/b", "c", "a/b"], true);

  assert.deepEqual(
    requests.map((request) => [request.path, request.verb, request.body]),
    [
      ["/me/messages/a%2Fb", "patch", { isRead: true }],
      ["/me/messages/c", "patch", { isRead: true }],
    ],
    "a repeated id costs one request, and nothing but the flag is sent",
  );
  // The fake answers without an isRead, so the flag that was asked for is the
  // honest report: Graph's word when it gives one, the request otherwise.
  assert.deepEqual(outcome, {
    updated: [
      { id: "a/b", isRead: true },
      { id: "c", isRead: true },
    ],
    failed: [],
  });
});

test("the flag reported back is the one Graph returned, not the one asked for", async () => {
  const { graph } = graphFor(() => ({ id: "m", isRead: false }));

  const outcome = await graph.setMessageRead(["m"], true);

  assert.deepEqual(outcome.updated, [{ id: "m", isRead: false }]);
});

test("one message Graph refuses is reported beside the ones that moved", async () => {
  const long = "Resource not found. ".repeat(60);
  const { graph, requests } = graphFor((path) => {
    if (path.endsWith("/gone")) throw new Error(`\n ${long}`);
    return { id: "m", isRead: true };
  });

  const outcome = await graph.setMessageRead(["here", "gone"], true);

  // The refusal is one id's, so the request that succeeded is not undone and the
  // caller is not told to redo the whole batch.
  assert.equal(requests.length, 2);
  assert.deepEqual(outcome.updated, [{ id: "here", isRead: true }]);
  assert.deepEqual(
    outcome.failed.map(({ id }) => id),
    ["gone"],
  );
  // A Graph error is a JSON blob over several lines; a result carries one line of
  // it, however long the blob was.
  const [reason] = outcome.failed.map(({ reason }) => reason);
  assert.equal(reason?.includes("\n"), false);
  assert.equal(
    reason?.length,
    maxFlagReasonChars + 1,
    "capped, with the ellipsis saying where it was cut",
  );
});

/**
 * The shape the Graph SDK really throws: its error handler rebuilds every
 * rejection as a `GraphError`, keeping the message and filling `code` from
 * `error.name`, and losing the class. A fake that threw `AuthRequiredError`
 * straight would prove nothing about that rewrite, so this is built the way
 * `GraphErrorHandler.constructError` builds it.
 */
function asSdkThrew(cause: Error): unknown {
  const rewritten = new GraphError(-1, "", cause) as GraphError & {
    code: string | null;
  };
  rewritten.code = cause.name;
  return rewritten;
}

test("a sign-in failure survives the SDK rewriting it, and only that one does", async () => {
  const { graph } = graphFor((path) => {
    if (path.endsWith("/signed-out"))
      throw asSdkThrew(new AuthRequiredError(noSessionMessage));
    throw asSdkThrew(new Error("Item not found."));
  });

  await assert.rejects(
    () => graph.getMessage("signed-out"),
    (error: unknown) =>
      error instanceof AuthRequiredError && error.message === noSessionMessage,
    "the sign-in failure reaches the caller as a sign-in failure",
  );
  // Every other refusal keeps the shape Graph gave it: unwrapping the one class
  // the caller acts on must not quietly relabel a missing message.
  await assert.rejects(
    () => graph.getMessage("missing"),
    (error: unknown) =>
      error instanceof AuthRequiredError === false &&
      String(error).includes("Item not found."),
  );
});

test("a missing session fails a flag batch once, not once per id", async () => {
  const { graph } = graphFor((path) => {
    // The SDK's rewritten form, because that is what a real request rejects with
    // and the batch has to recognise it after `traced` hands it back.
    if (path.endsWith("/b"))
      throw asSdkThrew(new AuthRequiredError(noSessionMessage));
    return { id: "m", isRead: true };
  });

  // Every id fails for the same reason and the remediation is one sign-in, so it
  // is raised as it came rather than listed a hundred times as data.
  await assert.rejects(
    () => graph.setMessageRead(["a", "b"], false),
    (error: unknown) =>
      error instanceof AuthRequiredError && error.message === noSessionMessage,
  );
});

test("uploadAttachment posts a file under 3 MB as one attachment", async () => {
  // `storedAttachment` answers the POST the way Graph does, so a body that would
  // draw "Required property is missing." from the service fails here too.
  const session = uploadSession();
  const { graph, requests } = graphFor(
    (path, body) =>
      path.endsWith("/attachments") && body !== undefined
        ? storedAttachment(body)
        : session.respond(path),
    { fetch: session.fetch },
  );

  const outcome = await graph.uploadAttachment("a/b", {
    name: "note.txt",
    contentType: "text/plain",
    size: 5,
    read: async () => Buffer.from("hello"),
  });

  assert.equal(outcome.attachment.transport, "single_request");
  assert.equal(outcome.attachment.bytes, 5);
  // The draft read on the way to the write comes back with it, so the answer can
  // echo the message rather than the upload.
  assert.equal(outcome.draft.id, "draft-1");
  const request = requests.find((entry) => entry.verb === "post");
  assert.equal(request?.path, "/me/messages/a%2Fb/attachments");
  assert.deepEqual(request?.body, {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: "note.txt",
    contentType: "text/plain",
    contentBytes: "aGVsbG8=",
  });
  assert.equal(
    "contentBase64" in (request?.body as Record<string, unknown>),
    false,
    "the tool's argument name is not the wire field Graph reads",
  );
  assert.deepEqual(
    session.puts,
    [],
    "3 MB is not reached, so no session is opened",
  );
});

test("an attachment write reads its target first and refuses a message that is not a draft", async () => {
  // `POST …/attachments` and `DELETE …/attachments/{id}` both work on a message
  // that has already been sent, which is the one thing nothing here may rewrite.
  const sent = () => ({ id: "m", isDraft: false });
  const session = uploadSession();
  const { graph, requests } = graphFor(sent, { fetch: session.fetch });

  for (const write of [
    () => graph.uploadAttachment("a/b", fileSource(5)),
    () => graph.deleteAttachment("a/b", "attachment-1"),
  ])
    await assert.rejects(write, /is not a draft, so nothing was changed/);

  assert.deepEqual(
    requests.filter((entry) => entry.verb !== "get"),
    [],
    "a sent message was written to",
  );
  assert.deepEqual(session.puts, []);
});

test("a file that cannot fit, or has nothing in it, is refused before any request", async () => {
  const { graph, requests } = graphFor(savedDraft, {
    maxUploadBytes: 1024,
    fetch: uploadSession().fetch,
  });

  await assert.rejects(
    () => graph.uploadAttachment("a/b", fileSource(1025)),
    (error: unknown) => {
      assert.match(
        messageOf(error),
        /big\.bin is 1025 bytes and the upload limit here is 1024 bytes \(OUTLOOK_MCP_MAX_UPLOAD_BYTES/,
      );
      assert.match(messageOf(error), /Graph caps at 157286400/);
      return true;
    },
  );
  await assert.rejects(
    () => graph.uploadAttachment("a/b", fileSource(0)),
    /has no bytes to attach/,
  );
  assert.deepEqual(requests, [], "a write that must fail was not attempted");

  // The default ceiling is the mailbox's, not Graph's, and is said in the error
  // a caller gets when nobody set anything.
  const defaults = graphFor(savedDraft);
  await assert.rejects(
    () =>
      defaults.graph.uploadAttachment(
        "a/b",
        fileSource(defaultMaxUploadBytes + 1),
      ),
    new RegExp(`the upload limit here is ${defaultMaxUploadBytes} bytes`),
  );
});

test("a file over 3 MB goes up an upload session in ordered ranges", async () => {
  const session = uploadSession();
  const { graph, requests } = graphFor(session.respond, {
    fetch: session.fetch,
  });
  const size = singleRequestAttachmentBytes + 5_000_000;

  const outcome = await graph.uploadAttachment("a/b", fileSource(size));

  assert.equal(outcome.attachment.transport, "upload_session");
  assert.equal(
    outcome.attachment.id,
    "att-1",
    "read out of the Location header",
  );
  assert.equal(outcome.attachment.bytes, size);
  const opened = requests.find((entry) => entry.verb === "post");
  assert.equal(
    opened?.path,
    "/me/messages/a%2Fb/attachments/createUploadSession",
  );
  assert.deepEqual(opened?.body, {
    AttachmentItem: {
      attachmentType: "file",
      name: "big.bin",
      contentType: "application/octet-stream",
      size,
    },
  });

  // The chunk size is the one Graph asks for: under the 4 MB it names, and on
  // the 320 KiB alignment its upload sessions are built on.
  assert.ok(
    uploadChunkBytes < 4 * 1024 * 1024,
    "a chunk over the 4 MB Graph asks for",
  );
  assert.equal(
    uploadChunkBytes % (320 * 1024),
    0,
    "a chunk off the 320 KiB alignment",
  );

  // Contiguous ranges, in order, all of that size until the tail, with the last
  // byte of the file in the last of them.
  let expected = 0;
  for (const put of session.puts) {
    const range = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(
      put.headers["Content-Range"] ?? "",
    );
    assert.ok(
      range,
      `no Content-Range on a chunk: ${put.headers["Content-Range"]}`,
    );
    const [start, end, total] = [
      Number(range[1]),
      Number(range[2]),
      Number(range[3]),
    ];
    assert.equal(start, expected, "ranges must follow one another");
    assert.equal(total, size);
    assert.equal(end, start + put.body.byteLength - 1);
    if (end + 1 < size)
      assert.equal(
        put.body.byteLength,
        uploadChunkBytes,
        "a chunk that is not the module's own size",
      );
    expected = end + 1;
  }
  assert.equal(expected, size, "the upload stopped short of the last byte");

  // The session URL is pre-authenticated by Graph, so nothing of the user's own
  // token belongs on it - and the byte ranges are what they were asked for.
  for (const put of session.puts) {
    assert.equal(put.headers["Authorization"], undefined);
    assert.equal(put.headers["Content-Type"], "application/octet-stream");
    assert.equal(put.url, session.url);
  }
  assert.ok(session.puts.length > 1, "a file this size is more than one chunk");
});

test("an upload resumes where the session says, and stops rather than repeat itself", async () => {
  const forward = uploadSession({
    answers: [
      { status: 200, body: JSON.stringify({ nextExpectedRanges: ["7680"] }) },
    ],
  });
  const resumed = graphFor(forward.respond, { fetch: forward.fetch });
  await resumed.graph.uploadAttachment(
    "a/b",
    fileSource(singleRequestAttachmentBytes + 10),
  );
  assert.match(
    forward.puts[1].headers["Content-Range"] ?? "",
    /^bytes 7680-/,
    "the server's next range was ignored",
  );

  const back = uploadSession({
    answers: [
      { status: 200, body: JSON.stringify({ nextExpectedRanges: ["0"] }) },
    ],
  });
  const stuck = graphFor(back.respond, { fetch: back.fetch });
  await assert.rejects(
    () =>
      stuck.graph.uploadAttachment(
        "a/b",
        fileSource(singleRequestAttachmentBytes + 10),
      ),
    /asked for byte 0 of .*after accepting byte/,
  );
  assert.equal(
    back.puts.length,
    1,
    "a loop that would not advance is cut short",
  );
});

test("a refused chunk says why, without carrying the session credential with it", async () => {
  // Graph's error text quotes the URL it refused, so the credential is exactly
  // what comes back in the body of a failed chunk.
  const url =
    "https://outlook.office.com/api/v2.0/Users('u')/Messages('m')/AttachmentSessions('s')?authtoken=secret-session-token";
  const session = uploadSession({
    uploadUrl: url,
    answers: [
      {
        status: 413,
        body: `{"error":{"message":"Attachment size exceeds the allowable limit for ${url}"}}`,
      },
    ],
  });
  const { graph } = graphFor(session.respond, { fetch: session.fetch });
  await assert.rejects(
    () =>
      graph.uploadAttachment(
        "a/b",
        fileSource(singleRequestAttachmentBytes + 10),
      ),
    (error: unknown) => {
      assert.match(messageOf(error), /with HTTP 413/);
      assert.match(messageOf(error), /authtoken=\[redacted\]/);
      assert.equal(messageOf(error).includes("secret-session-token"), false);
      assert.match(messageOf(error), /Nothing was attached/);
      return true;
    },
  );
});

test("an upload session that is expired, or points somewhere unexpected, sends no bytes", async () => {
  const expired = uploadSession({ expiresAt: "2000-01-01T00:00:00Z" });
  const past = graphFor(expired.respond, { fetch: expired.fetch });
  await assert.rejects(
    () =>
      past.graph.uploadAttachment(
        "a/b",
        fileSource(singleRequestAttachmentBytes + 10),
      ),
    /expired with .* still to send/,
  );
  assert.deepEqual(expired.puts, []);

  for (const uploadUrl of [
    "http://outlook.office.com/api/AttachmentSessions('s')?authtoken=t",
    "https://outlook.office.com:8443/api/AttachmentSessions('s')?authtoken=t",
    "https://someone@outlook.office.com/api/AttachmentSessions('s')",
    "not a url",
    "",
  ]) {
    const session = uploadSession({ uploadUrl });
    const { graph } = graphFor(session.respond, { fetch: session.fetch });
    await assert.rejects(
      () =>
        graph.uploadAttachment(
          "a/b",
          fileSource(singleRequestAttachmentBytes + 10),
        ),
      /will not send bytes to/,
      `${uploadUrl} was accepted`,
    );
    assert.deepEqual(session.puts, [], `${uploadUrl} was sent to`);
  }
});

test("the session URL never reaches a log line, not even at debug", async () => {
  const capture = captureLogs({ level: "debug" });
  const session = uploadSession();
  const { graph } = graphFor(session.respond, { fetch: session.fetch });
  await graph.uploadAttachment(
    "a/b",
    fileSource(singleRequestAttachmentBytes + 10),
  );
  capture.restore();
  assert.equal(capture.text().includes("secret-session-token"), false);
  assert.equal(capture.text().includes("AttachmentSessions"), false);
  assert.match(capture.text(), /attachment chunk uploaded/);
});

test("delete_attachment deletes only the one attachment of a draft", async () => {
  const { graph, requests } = graphFor(savedDraft);

  const draft = await graph.deleteAttachment("a/b", "att/1");

  assert.equal(draft.id, "draft-1");
  assert.equal(requests[1].path, "/me/messages/a%2Fb/attachments/att%2F1");
  assert.equal(requests[1].verb, "delete");
});

test("every Graph call leaves one debug line, and no line at the default level", async () => {
  const { graph } = graphFor();

  // A page of mail that arrived is not news: the caller writes the line that
  // says what the user asked for. At the default level this module is silent.
  const quiet = captureLogs({ level: "info" });
  try {
    await graph.listMessages({ top: 2 });
    assert.deepEqual(quiet.lines, []);
  } finally {
    quiet.restore();
  }

  const traced = captureLogs({ level: "debug" });
  try {
    await graph.listMessages({ top: 2 });
    const lines = traced.matching(/graph: graph request/);
    assert.equal(lines.length, 1);
    assert.match(
      lines[0] ?? "",
      /debug \d+ graph: graph request method=GET path=\/me\/mailFolders\/inbox\/messages ms=\d+ items=1/,
    );

    await graph.getMessage("AAMkAGI2TGV0MDE2OTY4NTUwMDAw");
    assert.match(traced.lines.at(-1) ?? "", /path=\/me\/messages\/AAMkAGI2/);

    // A cursor is a Graph URL whose query is the paging token. The path says
    // which page; the token is a credential that repeats on every request.
    await graph.listMessages({
      cursor:
        "https://graph.microsoft.com/v1.0/me/messages?%24skiptoken=TokenForNextPage",
    });
    assert.match(traced.lines.at(-1) ?? "", /path=\/v1\.0\/me\/messages\?…/);
    assert.ok(!traced.text().includes("TokenForNextPage"), traced.text());
  } finally {
    traced.restore();
  }
});

test("a failed Graph call is traced with its reason, and still throws", async () => {
  const { graph } = graphFor(() => {
    throw new Error("AADSTS700082: the token expired");
  });
  const traced = captureLogs({ level: "debug" });
  try {
    await assert.rejects(
      () => graph.getMessage("message-1"),
      /AADSTS700082/,
      "the trace does not swallow the failure",
    );
    assert.match(
      traced.text(),
      /graph request failed method=GET path=\/me\/messages\/message-1 ms=\d+ reason="AADSTS700082: the token expired"/,
    );
  } finally {
    traced.restore();
  }
});
