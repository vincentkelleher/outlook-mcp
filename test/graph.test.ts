import assert from "node:assert/strict";
import { test } from "node:test";
import {
  createGraph,
  defaultPageSize,
  maxAttachmentBytes,
  maxFlagReasonChars,
  maxPageSize,
  messagesPath,
} from "../src/core/graph.ts";
import type { GraphClient, GraphRequest } from "../src/core/graph.ts";
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

function responder(path: string): unknown {
  if (path === "/next-attachments") return { value: [{ id: "two" }] };
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

function fakeClient(respond: (path: string) => unknown = responder) {
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
          return respond(path) as T;
        },
        async patch<T>(body: unknown) {
          request.verb = "patch";
          request.body = body;
          return respond(path) as T;
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

function graphFor(respond?: (path: string) => unknown) {
  const fake = fakeClient(respond);
  return {
    graph: createGraph({ token: async () => "token", client: fake.factory }),
    requests: fake.requests,
  };
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

test("add_attachment posts a file attachment to the draft", async () => {
  const { graph, requests } = graphFor(savedDraft);

  await graph.addAttachment("a/b", {
    name: "note.txt",
    contentType: "text/plain",
    contentBase64: Buffer.from("hello").toString("base64"),
  });

  const request = requests[0];
  assert.equal(request.path, "/me/messages/a%2Fb/attachments");
  assert.equal(request.verb, "post");
  assert.deepEqual(request.body, {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: "note.txt",
    contentType: "text/plain",
    contentBase64: Buffer.from("hello").toString("base64"),
  });
});

test("an attachment at or over 3 MB is refused before any request", async () => {
  const { graph, requests } = graphFor(savedDraft);
  const attachment = (bytes: number) => ({
    name: "big.bin",
    contentType: "application/octet-stream",
    contentBase64: Buffer.alloc(bytes, 0x61).toString("base64"),
  });

  assert.throws(
    () => graph.addAttachment("a/b", attachment(maxAttachmentBytes)),
    (error: unknown) => {
      assert.match(
        messageOf(error),
        /big\.bin is 3145728 bytes and a single-request attachment must stay under 3145728 bytes \(3 MB\)/,
      );
      assert.match(messageOf(error), /upload session/);
      return true;
    },
  );
  assert.deepEqual(requests, [], "an upload that must fail was not attempted");

  // Just under the ceiling still goes up.
  await graph.addAttachment("a/b", attachment(maxAttachmentBytes - 1));
  assert.equal(requests.length, 1);
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
