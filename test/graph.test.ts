import assert from "node:assert/strict";
import { test } from "node:test";
import {
  createGraph,
  defaultPageSize,
  maxAttachmentBytes,
  maxPageSize,
  messagesPath,
} from "../src/core/graph.ts";
import type { GraphClient, GraphRequest } from "../src/core/graph.ts";

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
