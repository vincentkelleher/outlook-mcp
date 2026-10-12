import assert from "node:assert/strict";
import { test } from "node:test";
import { createGraph, messagesPath } from "../src/core/graph.ts";
import type { GraphClient, GraphRequest } from "../src/core/graph.ts";

const messageFields =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments";
const fullMessageFields = `${messageFields},toRecipients,ccRecipients,body`;

type Recorded = {
  path: string;
  headers: string[][];
  select?: string[];
  orderby?: string;
  top?: number;
  filter?: string;
  search?: string;
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
      const request: Recorded = { path, headers: [] };
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
  assert.equal(request.top, 10);
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
  const cursor = "/me/mailFolders/inbox/messages?$skiptoken=abc";
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

  const searchCursor = "/me/messages?$skiptoken=search";
  const searched = await graph.searchMessages({
    query: "ignored",
    top: 50,
    cursor: searchCursor,
  });
  assert.deepEqual(searched.items, ["first"]);
  assert.equal(requests[2].path, searchCursor);
  assert.deepEqual(requests[2].headers, []);
});
