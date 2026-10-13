import assert from "node:assert/strict";
import { test } from "node:test";
import { main } from "../src/cli.ts";
import type { CliDeps } from "../src/cli.ts";
import type { Graph } from "../src/core/graph.ts";

type Calls = {
  list: unknown[];
  search: unknown[];
  get: string[];
  attachments: string[];
};

function fakeGraph(): { graph: Graph; calls: Calls } {
  const calls: Calls = { list: [], search: [], get: [], attachments: [] };
  const graph: Graph = {
    async listMessages(options = {}) {
      calls.list.push(options);
      return { items: [{ id: "listed" }] };
    },
    async searchMessages(options) {
      calls.search.push(options);
      return { items: [{ id: "found" }] };
    },
    async getMessage(id) {
      calls.get.push(id);
      return { id };
    },
    async listFolders() {
      return { items: [{ id: "folder-1", displayName: "Inbox" }] };
    },
    async listAttachments(messageId) {
      calls.attachments.push(messageId);
      return [{ id: "attachment" }];
    },
    async getAttachment(messageId, attachmentId) {
      return { messageId, attachmentId };
    },
  };
  return { graph, calls };
}

async function runCli(args: string[], deps: CliDeps = {}) {
  const stdout: string[] = [];
  const stderr: string[] = [];
  const log = console.log;
  const error = console.error;
  console.log = (...lines: unknown[]) => stdout.push(lines.join(" "));
  console.error = (...lines: unknown[]) => stderr.push(lines.join(" "));
  let failure: unknown;
  try {
    await main(["node", "outlook", ...args], deps);
  } catch (thrown) {
    failure = thrown;
  } finally {
    console.log = log;
    console.error = error;
  }
  return { stdout: stdout.join("\n"), stderr: stderr.join("\n"), failure };
}

function message(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

test("list sends the unread flag and limit, and prints indented JSON", async () => {
  const { graph, calls } = fakeGraph();

  const first = await runCli(["emails", "list"], { graph });
  assert.equal(first.failure, undefined);
  assert.equal(first.stdout, JSON.stringify([{ id: "listed" }], null, 2));
  assert.deepEqual(JSON.parse(first.stdout), [{ id: "listed" }]);
  assert.equal(first.stderr, "");

  const unread = await runCli(["emails", "list", "--unread", "--limit", "3"], {
    graph,
  });
  assert.deepEqual(unread.stdout, JSON.stringify([{ id: "listed" }], null, 2));
  assert.deepEqual(calls.list, [
    { unread: undefined, top: 10 },
    { unread: true, top: 3 },
  ]);
});

test("search sends the query and its own larger maximum", async () => {
  const { graph, calls } = fakeGraph();

  const result = await runCli(
    ["emails", "search", "invoice", "--limit", "1000"],
    {
      graph,
    },
  );

  assert.equal(result.failure, undefined);
  assert.deepEqual(JSON.parse(result.stdout), [{ id: "found" }]);
  assert.deepEqual(calls.search, [{ query: "invoice", top: 1000 }]);
});

test("get and attachments pass ids straight through", async () => {
  const { graph, calls } = fakeGraph();

  const got = await runCli(["emails", "get", "a/b"], { graph });
  const attached = await runCli(["emails", "attachments", "m1"], { graph });

  assert.deepEqual(JSON.parse(got.stdout), { id: "a/b" });
  assert.deepEqual(JSON.parse(attached.stdout), [{ id: "attachment" }]);
  assert.deepEqual(calls.get, ["a/b"]);
  assert.deepEqual(calls.attachments, ["m1"]);
});

test("limits are validated against each command maximum", async () => {
  const { graph, calls } = fakeGraph();
  const cases: Array<{ args: string[]; maximum: number }> = [
    { args: ["emails", "list", "--limit", "0"], maximum: 100 },
    { args: ["emails", "list", "--limit", "101"], maximum: 100 },
    { args: ["emails", "list", "--limit", "1.5"], maximum: 100 },
    { args: ["emails", "list", "--limit", "nope"], maximum: 100 },
    { args: ["emails", "search", "q", "--limit", "0"], maximum: 1000 },
    { args: ["emails", "search", "q", "--limit", "1001"], maximum: 1000 },
  ];

  for (const { args, maximum } of cases) {
    const result = await runCli(args, { graph });
    assert.equal(
      message(result.failure),
      `--limit must be an integer from 1 to ${maximum}.`,
    );
    assert.equal(result.stdout, "");
  }
  assert.deepEqual(calls.list, []);
  assert.deepEqual(calls.search, []);
});
