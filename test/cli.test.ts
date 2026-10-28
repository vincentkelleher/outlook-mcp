import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { test } from "node:test";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { main } from "../src/cli.ts";
import { loadConfig, scopesFor } from "../src/core/config.ts";
import { writeLoginState } from "../src/core/device-login.ts";
import type { CliDeps } from "../src/cli.ts";
import type { Graph } from "../src/core/graph.ts";
import {
  defaultPageSize,
  maxFlaggedMessages,
  maxPageSize,
  maxSearchPageSize,
} from "../src/core/graph.ts";

type Calls = {
  list: unknown[];
  search: unknown[];
  get: string[];
  attachments: string[];
  flags: [string[], boolean][];
};

function fakeGraph(): { graph: Graph; calls: Calls } {
  const calls: Calls = {
    list: [],
    search: [],
    get: [],
    attachments: [],
    flags: [],
  };
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
    // The CLI has no draft commands, so any call here would be a surface the
    // design says does not exist.
    async createDraft() {
      throw new Error("the CLI must not create drafts");
    },
    async createReplyDraft() {
      throw new Error("the CLI must not create reply drafts");
    },
    async createForwardDraft() {
      throw new Error("the CLI must not create forward drafts");
    },
    async updateDraft() {
      throw new Error("the CLI must not update drafts");
    },
    async listDrafts() {
      throw new Error("the CLI must not list drafts");
    },
    async deleteDraft() {
      throw new Error("the CLI must not delete drafts");
    },
    async uploadAttachment() {
      throw new Error("the CLI must not upload attachments");
    },
    async deleteAttachment() {
      throw new Error("the CLI must not remove attachments");
    },
    // Flagging is the one write the CLI does have, alongside the draft tools.
    async setMessageRead(ids, read) {
      calls.flags.push([ids, read]);
      return {
        updated: ids.map((id) => ({ id, isRead: read })),
        failed: [],
      };
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
    // No --limit, so the CLI's default is the page size Graph defaults to too.
    { unread: undefined, top: defaultPageSize },
    { unread: true, top: 3 },
  ]);
});

test("search sends the query and its own larger maximum", async () => {
  const { graph, calls } = fakeGraph();

  const result = await runCli(
    ["emails", "search", "invoice", "--limit", `${maxSearchPageSize}`],
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

test("read and unread flag the ids given, and a batch over the cap costs nothing", async () => {
  const { graph, calls } = fakeGraph();

  const read = await runCli(["emails", "read", "a/b", "c"], { graph });
  const unread = await runCli(["emails", "unread", "a/b"], { graph });

  assert.equal(read.failure, undefined);
  assert.equal(unread.failure, undefined);
  assert.deepEqual(JSON.parse(read.stdout), {
    updated: [
      { id: "a/b", isRead: true },
      { id: "c", isRead: true },
    ],
    failed: [],
  });
  assert.deepEqual(calls.flags, [
    [["a/b", "c"], true],
    [["a/b"], false],
  ]);

  // The cap is the tool's own, and the CLI holds it too: a shell expansion that
  // sweeps in two hundred ids asks for two hundred PATCHes unless something here
  // says no, and it says no before it asks for anything.
  const tooMany = await runCli(
    [
      "emails",
      "read",
      ...Array.from(
        { length: maxFlaggedMessages + 1 },
        (_unused, index) => `m${index}`,
      ),
    ],
    { graph },
  );
  assert.match(
    message(tooMany.failure),
    new RegExp(
      `At most ${maxFlaggedMessages} ids per command, not ${maxFlaggedMessages + 1}`,
    ),
  );
  assert.equal(calls.flags.length, 2);
});

test("login --background reports an existing session instead of starting a second one", async () => {
  const { graph, calls } = fakeGraph();

  const result = await runCli(["login", "--background", "--json"], {
    graph,
    authStatus: async () => ({
      signedIn: true,
      account: "someone@example.com",
      authority: "https://login.microsoftonline.com/consumers/",
      cachePath: "/home/someone/.outlook-mcp/msal-cache.json",
    }),
  });

  assert.equal(result.failure, undefined);
  const printed = JSON.parse(result.stdout);
  assert.equal(printed.state, "authorized");
  assert.equal(printed.account, "someone@example.com");
  assert.match(printed.instructions, /Already signed in/);
  assert.deepEqual(calls.list, []);
});

test("auth status reports the session, the mode, and a code another process started", async () => {
  const home = mkdtempSync(join(tmpdir(), "outlook-mcp-cli-"));
  const saved = {
    OUTLOOK_MCP_HOME: process.env.OUTLOOK_MCP_HOME,
    OUTLOOK_MCP_STATE_DIR: process.env.OUTLOOK_MCP_STATE_DIR,
  };
  process.env.OUTLOOK_MCP_HOME = home;
  process.env.OUTLOOK_MCP_STATE_DIR = join(home, "state");
  const config = loadConfig({ ...process.env });
  try {
    const { graph, calls } = fakeGraph();
    const authStatus = async () => ({
      signedIn: false,
      account: "someone@example.com",
      reason: "needsConsent" as const,
      authority: "https://login.microsoftonline.com/consumers/",
      cachePath: join(home, "msal-cache.json"),
    });

    const bare = await runCli(["auth", "status"], { graph, authStatus });
    assert.equal(bare.failure, undefined);
    const printed = JSON.parse(bare.stdout);
    assert.equal(printed.signedIn, false);
    assert.equal(printed.reason, "needsConsent");
    // The mode and the permissions behind it, so a missing write tool explains
    // itself here instead of being guessed at — the tool says the same.
    assert.equal(printed.readonly, false);
    assert.deepEqual(printed.scopes, scopesFor(false));

    // A code belongs to the flow and not to whoever started it, so this is where
    // a caller with no MCP finds the one that is live — which is the whole
    // polling loop, from a shell.
    writeLoginState(config, {
      state: "pending",
      pid: 4242,
      verificationUri: "https://microsoft.com/devicelogin",
      userCode: "ABC-XYZ123",
      expiresAt: new Date(Date.now() + 600_000).toISOString(),
    });
    const pending = await runCli(["auth", "status"], { graph, authStatus });
    const carried = JSON.parse(pending.stdout);
    assert.equal(carried.signInPending, true);
    assert.equal(carried.verificationUri, "https://microsoft.com/devicelogin");
    assert.equal(carried.userCode, "ABC-XYZ123");

    // A report, not a gate: it asked nothing of the mailbox, and it exits 0
    // whoever reads `.signedIn` from the JSON.
    assert.deepEqual(calls.list, []);
  } finally {
    for (const [name, value] of Object.entries(saved)) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
    rmSync(home, { recursive: true, force: true });
  }
});

test("limits are validated against each command maximum", async () => {
  const { graph, calls } = fakeGraph();
  const cases: Array<{ args: string[]; maximum: number }> = [
    { args: ["emails", "list", "--limit", "0"], maximum: maxPageSize },
    {
      args: ["emails", "list", "--limit", `${maxPageSize + 1}`],
      maximum: maxPageSize,
    },
    { args: ["emails", "list", "--limit", "1.5"], maximum: maxPageSize },
    { args: ["emails", "list", "--limit", "nope"], maximum: maxPageSize },
    {
      args: ["emails", "search", "q", "--limit", "0"],
      maximum: maxSearchPageSize,
    },
    {
      args: ["emails", "search", "q", "--limit", `${maxSearchPageSize + 1}`],
      maximum: maxSearchPageSize,
    },
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
