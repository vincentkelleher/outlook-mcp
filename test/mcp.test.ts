import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/sdk/client";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../src/mcp.ts";
import { maxInlineBytes, tools } from "../src/tools.ts";
import type { ToolDeps } from "../src/tools.ts";
import { loadConfig } from "../src/core/config.ts";
import type { Config } from "../src/core/config.ts";
import { writeLoginState } from "../src/core/device-login.ts";
import { AuthRequiredError, noSessionMessage } from "../src/core/errors.ts";
import type { Graph, Page } from "../src/core/graph.ts";

const clientId = "00000000-0000-0000-0000-000000000000";
const readAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: true,
};

type TestContext = { after: (fn: () => void) => void };

function testConfig(t: TestContext): Config {
  const home = mkdtempSync(join(tmpdir(), "outlook-mcp-test-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  return loadConfig({
    HOME: home,
    OUTLOOK_OAUTH_CLIENT_ID: clientId,
    OUTLOOK_MCP_ATTACHMENT_DIR: join(home, "attachments"),
    OUTLOOK_MCP_STATE_DIR: join(home, "state"),
  });
}

type Calls = {
  list: unknown[];
  search: unknown[];
  get: string[];
  folders: number;
  attachments: string[];
  download: [string, string][];
  loginStarts: number;
  logouts: string[];
};

type Fake = { deps: ToolDeps; calls: Calls };

function fakeDeps(
  config: Config,
  options: { signedIn?: boolean; session?: boolean } = {},
): Fake {
  const calls: Calls = {
    list: [],
    search: [],
    get: [],
    folders: 0,
    attachments: [],
    download: [],
    loginStarts: 0,
    logouts: [],
  };
  // Without a session every Graph call fails the way the real token provider
  // does, which is what the auth remediation has to describe.
  const guard = () => {
    if (options.session === false)
      throw new AuthRequiredError(noSessionMessage);
  };
  const graph: Graph = {
    async listMessages(options = {}) {
      guard();
      calls.list.push(options);
      return {
        items: [{ id: "message-1", subject: "Hello" }],
        nextCursor: "https://graph/next?skiptoken=1",
      } as Page<unknown>;
    },
    async searchMessages(options) {
      guard();
      calls.search.push(options);
      return { items: [{ id: "found", subject: options.query }] };
    },
    async getMessage(id) {
      guard();
      calls.get.push(id);
      return { id, body: { content: "the whole body" } };
    },
    async listFolders() {
      guard();
      calls.folders += 1;
      return { items: [{ id: "folder-1", displayName: "Inbox" }] };
    },
    async listAttachments(messageId) {
      guard();
      calls.attachments.push(messageId);
      return [
        {
          id: "attachment-1",
          name: "note.txt",
          contentType: "text/plain",
          size: 5,
          isInline: false,
          contentBytes: "aGVsbG8=",
        },
      ];
    },
    async getAttachment(messageId, attachmentId) {
      guard();
      calls.download.push([messageId, attachmentId]);
      return {
        id: attachmentId,
        // A Graph file name is user data, including this traversal attempt.
        name: "../escape/note.txt",
        contentType: "text/plain",
        contentBytes:
          attachmentId === "oversized"
            ? Buffer.alloc(maxInlineBytes + 1, 0x61).toString("base64")
            : "aGVsbG8=",
      };
    },
  };
  const deps: ToolDeps = {
    graph,
    config,
    authStatus: async () => ({
      signedIn: options.signedIn ?? true,
      account: options.signedIn === false ? undefined : "someone@example.com",
      authority: config.authority,
      cachePath: config.cachePath,
    }),
    startLogin: () => {
      calls.loginStarts += 1;
      writeLoginState(config, {
        state: "pending",
        verificationUri: "https://microsoft.com/devicelogin",
        userCode: "ABC-XYZ",
        expiresAt: "2030-01-01T00:00:00.000Z",
      });
    },
    logout: (cachePath) => {
      calls.logouts.push(cachePath);
    },
  };
  return { deps, calls };
}

async function connect(t: TestContext, deps: ToolDeps) {
  const server = createServer(deps);
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "0.0.0" });
  await Promise.all([
    server.connect(serverTransport),
    client.connect(clientTransport),
  ]);
  t.after(() => void client.close());
  return client;
}

type Result = {
  content: { type: string; text?: string }[];
  isError?: boolean;
  structuredContent?: unknown;
};

function textOf(result: Result): string {
  return result.content.find((block) => block.type === "text")?.text ?? "";
}

function resourceTextOf(result: { contents: unknown[] }): string {
  return (result.contents[0] as { text?: string }).text ?? "";
}

async function call(
  client: Client,
  name: string,
  args: Record<string, unknown> = {},
): Promise<Result> {
  return (await client.callTool({
    name,
    arguments: args,
  })) as unknown as Result;
}

test("the handshake names the server and states what it cannot do", async (t) => {
  const client = await connect(t, fakeDeps(testConfig(t)).deps);

  assert.deepEqual(client.getServerVersion(), {
    name: "outlook",
    version: JSON.parse(
      readFileSync(new URL("../package.json", import.meta.url), "utf8"),
    ).version,
  });
  const instructions = client.getInstructions() ?? "";
  assert.match(
    instructions,
    /Nothing on this server can send, reply or forward/,
  );
  assert.match(instructions, /login/);
  assert.match(instructions, /untrusted/);
});

test("tools/list carries every tool, its annotations and the limits it enforces", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);
  const listed = await client.listTools();

  assert.deepEqual(listed.tools.map((tool) => tool.name).sort(), [
    "auth_status",
    "download_attachment",
    "get_email",
    "list_attachments",
    "list_emails",
    "login",
    "logout",
    "search_emails",
  ]);
  assert.deepEqual(
    listed.tools.map((tool) => tool.name).sort(),
    tools.map((tool) => tool.name).sort(),
  );
  assert.equal(
    listed.tools.find((tool) => tool.name === "list_emails")?.title,
    "List emails",
  );

  const signed = [
    "list_emails",
    "search_emails",
    "get_email",
    "list_attachments",
    "download_attachment",
    "auth_status",
  ];
  for (const name of signed) {
    assert.deepEqual(
      listed.tools.find((tool) => tool.name === name)?.annotations,
      readAnnotations,
      `${name} annotations`,
    );
  }
  const logoutTool = listed.tools.find((tool) => tool.name === "logout");
  assert.equal(logoutTool?.annotations?.destructiveHint, true);
  assert.equal(logoutTool?.annotations?.readOnlyHint, false);

  // The advertised bounds and the enforced bounds are the same schema, so a
  // client reading tools/list cannot be told something the server disagrees with.
  const bounds = (name: string) => {
    const schema = listed.tools.find((tool) => tool.name === name)
      ?.inputSchema as
      | { properties?: { limit?: { minimum?: number; maximum?: number } } }
      | undefined;
    return {
      minimum: schema?.properties?.limit?.minimum,
      maximum: schema?.properties?.limit?.maximum,
    };
  };
  assert.deepEqual(bounds("list_emails"), { minimum: 1, maximum: 100 });
  assert.deepEqual(bounds("search_emails"), { minimum: 1, maximum: 1000 });
  for (const [name, maximum, required] of [
    ["list_emails", 100, {}],
    ["search_emails", 1000, { query: "invoice" }],
  ] as const) {
    const at = (limit: number) => call(client, name, { ...required, limit });
    assert.equal(
      (await at(maximum)).isError,
      undefined,
      `${name} at ${maximum}`,
    );
    assert.equal((await at(maximum + 1)).isError, true, `${name} over maximum`);
    assert.equal((await at(0)).isError, true, `${name} under minimum`);
  }
  assert.equal(calls.list.length, 1);
  assert.equal(calls.search.length, 1);
});

test("list_emails, search_emails and get_email answer with compact JSON", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const listed = await call(client, "list_emails", {
    folder: "Archive",
    unread: true,
    limit: 3,
  });
  assert.equal(listed.isError, undefined);
  assert.deepEqual(calls.list, [
    { folder: "Archive", unread: true, top: 3, cursor: undefined },
  ]);
  const payload = JSON.parse(textOf(listed));
  assert.deepEqual(payload.items, [{ id: "message-1", subject: "Hello" }]);
  assert.equal(payload.nextCursor, "https://graph/next?skiptoken=1");
  // No indentation: a page of mail must fit a client's text budget.
  assert.equal(textOf(listed), JSON.stringify(payload));
  assert.deepEqual(listed.structuredContent, payload);

  const found = await call(client, "search_emails", { query: "from:ana" });
  assert.deepEqual(JSON.parse(textOf(found)).items, [
    { id: "found", subject: "from:ana" },
  ]);
  assert.deepEqual(calls.search, [
    { query: "from:ana", top: undefined, cursor: undefined },
  ]);

  const got = await call(client, "get_email", { id: "message-1" });
  assert.deepEqual(JSON.parse(textOf(got)), {
    id: "message-1",
    body: { content: "the whole body" },
  });
  assert.deepEqual(calls.get, ["message-1"]);
});

test("a missing session answers with the next step, not a stack trace", async (t) => {
  const client = await connect(
    t,
    fakeDeps(testConfig(t), { session: false }).deps,
  );

  const result = await call(client, "list_emails");
  assert.equal(result.isError, true);
  const text = textOf(result);
  assert.match(text, /No Outlook session\./);
  assert.match(text, /Call the login tool/);
  assert.match(text, /verification URL and user code/);
  assert.match(text, /retry this call\./);
  assert.equal(text.includes("\n"), false, "error text is one line");
  assert.equal(text.includes("AuthRequiredError"), false);
  assert.equal(text.includes("at "), false);
});

test("arguments that break the schema are rejected before Graph is called", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const wrongType = await call(client, "get_email", { id: 7 });
  assert.equal(wrongType.isError, true);
  assert.match(textOf(wrongType), /Invalid arguments for tool get_email/);

  const unknownKey = await call(client, "list_emails", { folder: 1 });
  assert.equal(unknownKey.isError, true);
  const surprise = await call(client, "auth_status", { surprise: true });
  assert.equal(surprise.isError, true);
  assert.deepEqual(calls.list, []);
  assert.deepEqual(calls.get, []);
});

test("list_attachments returns metadata and never file contents", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const result = await call(client, "list_attachments", { id: "message-1" });
  const valid = await call(client, "list_attachments", {
    messageId: "message-1",
  });

  assert.equal(result.isError, true, "messageId is required");
  assert.deepEqual(JSON.parse(textOf(valid)).items, [
    {
      id: "attachment-1",
      name: "note.txt",
      contentType: "text/plain",
      size: 5,
      isInline: false,
    },
  ]);
  assert.equal(textOf(valid).includes("contentBytes"), false);
  assert.equal(textOf(valid).includes("aGVsbG8="), false);
  assert.deepEqual(calls.attachments, ["message-1"]);
});

test("download_attachment writes inside the allowed directory only", async (t) => {
  const config = testConfig(t);
  const { deps, calls } = fakeDeps(config);
  const client = await connect(t, deps);

  const written = await call(client, "download_attachment", {
    messageId: "message-1",
    attachmentId: "attachment-1",
  });
  assert.equal(written.isError, undefined);
  const payload = JSON.parse(textOf(written));
  // The Graph name carries path traversal; the stored name may not.
  assert.equal(payload.name, "note.txt");
  assert.equal(payload.path, join(config.attachmentDir, "note.txt"));
  assert.equal(payload.bytes, 5);
  assert.equal(readFileSync(payload.path, "utf8"), "hello");
  assert.equal(payload.contentBase64, undefined, "inline content is opt-in");

  const inline = await call(client, "download_attachment", {
    messageId: "message-1",
    attachmentId: "attachment-1",
    dir: "case-7",
    encoding: "base64",
  });
  const inlinePayload = JSON.parse(textOf(inline));
  assert.equal(inlinePayload.contentBase64, "aGVsbG8=");
  assert.equal(
    inlinePayload.path,
    join(config.attachmentDir, "case-7", "note.txt"),
  );

  const tooBig = await call(client, "download_attachment", {
    messageId: "message-1",
    attachmentId: "oversized",
    encoding: "base64",
  });
  const bigPayload = JSON.parse(textOf(tooBig));
  assert.equal(bigPayload.bytes, maxInlineBytes + 1);
  assert.equal(bigPayload.contentBase64, undefined, "over the inline cap");
  assert.match(bigPayload.note, /over the .* inline cap/);
  assert.deepEqual(calls.download, [
    ["message-1", "attachment-1"],
    ["message-1", "attachment-1"],
    ["message-1", "oversized"],
  ]);

  for (const dir of ["../escape", "/etc", "nested/../../outside"]) {
    const refused = await call(client, "download_attachment", {
      messageId: "message-1",
      attachmentId: "attachment-1",
      dir,
    });
    assert.equal(refused.isError, true, `${dir} was accepted`);
    assert.match(
      textOf(refused),
      /dir must be inside the attachment directory/,
    );
    assert.ok(textOf(refused).includes(config.attachmentDir));
  }
  assert.equal(
    existsSync(join(config.attachmentDir, "..", "escape")),
    false,
    "a rejected dir wrote nothing",
  );
});

test("auth_status reports the session and the last sign-in failure", async (t) => {
  const config = testConfig(t);
  const { deps } = fakeDeps(config);
  const client = await connect(t, deps);

  assert.deepEqual(JSON.parse(textOf(await call(client, "auth_status"))), {
    signedIn: true,
    account: "someone@example.com",
    authority: config.authority,
    cachePath: config.cachePath,
  });

  writeLoginState(config, { state: "error", message: "no client id" });
  assert.deepEqual(
    JSON.parse(textOf(await call(client, "auth_status"))).lastLoginError,
    "no client id",
  );
});

test("login hands back the code without waiting, and never restarts a session", async (t) => {
  const config = testConfig(t);
  const { deps, calls } = fakeDeps(config, { signedIn: false });
  const client = await connect(t, deps);

  const launch = JSON.parse(textOf(await call(client, "login")));
  assert.equal(calls.loginStarts, 1);
  assert.equal(launch.state, "pending");
  assert.equal(launch.verificationUri, "https://microsoft.com/devicelogin");
  assert.equal(launch.userCode, "ABC-XYZ");
  assert.equal(launch.expiresAt, "2030-01-01T00:00:00.000Z");
  assert.match(launch.instructions, /https:\/\/microsoft\.com\/devicelogin/);
  assert.match(launch.instructions, /ABC-XYZ/);
  assert.match(launch.instructions, /auth_status/);

  const existing = fakeDeps(config, { signedIn: true });
  existing.deps.startLogin = () => {
    throw new Error("a session exists, so nothing should be spawned");
  };
  const signedIn = await connect(t, existing.deps);
  const answer = JSON.parse(textOf(await call(signedIn, "login")));
  assert.equal(answer.state, "authorized");
  assert.equal(answer.account, "someone@example.com");
  assert.equal(existing.calls.loginStarts, 0);
});

test("logout deletes the saved session and says so", async (t) => {
  const config = testConfig(t);
  const { deps, calls } = fakeDeps(config);
  const client = await connect(t, deps);

  const result = await call(client, "logout");

  assert.equal(result.isError, undefined);
  assert.deepEqual(calls.logouts, [config.cachePath]);
  assert.deepEqual(JSON.parse(textOf(result)), {
    signedOut: true,
    cachePath: config.cachePath,
  });
});

test("resources expose folders, the inbox and one message", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const listed = await client.listResources();
  assert.deepEqual(
    listed.resources.map((resource) => resource.uri),
    ["outlook://folders", "outlook://inbox"],
  );
  const templates = await client.listResourceTemplates();
  assert.deepEqual(
    templates.resourceTemplates.map((template) => template.uriTemplate).sort(),
    ["outlook://inbox{?unread,limit}", "outlook://messages/{id}"],
  );

  const folders = await client.readResource({ uri: "outlook://folders" });
  assert.deepEqual(JSON.parse(resourceTextOf(folders)), {
    items: [{ id: "folder-1", displayName: "Inbox" }],
  });

  const inbox = await client.readResource({ uri: "outlook://inbox" });
  assert.deepEqual(JSON.parse(resourceTextOf(inbox)).items, [
    { id: "message-1", subject: "Hello" },
  ]);
  await client.readResource({ uri: "outlook://inbox?unread=true&limit=3" });
  assert.deepEqual(calls.list.at(-1), {
    folder: "inbox",
    unread: true,
    top: 3,
  });

  const message = await client.readResource({
    uri: "outlook://messages/a%2Fb",
  });
  assert.deepEqual(JSON.parse(resourceTextOf(message)), {
    id: "a/b",
    body: { content: "the whole body" },
  });

  await assert.rejects(
    client.readResource({ uri: "outlook://inbox?unread=false&limit=500" }),
    /limit must be an integer from 1 to 100/,
  );
});

test("a resource with no session explains the login step", async (t) => {
  const client = await connect(
    t,
    fakeDeps(testConfig(t), { session: false }).deps,
  );

  await assert.rejects(
    client.readResource({ uri: "outlook://folders" }),
    /No Outlook session\..*Call the login tool.*retry this call\./s,
  );
});
