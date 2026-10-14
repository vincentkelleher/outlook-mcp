import assert from "node:assert/strict";
import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  symlinkSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/sdk/client";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../src/mcp.ts";
import {
  draftToolNames,
  maxEchoedBodyChars,
  maxInlineBytes,
  readToolNames,
} from "../src/tools.ts";
import type { ToolDeps } from "../src/tools.ts";
import { loadConfig } from "../src/core/config.ts";
import type { Config } from "../src/core/config.ts";
import { writeLoginState } from "../src/core/device-login.ts";
import { AuthRequiredError, noSessionMessage } from "../src/core/errors.ts";
import type { Graph, Page } from "../src/core/graph.ts";
import {
  defaultPageSize,
  maxPageSize,
  maxSearchPageSize,
} from "../src/core/graph.ts";

const clientId = "00000000-0000-0000-0000-000000000000";
const readAnnotations = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: true,
};
const writeAnnotations = {
  readOnlyHint: false,
  destructiveHint: false,
  idempotentHint: false,
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
  created: unknown[];
  replied: [string, string][];
  forwarded: [string, string[], string][];
  patched: [string, unknown][];
  draftPages: unknown[];
  deleted: string[];
  uploaded: [string, unknown][];
  loginStarts: number;
  logouts: string[];
};

type Fake = { deps: ToolDeps; calls: Calls };

/** What a draft write returns: Graph hands back the message it stored. */
const storedDraft = {
  id: "draft-1",
  subject: "Weekly report",
  body: { content: "The numbers are in.", contentType: "text" },
  toRecipients: [{ emailAddress: { address: "ana@example.com", name: "Ana" } }],
  ccRecipients: [{ emailAddress: { address: "bo@example.com" } }],
  categories: ["work"],
  webLink: "https://outlook.office.com/mail/deeplink/compose?id=draft-1",
  isDraft: true,
};

/** Graph takes addresses nested; the fake unwraps them the same way. */
const addressed = (list: string[] | undefined) =>
  list?.map((address) => ({ emailAddress: { address } }));

/**
 * File names as a sender can send them: traversal, a Windows separator, a name
 * that is nothing but dots, and one ordinary name.
 */
const sentNames: Record<string, string> = {
  "attachment-1": "../escape/note.txt",
  oversized: "quarterly.pdf",
  traversal: "../../../../tmp/stolen.txt",
  backslash: "..\\..\\..\\tmp\\stolen.txt",
  dots: "....",
  plain: "quarterly report.pdf",
};

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
    created: [],
    replied: [],
    forwarded: [],
    patched: [],
    draftPages: [],
    deleted: [],
    uploaded: [],
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
        // A Graph file name is user data, whoever sent it.
        name: sentNames[attachmentId] ?? "note.txt",
        contentType: "text/plain",
        contentBytes:
          attachmentId === "oversized"
            ? Buffer.alloc(maxInlineBytes + 1, 0x61).toString("base64")
            : "aGVsbG8=",
      };
    },
    async createDraft(draft) {
      guard();
      calls.created.push(draft);
      return { ...storedDraft, toRecipients: addressed(draft.to) };
    },
    async createReplyDraft(messageId, comment) {
      guard();
      calls.replied.push([messageId, comment]);
      return {
        ...storedDraft,
        id: "draft-reply",
        subject: "RE: Weekly report",
        body: {
          content: `${comment}\n\nOriginal message`,
          contentType: "text",
        },
      };
    },
    async createForwardDraft(messageId, to, comment) {
      guard();
      calls.forwarded.push([messageId, to, comment]);
      return {
        ...storedDraft,
        id: "draft-forward",
        subject: "FW: Weekly report",
        toRecipients: addressed(to),
        ccRecipients: [],
        body: { content: comment, contentType: "text" },
      };
    },
    async updateDraft(messageId, patch) {
      guard();
      calls.patched.push([messageId, patch]);
      return {
        ...storedDraft,
        id: messageId,
        subject: patch.subject ?? storedDraft.subject,
        toRecipients: patch.to ? addressed(patch.to) : storedDraft.toRecipients,
        body:
          patch.body === undefined
            ? storedDraft.body
            : { content: patch.body, contentType: patch.bodyFormat ?? "text" },
      };
    },
    async listDrafts(options = {}) {
      guard();
      calls.draftPages.push(options);
      return {
        items: [
          {
            id: "draft-1",
            subject: "Weekly report",
            isDraft: true,
            toRecipients: storedDraft.toRecipients,
          },
        ],
        nextCursor: "https://graph/next-drafts?skiptoken=1",
      };
    },
    async deleteDraft(messageId) {
      guard();
      calls.deleted.push(messageId);
    },
    async addAttachment(messageId, attachment) {
      guard();
      calls.uploaded.push([messageId, attachment]);
      return { ...storedDraft, id: messageId };
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
    logout: (session) => {
      calls.logouts.push(session.cachePath);
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

  // The full surface is the registry's own two halves: the read and sign-in
  // tools plus the draft tools this mode exposes. The read-only half of the same
  // pair is pinned in test/no-send.test.ts.
  assert.deepEqual(
    listed.tools.map((tool) => tool.name).sort(),
    [...readToolNames, ...draftToolNames].sort(),
  );
  assert.equal(
    listed.tools.find((tool) => tool.name === "list_emails")?.title,
    "List emails",
  );

  // Everything that only reads the mailbox, which is the read half less the two
  // sign-in tools that write the local session instead of reading mail.
  const signed = readToolNames.filter(
    (name) => name !== "login" && name !== "logout",
  );
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

  // The annotations are what make a careful client ask the human first, so they
  // are asserted on everything that writes a draft. `list_drafts` is the one
  // member of the draft group that only reads, and `delete_draft` is the one
  // that destroys, so both sit outside the plain write shape.
  for (const name of draftToolNames.filter(
    (name) => name !== "delete_draft" && name !== "list_drafts",
  )) {
    assert.deepEqual(
      listed.tools.find((tool) => tool.name === name)?.annotations,
      writeAnnotations,
      `${name} annotations`,
    );
  }
  assert.deepEqual(
    listed.tools.find((tool) => tool.name === "list_drafts")?.annotations,
    readAnnotations,
    "listing drafts is a GET and must not advertise itself as a write",
  );
  const deleteTool = listed.tools.find((tool) => tool.name === "delete_draft");
  assert.deepEqual(deleteTool?.annotations, {
    ...writeAnnotations,
    destructiveHint: true,
  });

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
  assert.deepEqual(bounds("list_emails"), {
    minimum: 1,
    maximum: maxPageSize,
  });
  assert.deepEqual(bounds("list_drafts"), {
    minimum: 1,
    maximum: maxPageSize,
  });
  assert.deepEqual(bounds("search_emails"), {
    minimum: 1,
    maximum: maxSearchPageSize,
  });

  // The page sizes are one number in one place, so the sentence an agent reads
  // cannot promise a bound the schema then refuses.
  const limitText = (name: string) =>
    (
      listed.tools.find((tool) => tool.name === name)?.inputSchema as {
        properties?: { limit?: { description?: string } };
      }
    ).properties?.limit?.description ?? "";
  for (const [name, maximum] of [
    ["list_emails", maxPageSize],
    ["list_drafts", maxPageSize],
    ["search_emails", maxSearchPageSize],
  ] as const) {
    assert.match(
      limitText(name),
      new RegExp(`1 to ${maximum}\\b`),
      `${name} limit text`,
    );
    assert.match(
      limitText(name),
      new RegExp(`Defaults to ${defaultPageSize}\\b`),
      `${name} default text`,
    );
  }

  for (const [name, maximum, required] of [
    ["list_emails", maxPageSize, {}],
    ["search_emails", maxSearchPageSize, { query: "invoice" }],
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

test("an attachment name from the sender cannot lead the write out of the directory", async (t) => {
  const config = testConfig(t);
  const { deps } = fakeDeps(config);
  const client = await connect(t, deps);

  // The name is the sender's file name, so it is reduced to one leaf and the
  // destination is verified: traversal, separators and dot-only names stay put.
  for (const [attachmentId, leaf] of [
    ["traversal", "stolen.txt"],
    ["backslash", "stolen.txt"],
    ["dots", "attachment"],
    ["plain", "quarterly report.pdf"],
  ] as const) {
    const result = await call(client, "download_attachment", {
      messageId: "message-1",
      attachmentId,
    });
    assert.equal(result.isError, undefined, textOf(result));
    const payload = JSON.parse(textOf(result));
    assert.equal(payload.name, leaf, `${attachmentId} kept a dangerous name`);
    assert.equal(payload.path, join(config.attachmentDir, leaf));
    assert.equal(readFileSync(payload.path, "utf8"), "hello");
  }
  assert.deepEqual(readdirSync(config.attachmentDir).sort(), [
    "attachment",
    "quarterly report.pdf",
    "stolen.txt",
  ]);
  assert.equal(
    existsSync(join(config.attachmentDir, "..", "stolen.txt")),
    false,
    "a name walked out of the attachment directory",
  );

  // And a symlink already waiting at that destination would carry the write with
  // it, so nothing is written through one either.
  const outside = mkdtempSync(join(tmpdir(), "outlook-mcp-outside-"));
  t.after(() => rmSync(outside, { recursive: true, force: true }));
  symlinkSync(
    join(outside, "note.txt"),
    join(config.attachmentDir, "note.txt"),
  );

  const refused = await call(client, "download_attachment", {
    messageId: "message-1",
    attachmentId: "attachment-1",
  });
  assert.equal(refused.isError, true);
  assert.match(textOf(refused), /outside the attachment directory/);
  assert.equal(
    existsSync(join(outside, "note.txt")),
    false,
    "the write followed the symlink",
  );
});

test("create_draft writes what it is given and reads back the stored draft", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const result = await call(client, "create_draft", {
    to: ["ana@example.com"],
    cc: ["bo@example.com"],
    subject: "Weekly report",
    body: "The numbers are in.",
    categories: ["work"],
  });

  assert.equal(result.isError, undefined);
  assert.deepEqual(calls.created, [
    {
      to: ["ana@example.com"],
      cc: ["bo@example.com"],
      subject: "Weekly report",
      body: "The numbers are in.",
      // text is the default, taken from the schema the client is shown.
      bodyFormat: "text",
      categories: ["work"],
    },
  ]);
  const payload = JSON.parse(textOf(result));
  assert.equal(payload.id, "draft-1");
  assert.equal(payload.subject, "Weekly report");
  assert.deepEqual(payload.to, ["ana@example.com"]);
  assert.deepEqual(payload.cc, ["bo@example.com"]);
  assert.equal(payload.body, "The numbers are in.");
  assert.equal(payload.bodyFormat, "text");
  assert.deepEqual(payload.categories, ["work"]);
  assert.equal(payload.sent, false);
  assert.equal(payload.webLink, storedDraft.webLink);
  assert.match(payload.note, /This is a draft: nothing has been sent/);
  assert.match(payload.note, /Drafts folder/);
  assert.deepEqual(result.structuredContent, payload);
});

test("reply and forward drafts carry the comment and stay threaded", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const reply = await call(client, "create_reply_draft", {
    messageId: "message-1",
    comment: "Thanks for this.",
  });
  assert.equal(reply.isError, undefined);
  assert.deepEqual(calls.replied, [["message-1", "Thanks for this."]]);
  const answered = JSON.parse(textOf(reply));
  assert.equal(answered.subject, "RE: Weekly report");
  assert.match(answered.body, /^Thanks for this\./);
  assert.equal(answered.sent, false);

  const forward = await call(client, "create_forward_draft", {
    messageId: "message-1",
    to: ["bo@example.com"],
    comment: "Passing this on.",
  });
  assert.equal(forward.isError, undefined);
  assert.deepEqual(calls.forwarded, [
    ["message-1", ["bo@example.com"], "Passing this on."],
  ]);
  const moved = JSON.parse(textOf(forward));
  assert.equal(moved.subject, "FW: Weekly report");
  assert.deepEqual(moved.to, ["bo@example.com"]);
  assert.equal(moved.sent, false);
});

test("update_draft writes only the fields it is given", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const revised = await call(client, "update_draft", {
    messageId: "draft-1",
    subject: "Weekly report, corrected",
  });
  assert.equal(revised.isError, undefined);
  assert.deepEqual(calls.patched, [
    ["draft-1", { subject: "Weekly report, corrected" }],
  ]);
  assert.equal(JSON.parse(textOf(revised)).subject, "Weekly report, corrected");

  const nothing = await call(client, "update_draft", { messageId: "draft-1" });
  assert.equal(nothing.isError, true);
  assert.match(textOf(nothing), /at least one of/);
  assert.equal(calls.patched.length, 1, "an empty patch reached Graph");

  // A format with no body would be a patch that changes nothing at all.
  const formatOnly = await call(client, "update_draft", {
    messageId: "draft-1",
    bodyFormat: "html",
  });
  assert.equal(formatOnly.isError, true);
  assert.match(textOf(formatOnly), /so send body as well/);
  assert.equal(calls.patched.length, 1);

  // A reply echoes the quoted original, so the echo has a ceiling of its own.
  const stored = "x".repeat(maxEchoedBodyChars + 500);
  const long = await call(client, "update_draft", {
    messageId: "draft-1",
    body: stored,
  });
  const echoed = JSON.parse(textOf(long));
  assert.equal(echoed.body, stored.slice(0, maxEchoedBodyChars));
  assert.match(
    echoed.note,
    new RegExp(
      `first ${maxEchoedBodyChars} of ${stored.length} stored characters`,
    ),
    "the echo says it is partial",
  );
});

test("list_drafts reads the drafts and says none of them is sent", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const result = await call(client, "list_drafts", { limit: 5 });

  assert.equal(result.isError, undefined);
  assert.deepEqual(calls.draftPages, [{ top: 5, cursor: undefined }]);
  const payload = JSON.parse(textOf(result));
  assert.equal(payload.items[0].id, "draft-1");
  assert.equal(payload.nextCursor, "https://graph/next-drafts?skiptoken=1");
  assert.match(payload.note, /None of these drafts have been sent/);
});

test("delete_draft removes one draft and reports it", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const result = await call(client, "delete_draft", { messageId: "draft-1" });

  assert.equal(result.isError, undefined);
  assert.deepEqual(calls.deleted, ["draft-1"]);
  assert.deepEqual(JSON.parse(textOf(result)), {
    deleted: true,
    messageId: "draft-1",
    note: "The draft is gone from the mailbox. Nothing was sent.",
  });
});

test("add_attachment uploads a file to the draft within the schema's bounds", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const added = await call(client, "add_attachment", {
    messageId: "draft-1",
    name: "note.txt",
    contentType: "text/plain",
    contentBase64: "aGVsbG8=",
  });
  assert.equal(added.isError, undefined);
  assert.deepEqual(calls.uploaded, [
    [
      "draft-1",
      {
        name: "note.txt",
        contentType: "text/plain",
        contentBase64: "aGVsbG8=",
      },
    ],
  ]);
  assert.equal(JSON.parse(textOf(added)).id, "draft-1");

  await call(client, "add_attachment", {
    messageId: "draft-2",
    name: "note.txt",
    contentBase64: "aGVsbG8=",
  });
  const defaulted = calls.uploaded[1][1] as { contentType: string };
  assert.equal(defaulted.contentType, "application/octet-stream");

  // Content that is not base64 never reaches Graph, let alone the draft.
  const broken = await call(client, "add_attachment", {
    messageId: "draft-1",
    name: "note.txt",
    contentType: "text/plain",
    contentBase64: "not base64 at all",
  });
  assert.equal(broken.isError, true);
  assert.match(textOf(broken), /Invalid arguments for tool add_attachment/);
  assert.equal(calls.uploaded.length, 2);
});

test("recipients are validated addresses before anything is written", async (t) => {
  const { deps, calls } = fakeDeps(testConfig(t));
  const client = await connect(t, deps);

  const displayForm = await call(client, "create_draft", {
    to: ["Ana <ana@example.com>"],
    subject: "Weekly report",
  });
  assert.equal(displayForm.isError, true);
  assert.match(textOf(displayForm), /Invalid email address/);

  const nobody = await call(client, "create_draft", { to: [] });
  assert.equal(nobody.isError, true);

  const forwarded = await call(client, "create_forward_draft", {
    messageId: "message-1",
    to: ["ana at example.com"],
    comment: "Passing this on.",
  });
  assert.equal(forwarded.isError, true);

  assert.deepEqual(calls.created, []);
  assert.deepEqual(calls.forwarded, []);
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
    readonly: false,
    scopes: ["User.Read", "Mail.ReadWrite", "email", "profile"],
  });

  writeLoginState(config, { state: "error", message: "no client id" });
  assert.deepEqual(
    JSON.parse(textOf(await call(client, "auth_status"))).lastLoginError,
    "no client id",
  );
});

test("auth_status carries the code of a sign-in that is still open", async (t) => {
  const config = testConfig(t);
  const { deps } = fakeDeps(config);
  const client = await connect(t, deps);

  // The code belongs to the flow, not to whoever called login, so an agent told
  // to poll auth_status has to be able to find it here.
  writeLoginState(config, {
    state: "pending",
    verificationUri: "https://microsoft.com/devicelogin",
    userCode: "ABC-XYZ",
    expiresAt: "2030-01-01T00:00:00.000Z",
  });
  const open = JSON.parse(textOf(await call(client, "auth_status")));
  assert.equal(open.signInPending, true);
  assert.equal(open.verificationUri, "https://microsoft.com/devicelogin");
  assert.equal(open.userCode, "ABC-XYZ");
  assert.equal(open.expiresAt, "2030-01-01T00:00:00.000Z");
  assert.equal(open.lastLoginError, undefined);

  // Past its own deadline nobody is polling Microsoft any more, so the code is
  // not offered, and a flow that is not happening is not reported as one.
  writeLoginState(config, {
    state: "pending",
    userCode: "DEAD-CODE",
    expiresAt: new Date(Date.now() - 60_000).toISOString(),
  });
  const dead = JSON.parse(textOf(await call(client, "auth_status")));
  assert.equal(
    dead.signInPending,
    undefined,
    "a dead code is still advertised",
  );
  assert.equal(dead.userCode, undefined);
  assert.equal(dead.signedIn, true);

  writeLoginState(config, { state: "authorized" });
  assert.equal(
    JSON.parse(textOf(await call(client, "auth_status"))).signInPending,
    undefined,
  );
});

test("auth_status reports why an account that exists cannot be used", async (t) => {
  const config = testConfig(t);
  const { deps } = fakeDeps(config);
  deps.authStatus = async () => ({
    signedIn: false,
    account: "someone@example.com",
    reason: "needsConsent",
    authority: config.authority,
    cachePath: config.cachePath,
  });
  const client = await connect(t, deps);

  const reported = JSON.parse(textOf(await call(client, "auth_status")));
  assert.equal(reported.signedIn, false);
  assert.equal(reported.reason, "needsConsent");
  assert.deepEqual(reported.scopes, [
    "User.Read",
    "Mail.ReadWrite",
    "email",
    "profile",
  ]);
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

test("login starts a flow when the cached grant is too narrow", async (t) => {
  const config = testConfig(t);
  const { deps, calls } = fakeDeps(config);
  // An account is present, which is exactly the state that used to be answered
  // with "already signed in" while every scoped call kept failing.
  deps.authStatus = async () => ({
    signedIn: false,
    account: "someone@example.com",
    reason: "needsConsent",
    authority: config.authority,
    cachePath: config.cachePath,
  });
  const client = await connect(t, deps);

  const launch = JSON.parse(textOf(await call(client, "login")));

  assert.equal(calls.loginStarts, 1, "re-consent was short-circuited");
  assert.equal(launch.state, "pending");
  assert.equal(launch.userCode, "ABC-XYZ");
  assert.match(launch.instructions, /ABC-XYZ/);
});

test("login starts no flow when the session could not be checked", async (t) => {
  // Microsoft unreachable is not a grant too narrow. Starting a device-code flow
  // here would spend one of a rate-limited kind, on a machine whose sign-in is
  // fine, and would need the network that just failed.
  const config = testConfig(t);
  const { deps, calls } = fakeDeps(config);
  deps.authStatus = async () => ({
    signedIn: false,
    account: "someone@example.com",
    reason: "unavailable",
    authority: config.authority,
    cachePath: config.cachePath,
  });
  const client = await connect(t, deps);

  const answer = JSON.parse(textOf(await call(client, "login")));

  assert.equal(calls.loginStarts, 0, "a failed check started a flow");
  assert.equal(answer.state, "error");
  assert.match(answer.instructions, /Retry the call later/);
  assert.doesNotMatch(answer.instructions, /ABC-XYZ|verification/);
});

test("a usable token means login starts no flow", async (t) => {
  const config = testConfig(t);
  const { deps, calls } = fakeDeps(config, { signedIn: true });
  const client = await connect(t, deps);

  const answer = JSON.parse(textOf(await call(client, "login")));

  // A token that works is the only reason not to start a flow.
  assert.equal(answer.state, "authorized");
  assert.equal(answer.account, "someone@example.com");
  assert.equal(calls.loginStarts, 0);
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
    new RegExp(`limit must be an integer from 1 to ${maxPageSize}`),
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
