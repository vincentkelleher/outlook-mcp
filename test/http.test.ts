import assert from "node:assert/strict";
import { request } from "node:http";
import type { IncomingHttpHeaders } from "node:http";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import {
  Client,
  StreamableHTTPClientTransport,
} from "@modelcontextprotocol/client";
import { loadConfig } from "../src/core/config.ts";
import { admitRequest, isLoopbackHost } from "../src/core/http.ts";
import type { Admission } from "../src/core/http.ts";
import { startHttp } from "../src/mcp.ts";
import { draftToolNames } from "../src/tools.ts";
import type { ToolDeps } from "../src/tools.ts";
import type { Graph } from "../src/core/graph.ts";
import type { Config } from "../src/core/config.ts";

const clientId = "00000000-0000-0000-0000-000000000000";

type TestContext = { after: (fn: () => void) => void };

/** A config under a throwaway home, removed when the test ends. */
function configFor(
  t: TestContext,
  env: Record<string, string | undefined> = {},
) {
  const home = mkdtempSync(join(tmpdir(), "outlook-http-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  return loadConfig({
    HOME: home,
    OUTLOOK_OAUTH_CLIENT_ID: clientId,
    OUTLOOK_MCP_STATE_DIR: join(home, "state"),
    OUTLOOK_MCP_ATTACHMENT_DIR: join(home, "attachments"),
    ...env,
  });
}

/** A mailbox that answers, so a transport test is about the transport. */
function fakeDeps(config: Config): ToolDeps {
  const graph = {
    async listMessages() {
      return { items: [{ id: "message-1", subject: "Hello" }] };
    },
    async searchMessages() {
      return { items: [] };
    },
    async getMessage(id: string) {
      return { id, body: { content: "the whole body" } };
    },
    async listFolders() {
      return { items: [{ id: "folder-1", displayName: "Inbox" }] };
    },
    async listAttachments() {
      return [];
    },
    async getAttachment() {
      return { id: "attachment-1" };
    },
    async createDraft() {
      return { id: "draft-1", isDraft: true };
    },
    async createReplyDraft() {
      return { id: "draft-1", isDraft: true };
    },
    async createForwardDraft() {
      return { id: "draft-1", isDraft: true };
    },
    async updateDraft() {
      return { id: "draft-1", isDraft: true };
    },
    async listDrafts() {
      return { items: [] };
    },
    async deleteDraft() {},
    async addAttachment() {
      return { id: "draft-1" };
    },
  } as unknown as Graph;
  return {
    graph,
    config,
    authStatus: async () => ({
      signedIn: true,
      account: "someone@example.com",
      authority: config.authority,
      cachePath: config.cachePath,
    }),
    startLogin: () => {},
    logout: () => {},
  };
}

/** Serve a fake mailbox on a free loopback port, and stop it when the test ends. */
async function serving(
  t: TestContext,
  env: Record<string, string | undefined> = {},
) {
  const handle = await startHttp(fakeDeps(configFor(t, env)), {
    host: "127.0.0.1",
    port: 0,
  });
  t.after(() => handle.close());
  return handle;
}

/** One raw request, with a Host header a browser would never be allowed to send. */
function raw(input: {
  url: string;
  method?: string;
  host?: string;
  headers?: Record<string, string>;
  body?: string;
}): Promise<{
  status: number;
  body: string;
  headers: IncomingHttpHeaders;
}> {
  return new Promise((resolve, reject) => {
    const target = new URL(input.url);
    const req = request(
      {
        hostname: target.hostname,
        port: target.port,
        path: `${target.pathname}${target.search}`,
        method: input.method ?? "GET",
        headers: {
          ...(input.host === undefined ? {} : { host: input.host }),
          ...input.headers,
        },
      },
      (res) => {
        const chunks: Buffer[] = [];
        res.on("data", (chunk: Buffer) => chunks.push(chunk));
        res.on("end", () =>
          resolve({
            status: res.statusCode ?? 0,
            body: Buffer.concat(chunks).toString("utf8"),
            headers: res.headers,
          }),
        );
      },
    );
    req.on("error", reject);
    if (input.body !== undefined) req.write(input.body);
    req.end();
  });
}

test("loopback names answer, and everything else is refused by name", () => {
  for (const host of [
    "127.0.0.1:3939",
    "localhost:3939",
    "[::1]:3939",
    "LOCALHOST",
    "localhost",
  ])
    assert.equal(
      admitRequest({ host, allowedHosts: [], allowedOrigins: [] }).allowed,
      true,
      host,
    );
  assert.equal(isLoopbackHost("127.0.0.1"), true);
  assert.equal(isLoopbackHost("evil.example"), false);

  // The answer says what to set, because "403" alone sends nobody to the docs.
  const refused = refusalOf(
    admitRequest({
      host: "evil.example:3939",
      allowedHosts: [],
      allowedOrigins: [],
    }),
  );
  assert.equal(refused.status, 403);

  assert.equal(
    admitRequest({
      host: "outlook.internal:3939",
      allowedHosts: ["outlook.internal"],
      allowedOrigins: [],
    }).allowed,
    true,
  );
});

test("a request that names no host, or two, is not served", () => {
  assert.equal(
    refusalOf(admitRequest({ allowedHosts: [], allowedOrigins: [] })).status,
    400,
  );
  // A repeated Host header is a request that cannot be answered honestly.
  assert.equal(
    refusalOf(
      admitRequest({
        host: ["127.0.0.1:3939", "evil.example"],
        allowedHosts: [],
        allowedOrigins: [],
      }),
    ).status,
    400,
  );
});

test("a browser origin is served only when it is the origin addressed", () => {
  const same = admitRequest({
    host: "127.0.0.1:3939",
    origin: "http://127.0.0.1:3939",
    allowedHosts: [],
    allowedOrigins: [],
  });
  assert.equal(same.allowed, true);

  // DNS rebinding: the page is someone else's, the Host is now this machine's.
  const rebound = refusalOf(
    admitRequest({
      host: "127.0.0.1:3939",
      origin: "http://evil.example",
      allowedHosts: [],
      allowedOrigins: [],
    }),
  );
  assert.equal(rebound.status, 403);
  assert.match(rebound.reason, /OUTLOOK_MCP_ALLOWED_ORIGINS/);

  assert.equal(
    admitRequest({
      host: "127.0.0.1:3939",
      origin: "http://localhost:5173",
      allowedHosts: [],
      allowedOrigins: ["http://localhost:5173"],
    }).allowed,
    true,
  );
  // `http://localhost` and `localhost:80` are one origin, not two spellings.
  assert.equal(
    admitRequest({
      host: "localhost",
      origin: "http://localhost",
      allowedHosts: [],
      allowedOrigins: [],
    }).allowed,
    true,
  );
});

test("the HTTP server completes a real MCP session and reports it", async (t) => {
  const handle = await serving(t);
  assert.match(handle.url, /^http:\/\/127\.0\.0\.1:\d+\/mcp$/);

  const client = new Client({ name: "http-test", version: "0.0.0" });
  await client.connect(new StreamableHTTPClientTransport(new URL(handle.url)));
  const tools = await client.listTools();
  assert.ok(tools.tools.length > 10, "the registry arrived over HTTP");
  assert.ok(
    tools.tools.some((tool) => tool.name === "list_emails"),
    "list_emails is missing",
  );

  const status = (await client.callTool({
    name: "auth_status",
    arguments: {},
  })) as { structuredContent?: { signedIn?: boolean } };
  assert.equal(status.structuredContent?.signedIn, true);

  const inbox = await client.readResource({ uri: "outlook://folders" }, {});
  assert.match(JSON.stringify(inbox), /folder-1/);

  const health = await healthOf(handle);
  assert.equal(health.status, "ok");
  assert.equal(health.sessions, 1, "the open session is counted");

  await client.close();
});

test("a DELETE retires its session, and the count follows it", async (t) => {
  const handle = await serving(t);
  const opened = await raw({
    url: handle.url,
    method: "POST",
    headers: {
      "content-type": "application/json",
      accept: "application/json, text/event-stream",
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2025-06-18",
        capabilities: {},
        clientInfo: { name: "raw", version: "0" },
      },
    }),
  });
  assert.equal(opened.status, 200);
  const sessionId = opened.headers["mcp-session-id"];
  assert.equal(
    typeof sessionId,
    "string",
    "the session id is offered to the client",
  );
  assert.equal((await healthOf(handle)).sessions, 1);

  const closed = await raw({
    url: handle.url,
    method: "DELETE",
    headers: { "mcp-session-id": String(sessionId) },
  });
  assert.equal(closed.status, 200);
  assert.equal((await healthOf(handle)).sessions, 0, "the session was retired");

  // A client that reconnects under a retired id is told to start again, not
  // silently served a session it was never given.
  const reused = await raw({
    url: handle.url,
    method: "POST",
    headers: {
      "content-type": "application/json",
      accept: "application/json, text/event-stream",
      "mcp-session-id": String(sessionId),
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 2,
      method: "tools/list",
      params: {},
    }),
  });
  assert.equal(reused.status, 404);
});

test("readonly mode reaches the tool list over HTTP too", async (t) => {
  const handle = await serving(t, { OUTLOOK_MCP_READONLY: "1" });
  const client = new Client({ name: "http-test", version: "0.0.0" });
  await client.connect(new StreamableHTTPClientTransport(new URL(handle.url)));
  const names = (await client.listTools()).tools.map((tool) => tool.name);
  for (const draft of draftToolNames)
    assert.equal(names.includes(draft), false, `${draft} leaked into readonly`);
  await client.close();
});

test("anything that is not the MCP endpoint is answered, not served", async (t) => {
  const handle = await serving(t);
  const missing = await raw({ url: `http://127.0.0.1:${portOf(handle)}/nope` });
  assert.equal(missing.status, 404);
  assert.match(missing.body, /MCP endpoint is \/mcp/);

  const junk = await raw({
    url: handle.url,
    method: "POST",
    headers: { "content-type": "application/json" },
    body: "{not json",
  });
  assert.equal(junk.status, 400);
  assert.match(junk.body, /not JSON/);

  // A session-opening call is the only thing that may create one.
  const unopened = await raw({
    url: handle.url,
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "tools/list",
      params: {},
    }),
  });
  assert.equal(unopened.status, 400);
  assert.match(unopened.body, /initialize request/);

  const stale = await raw({
    url: handle.url,
    method: "POST",
    headers: {
      "content-type": "application/json",
      "mcp-session-id": "00000000-dead-beef-0000-000000000000",
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2025-06-18",
        capabilities: {},
        clientInfo: { name: "stale", version: "0" },
      },
    }),
  });
  assert.equal(stale.status, 404, "an unknown session opens nothing");
});

test("a Host header naming another machine is refused before the body", async (t) => {
  const handle = await serving(t);
  const rebound = await raw({
    url: handle.url,
    method: "POST",
    host: "evil.example:3939",
    headers: {
      "content-type": "application/json",
      origin: "http://evil.example",
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2025-06-18",
        capabilities: {},
        clientInfo: { name: "rebound", version: "0" },
      },
    }),
  });
  assert.equal(rebound.status, 403);
  assert.match(rebound.body, /OUTLOOK_MCP_ALLOWED_HOSTS/);
  // The session was never opened, so nothing was created to clean up.
  assert.equal((await healthOf(handle)).sessions, 0);
});

test("a port already in use fails the start, not the first request", async (t) => {
  const handle = await serving(t);
  const port = portOf(handle);
  await assert.rejects(
    () => startHttp(fakeDeps(configFor(t)), { host: "127.0.0.1", port }),
    (error: unknown) =>
      (error as NodeJS.ErrnoException).code === "EADDRINUSE" ||
      /EADDRINUSE/.test(String(error)),
  );
});

/** Narrow an admission to the half that carries the answer. */
function refusalOf(admission: Admission): { status: number; reason: string } {
  if (admission.allowed) assert.fail("the request was expected to be refused");
  return { status: admission.status, reason: admission.reason };
}

function portOf(handle: { url: string }): number {
  return Number(new URL(handle.url).port);
}

async function healthOf(handle: { url: string }): Promise<{
  status: string;
  sessions: number;
  readonly: boolean;
}> {
  return (
    await fetch(`http://127.0.0.1:${portOf(handle)}/health`)
  ).json() as Promise<{ status: string; sessions: number; readonly: boolean }>;
}
