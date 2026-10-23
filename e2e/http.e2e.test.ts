import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/sdk/client";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const binArgs = (process.env.OUTLOOK_E2E_BIN ?? "src/index.ts").split(" ");
const placeholderClientId = "00000000-0000-0000-0000-000000000000";

type Running = {
  url: string;
  stdout: string[];
  stderr: string;
  stop: () => Promise<number | null>;
};

/**
 * Start the real entry over HTTP on a port the OS picks, and wait for the line
 * it prints when the listener is up: that line is what a person reading
 * `docker logs` reads, so waiting for it is the test of the message too.
 */
async function start(t: { after: (fn: () => void) => void }): Promise<Running> {
  const home = mkdtempSync(join(tmpdir(), "outlook-skill-http-e2e-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const child = spawn(
    process.execPath,
    [...binArgs, "serve", "--http", "--port", "0"],
    {
      cwd: repo,
      env: {
        ...process.env,
        HOME: home,
        OUTLOOK_OAUTH_CLIENT_ID: placeholderClientId,
        OUTLOOK_MCP_STATE_DIR: join(home, "state"),
        OUTLOOK_MCP_ATTACHMENT_DIR: join(home, "attachments"),
        // Never inherited from the shell here, exactly as the stdio suite holds it.
        OUTLOOK_MCP_READONLY: "",
      },
      stdio: ["ignore", "pipe", "pipe"],
    },
  );

  // A server nobody stops holds its stdout pipe open, and the test runner waits
  // on that read for ever. SIGTERM is the shutdown the CLI itself advertises.
  t.after(() => {
    if (!child.killed) child.kill("SIGTERM");
  });

  const stdout: string[] = [];
  let stderr = "";
  const printed: { url?: string; reject?: (error: unknown) => void } = {};
  child.stdout.on("data", (chunk: Buffer) => stdout.push(chunk.toString()));
  child.stderr.on("data", (chunk: Buffer) => {
    stderr += chunk.toString();
    const url = /http:\/\/127\.0\.0\.1:(\d+)\/mcp/.exec(stderr);
    if (url && !printed.url) printed.url = url[0];
  });
  // A process that dies before it binds has an error to read, not a port.
  child.on("error", (error) => printed.reject?.(error));

  const url = await Promise.race([
    new Promise<string>((resolve, reject) => {
      printed.reject = reject;
      const wait = setInterval(() => {
        if (printed.url) {
          clearInterval(wait);
          resolve(printed.url);
        }
      }, 25);
      const givenUp = setTimeout(() => {
        clearInterval(wait);
        reject(
          new Error(`the server never printed its URL; stderr: ${stderr}`),
        );
      }, 20_000);
      givenUp.unref();
    }),
    new Promise<never>((_resolve, reject) =>
      child.on("exit", (code) =>
        reject(
          new Error(`the server exited early with ${code}; stderr: ${stderr}`),
        ),
      ),
    ),
  ]);

  return {
    url,
    stdout,
    get stderr() {
      return stderr;
    },
    stop: () =>
      new Promise((resolve) => {
        child.on("exit", (code) => resolve(code));
        child.kill("SIGTERM");
      }),
  };
}

test("serve --http completes a real handshake with a real client", async (t) => {
  const server = await start(t);
  const client = new Client({ name: "e2e-http", version: "0.0.0" });
  await client.connect(new StreamableHTTPClientTransport(new URL(server.url)));

  const tools = await client.listTools();
  assert.ok(
    tools.tools.some((tool) => tool.name === "create_draft"),
    "the draft tools are offered outside read-only mode",
  );
  const status = (await client.callTool({
    name: "auth_status",
    arguments: {},
  })) as { structuredContent?: { signedIn?: boolean; cachePath?: string } };
  // No session in a throwaway home: the answer is the remediation, not a crash.
  assert.equal(status.structuredContent?.signedIn, false);

  const health = (await (
    await fetch(new URL("/health", server.url))
  ).json()) as { status: string; readonly: boolean; sessions: number };
  assert.equal(health.status, "ok");
  assert.equal(health.readonly, false);
  assert.equal(health.sessions, 1);

  await client.close();
  // Nothing but a diagnostic may reach stdout, even where JSON-RPC does not.
  assert.equal(server.stdout.join(""), "");
  assert.match(
    server.stderr,
    /Outlook MCP serving on http:\/\/127\.0\.0\.1:\d+\/mcp/,
  );
  assert.match(server.stderr, /Nothing on that port authenticates/);

  // And the logs, at the level a server runs at by default: what it bound, which
  // session opened, and what that session asked for. This is the output a person
  // reads with `docker logs`, so it is asserted here rather than only in a unit
  // test that stubs the transport.
  assert.match(
    server.stderr,
    /info \d+ http: serving MCP over Streamable HTTP url=http:\/\/127\.0\.0\.1:\d+\/mcp/,
  );
  assert.match(
    server.stderr,
    /info \d+ http: session opened sessionId=[0-9a-f-]{36} sessions=1 readonly=false/,
  );
  assert.match(
    server.stderr,
    /info \d+ mcp: tool call sessionId=[0-9a-f-]{36} tool=auth_status/,
  );
  assert.match(
    server.stderr,
    /info \d+ mcp: tool call completed sessionId=[0-9a-f-]{36} tool=auth_status ms=\d+ bytes=\d+/,
  );
  // Sign-in state is a question the tool answered, so the answer is in the log
  // too — with no account name in it, because there is none.
  assert.match(
    server.stderr,
    /info \d+ auth: session check found no cached account ms=\d+/,
  );
});

test("SIGTERM closes the listener and the process exits", async (t) => {
  const server = await start(t);
  const code = await server.stop();
  assert.equal(code, 0, `exit code ${code}; stderr: ${server.stderr}`);
  await assert.rejects(() => fetch(new URL("/health", server.url)));
});

test("the port flag answers the environment, and a bad one is refused", async () => {
  const home = mkdtempSync(join(tmpdir(), "outlook-skill-http-e2e-"));
  const child = spawn(
    process.execPath,
    [...binArgs, "serve", "--http", "--port", "70000"],
    {
      cwd: repo,
      env: {
        ...process.env,
        HOME: home,
        OUTLOOK_OAUTH_CLIENT_ID: placeholderClientId,
      },
      stdio: ["ignore", "ignore", "pipe"],
    },
  );
  const output = await new Promise<{ code: number | null; text: string }>(
    (resolve) => {
      let text = "";
      child.stderr?.on("data", (chunk: Buffer) => (text += chunk.toString()));
      child.on("close", (exitCode) => resolve({ code: exitCode, text }));
    },
  );
  rmSync(home, { recursive: true, force: true });
  assert.equal(output.code, 1);
  assert.match(output.text, /--port must be an integer from 0 to 65535/);
});
