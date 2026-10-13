import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { draftToolNames, readToolNames } from "../src/tools.ts";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const binArgs = (process.env.OUTLOOK_E2E_BIN ?? "src/index.ts").split(" ");
const placeholderClientId = "00000000-0000-0000-0000-000000000000";

type Frame = {
  id?: number;
  jsonrpc?: string;
  result?: {
    serverInfo?: { name?: string; version?: string };
    protocolVersion?: string;
    instructions?: string;
    tools?: { name: string }[];
    isError?: boolean;
    content?: { type: string; text?: string }[];
  };
};

/**
 * Serves the entry module over real stdio pipes and records every raw byte it
 * writes, because a single stray byte on stdout corrupts JSON-RPC.
 */
function startServer(home: string, env: Record<string, string> = {}) {
  const child = spawn(process.execPath, binArgs, {
    cwd: repo,
    env: {
      ...process.env,
      HOME: home,
      OUTLOOK_OAUTH_CLIENT_ID: placeholderClientId,
      OUTLOOK_MCP_STATE_DIR: join(home, "state"),
      OUTLOOK_MCP_ATTACHMENT_DIR: join(home, "attachments"),
      // An env var, never a flag: one process cannot disagree with itself about
      // whether it may write, and login and serve must ask Microsoft for the
      // same permissions. It is set here or not at all, never inherited from the
      // shell, and never per command.
      OUTLOOK_MCP_READONLY: "",
      ...env,
    },
    stdio: ["pipe", "pipe", "pipe"],
  });
  const raw: Buffer[] = [];
  const lines: string[] = [];
  const notJson: string[] = [];
  const stderr: Buffer[] = [];
  let buffered = Buffer.alloc(0);
  const waiters: { id: number; resolve: (frame: Frame) => void }[] = [];

  child.stdout.on("data", (chunk: Buffer) => {
    raw.push(chunk);
    buffered = Buffer.concat([buffered, chunk]);
    let cut = buffered.indexOf(0x0a);
    while (cut >= 0) {
      const line = buffered.subarray(0, cut).toString("utf8");
      buffered = buffered.subarray(cut + 1);
      lines.push(line);
      let frame: Frame;
      try {
        frame = JSON.parse(line) as Frame;
      } catch {
        notJson.push(line);
        cut = buffered.indexOf(0x0a);
        continue;
      }
      const index = waiters.findIndex((waiter) => waiter.id === frame.id);
      if (index >= 0) waiters.splice(index, 1)[0].resolve(frame);
      cut = buffered.indexOf(0x0a);
    }
  });
  child.stderr.on("data", (chunk: Buffer) => stderr.push(chunk));

  const request = (id: number, method: string, params?: unknown) =>
    new Promise<Frame>((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(
          new Error(
            `no response to ${method} (id ${id}) in 20s; stdout: ${lines.join(" | ")}; stderr: ${Buffer.concat(stderr).toString("utf8")}`,
          ),
        );
      }, 20000);
      waiters.push({
        id,
        resolve: (frame) => {
          clearTimeout(timer);
          resolve(frame);
        },
      });
      child.stdin.write(
        `${JSON.stringify({ jsonrpc: "2.0", id, method, ...(params === undefined ? {} : { params }) })}\n`,
      );
    });

  return {
    child,
    lines,
    notJson,
    stderrText: () => Buffer.concat(stderr).toString("utf8"),
    stdoutText: () => Buffer.concat(raw).toString("utf8"),
    notify: (method: string) =>
      child.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", method })}\n`),
    request,
  };
}

test("no arguments serves MCP over stdio with only newline-delimited JSON on stdout", async (t) => {
  const home = mkdtempSync(join(tmpdir(), "outlook-mcp-e2e-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const server = startServer(home);
  t.after(() => server.child.kill());

  const initialized = await server.request(1, "initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "outlook-e2e", version: "0.0.0" },
  });
  assert.equal(initialized.jsonrpc, "2.0");
  assert.equal(initialized.result?.serverInfo?.name, "outlook");
  assert.match(initialized.result?.protocolVersion ?? "", /^20\d\d-\d\d-\d\d$/);
  assert.match(
    initialized.result?.instructions ?? "",
    /send, reply or forward/,
  );
  server.notify("notifications/initialized");

  const listed = await server.request(2, "tools/list");
  assert.deepEqual(
    listed.result?.tools?.map((tool) => tool.name).sort(),
    [...readToolNames, ...draftToolNames].sort(),
  );

  const status = await server.request(3, "tools/call", {
    name: "auth_status",
    arguments: {},
  });
  assert.equal(status.result?.isError, undefined);
  assert.equal(
    JSON.parse(status.result?.content?.[0]?.text ?? "{}").signedIn,
    false,
  );

  server.child.stdin.end();
  server.child.kill();

  const text = server.stdoutText();
  assert.equal(text.length > 0, true, "the server answered");
  assert.equal(text.endsWith("\n"), true, "every frame ends with a newline");
  const frames = text.split("\n");
  assert.equal(frames.at(-1), "", "stdout stops after the last frame");
  // Byte accounting: one line per response, no preamble, no progress text.
  assert.equal(
    frames.length - 1,
    3,
    `expected three frames, got ${JSON.stringify(frames)}`,
  );
  assert.deepEqual(server.notJson, [], "stdout carries only JSON");
  for (const line of frames.slice(0, -1)) {
    assert.equal(line.startsWith("{"), true, line);
    assert.equal((JSON.parse(line) as Frame).jsonrpc, "2.0");
  }
  assert.equal(text.includes("\u0000"), false);
  // stderr is where anything chatty belongs, so it is deliberately unchecked.
  assert.equal(server.lines.length, 3);
});

test("read-only mode serves only the read tools and says so on stderr", async (t) => {
  const home = mkdtempSync(join(tmpdir(), "outlook-mcp-e2e-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const server = startServer(home, { OUTLOOK_MCP_READONLY: "1" });
  t.after(() => server.child.kill());

  const initialized = await server.request(1, "initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "outlook-e2e", version: "0.0.0" },
  });
  assert.match(initialized.result?.instructions ?? "", /readonly mode/);
  server.notify("notifications/initialized");

  const listed = await server.request(2, "tools/list");
  // Pinned against the read half of the registry, so a draft tool served in this
  // mode fails the comparison instead of needing its own assertion.
  const names = listed.result?.tools?.map((tool) => tool.name) ?? [];
  assert.deepEqual(names.sort(), [...readToolNames].sort());

  // The session reports the mode and the permissions behind it, so an agent
  // that cannot find create_draft can see why rather than guess.
  const status = await server.request(3, "tools/call", {
    name: "auth_status",
    arguments: {},
  });
  const reported = JSON.parse(status.result?.content?.[0]?.text ?? "{}");
  assert.equal(reported.readonly, true);
  assert.deepEqual(reported.scopes, [
    "User.Read",
    "Mail.Read",
    "email",
    "profile",
  ]);

  server.child.stdin.end();
  server.child.kill();

  assert.match(
    server.stderrText(),
    /Outlook MCP serving read-only \(User\.Read, Mail\.Read, email, profile\); draft tools disabled/,
  );
  // The line went to stderr, so stdout is still nothing but frames.
  assert.deepEqual(server.notJson, [], "stdout carries only JSON");
  assert.equal(server.stdoutText().includes("read-only"), false);
});
