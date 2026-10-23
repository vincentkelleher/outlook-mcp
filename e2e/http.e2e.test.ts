import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import {
  Client,
  StreamableHTTPClientTransport,
} from "@modelcontextprotocol/client";
import { readToolNames } from "../src/tools.ts";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const binArgs = (process.env.OUTLOOK_E2E_BIN ?? "src/index.ts").split(" ");

for (const signal of ["SIGINT", "SIGTERM"] as const) {
  test(
    `HTTP CLI serves /mcp and shuts down on ${signal}`,
    { timeout: 15000 },
    async (t) => {
      const home = mkdtempSync(join(tmpdir(), "outlook-http-e2e-"));
      const child = spawn(
        process.execPath,
        [...binArgs, "serve", "--transport", "http", "--port", "0"],
        {
          cwd: repo,
          env: {
            ...process.env,
            HOME: home,
            OUTLOOK_SKILL_HOME: join(home, "session"),
            OUTLOOK_MCP_STATE_DIR: join(home, "state"),
            OUTLOOK_OAUTH_CLIENT_ID: "00000000-0000-0000-0000-000000000000",
            OUTLOOK_MCP_READONLY: "1",
          },
          stdio: ["pipe", "pipe", "pipe"],
        },
      );
      t.after(() => {
        child.kill("SIGKILL");
        rmSync(home, { recursive: true, force: true });
      });
      let stdout = "";
      let stderr = "";
      child.stdout.on("data", (chunk) => {
        stdout += chunk;
      });
      const url = await new Promise<string>((resolve, reject) => {
        child.once("error", reject);
        child.once("exit", (code) =>
          reject(new Error(`HTTP exited (${code}): ${stderr}`)),
        );
        child.stderr.on("data", (chunk) => {
          stderr += chunk;
          const match = /listening at (http:\/\/127\.0\.0\.1:\d+\/mcp)/.exec(
            stderr,
          );
          if (match) resolve(match[1]);
        });
      });
      const client = new Client({ name: "http-e2e", version: "0.0.0" });
      t.after(() => client.close());
      await client.connect(new StreamableHTTPClientTransport(new URL(url)));
      assert.deepEqual(
        (await client.listTools()).tools.map((tool) => tool.name).sort(),
        [...readToolNames].sort(),
      );
      const status = await client.callTool({
        name: "auth_status",
        arguments: {},
      });
      assert.equal(status.isError, undefined);
      const payload = status.structuredContent as Record<string, unknown>;
      assert.equal(payload.signedIn, false);
      assert.equal(payload.readonly, true);
      const blocked = await fetch(url, {
        headers: { Origin: "https://evil.example" },
      });
      assert.equal(blocked.status, 403);
      await blocked.text();
      await client.close();
      const exited = once(child, "exit");
      child.kill(signal);
      assert.deepEqual(await exited, [0, null]);
      assert.equal(
        stdout,
        "",
        "HTTP startup and diagnostics never write to stdout",
      );
      await assert.rejects(fetch(url));
    },
  );
}
