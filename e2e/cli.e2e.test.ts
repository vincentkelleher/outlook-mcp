import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const binArgs = (process.env.OUTLOOK_E2E_BIN ?? "src/index.ts").split(" ");
const placeholderClientId = "00000000-0000-0000-0000-000000000000";

function outlook(home: string, ...args: string[]) {
  const child = spawnSync(process.execPath, [...binArgs, ...args], {
    cwd: repo,
    encoding: "utf8",
    env: {
      ...process.env,
      HOME: home,
      OUTLOOK_OAUTH_CLIENT_ID: placeholderClientId,
    },
  });
  return { exitCode: child.status, stdout: child.stdout, stderr: child.stderr };
}

function testHome(t: { after: (fn: () => void) => void }) {
  const home = mkdtempSync(join(tmpdir(), "outlook-skill-e2e-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  return home;
}

test("CLI exposes every documented command", (t) => {
  const home = testHome(t);
  const root = outlook(home, "--help");
  const emails = outlook(home, "emails", "--help");

  assert.equal(root.exitCode, 0);
  assert.match(root.stdout, /login/);
  assert.match(root.stdout, /logout/);
  assert.match(root.stdout, /emails/);
  assert.equal(emails.exitCode, 0);
  assert.match(emails.stdout, /list/);
  assert.match(emails.stdout, /get <id>/);
  assert.match(emails.stdout, /attachments <message-id>/);
  assert.match(emails.stdout, /search \[options\] <query>/);
  // Drafts are MCP-only on purpose: a human writes them in Outlook, not here.
  assert.equal(
    `${root.stdout}${emails.stdout}`.toLowerCase().includes("draft"),
    false,
    "the CLI offers no draft command",
  );
});

test("logout command removes the saved session", (t) => {
  const home = testHome(t);
  const cache = join(home, ".outlook-skill", "msal-cache.json");
  mkdirSync(join(home, ".outlook-skill"), { recursive: true });
  writeFileSync(cache, "session");

  const result = outlook(home, "logout");

  assert.equal(result.exitCode, 0);
  assert.match(result.stderr, /Signed out\./);
  assert.equal(existsSync(cache), false);
});

test("email commands reject invalid limits without contacting Microsoft Graph", (t) => {
  const home = testHome(t);
  for (const args of [
    ["emails", "list", "--limit", "0"],
    ["emails", "list", "--limit", "101"],
    ["emails", "search", "test", "--limit", "0"],
    ["emails", "search", "test", "--limit", "1001"],
  ]) {
    const result = outlook(home, ...args);
    assert.equal(result.exitCode, 1, `${args.join(" ")}: ${result.stderr}`);
    assert.match(result.stderr, /--limit must be an integer/);
    assert.equal(
      result.stderr.includes("No Outlook session"),
      false,
      `${args.join(" ")} reached Microsoft Graph`,
    );
  }
});

test("email commands report a missing session", (t) => {
  const home = testHome(t);
  for (const args of [
    ["emails", "list"],
    ["emails", "get", "message-id"],
    ["emails", "attachments", "message-id"],
    ["emails", "search", "test"],
  ]) {
    const result = outlook(home, ...args);
    assert.equal(result.exitCode, 1, `${args.join(" ")}: ${result.stderr}`);
    assert.match(
      result.stderr,
      /No Outlook session\. Run the login command first\./,
    );
  }
});
