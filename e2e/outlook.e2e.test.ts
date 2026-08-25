import { expect, test } from "bun:test";
import {
  existsSync,
  mkdtempSync,
  mkdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

async function outlook(home: string, ...args: string[]) {
  const child = Bun.spawn(["bun", "scripts/outlook.ts", ...args], {
    stdout: "pipe",
    stderr: "pipe",
    env: {
      ...process.env,
      HOME: home,
      OUTLOOK_OAUTH_CLIENT_ID: "00000000-0000-0000-0000-000000000000",
    },
  });
  return {
    exitCode: await child.exited,
    stdout: await new Response(child.stdout).text(),
    stderr: await new Response(child.stderr).text(),
  };
}

function testHome() {
  return mkdtempSync(join(tmpdir(), "outlook-skill-e2e-"));
}

test("CLI exposes every documented command", async () => {
  const home = testHome();
  const root = await outlook(home, "--help");
  const emails = await outlook(home, "emails", "--help");

  expect(root.exitCode).toBe(0);
  expect(root.stdout).toContain("login");
  expect(root.stdout).toContain("logout");
  expect(root.stdout).toContain("emails");
  expect(emails.exitCode).toBe(0);
  expect(emails.stdout).toContain("list");
  expect(emails.stdout).toContain("get <id>");
  expect(emails.stdout).toContain("attachments <message-id>");
  expect(emails.stdout).toContain("search [options] <query>");
});

test("logout command removes the saved session", async () => {
  const home = testHome();
  const cache = join(home, ".outlook-skill", "msal-cache.json");
  mkdirSync(join(home, ".outlook-skill"), { recursive: true });
  writeFileSync(cache, "session");

  const result = await outlook(home, "logout");

  expect(result.exitCode).toBe(0);
  expect(result.stderr).toContain("Signed out.");
  expect(existsSync(cache)).toBe(false);
  rmSync(home, { recursive: true, force: true });
});

test("email commands reject invalid limits without contacting Microsoft Graph", async () => {
  const home = testHome();
  for (const args of [
    ["emails", "list", "--limit", "0"],
    ["emails", "list", "--limit", "101"],
    ["emails", "search", "test", "--limit", "0"],
    ["emails", "search", "test", "--limit", "1001"],
  ]) {
    const result = await outlook(home, ...args);
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain("--limit must be an integer");
  }
});

test("email commands report a missing session", async () => {
  const home = testHome();
  for (const args of [
    ["emails", "list"],
    ["emails", "get", "message-id"],
    ["emails", "attachments", "message-id"],
    ["emails", "search", "test"],
  ]) {
    const result = await outlook(home, ...args);
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain(
      "No Outlook session. Run the login command first.",
    );
  }
});
