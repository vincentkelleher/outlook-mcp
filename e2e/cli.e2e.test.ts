import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
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
  assert.match(root.stdout, /config \[options\] <client>/);
  // Drafts are MCP-only on purpose: a human writes them in Outlook, not here.
  assert.equal(
    `${root.stdout}${emails.stdout}`.toLowerCase().includes("draft"),
    false,
    "the CLI offers no draft command",
  );
});

test("config prints the shape a client actually reads", (t) => {
  const home = testHome(t);
  const pi = outlook(home, "config", "pi");
  const codex = outlook(home, "config", "codex");

  assert.equal(pi.exitCode, 0, pi.stderr);
  const server = JSON.parse(pi.stdout).mcpServers.outlook;
  assert.equal(server.command, "node");
  // Pi defaults MCP tools to codemode, so the entry has to say otherwise.
  assert.equal(server.exposure, "direct");
  assert.equal(
    server.env.OUTLOOK_OAUTH_CLIENT_ID,
    "${OUTLOOK_OAUTH_CLIENT_ID}",
  );
  // An absolute path into this checkout, built or not, never a relative guess.
  assert.match(
    server.args[0],
    /outlook-skill\/(dist\/index\.js|src\/index\.ts)$/,
    server.args[0],
  );

  assert.equal(codex.exitCode, 0, codex.stderr);
  assert.match(codex.stdout, /^\[mcp_servers\.outlook\]/m);
  // Codex documents no interpolation, so the variable is forwarded instead.
  assert.match(codex.stdout, /env_vars = \["OUTLOOK_OAUTH_CLIENT_ID"\]/);
  assert.equal(codex.stdout.includes("${"), false);
});

test("config --write merges into the user file, repeats itself, and refuses garbage", (t) => {
  const home = testHome(t);
  const file = join(home, ".codex", "config.toml");
  mkdirSync(dirname(file), { recursive: true });
  writeFileSync(file, '[mcp_servers.other]\ncommand = "y"\n');

  const first = outlook(home, "config", "codex", "--write");
  assert.equal(first.exitCode, 0, first.stderr);
  assert.match(first.stdout, /updated with the outlook server/);
  const written = readFileSync(file, "utf8");
  assert.match(written, /\[mcp_servers\.other\]/);
  assert.match(written, /\[mcp_servers\.outlook\]/);

  const second = outlook(home, "config", "codex", "--write");
  assert.equal(second.exitCode, 0, second.stderr);
  assert.match(second.stdout, /already had this outlook server/);
  assert.equal(readFileSync(file, "utf8"), written);

  writeFileSync(file, "not toml at all\n");
  const refused = outlook(home, "config", "codex", "--write");
  assert.equal(refused.exitCode, 1);
  assert.match(
    refused.stderr,
    /never overwrites a configuration file it cannot parse/,
  );
  assert.equal(readFileSync(file, "utf8"), "not toml at all\n");
});

test("config names every client when one is misspelt", (t) => {
  const home = testHome(t);
  const result = outlook(home, "config", "notepad");
  assert.equal(result.exitCode, 1);
  assert.match(result.stderr, /Unknown client "notepad"/);
  assert.match(result.stderr, /gemini-cli/);
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
