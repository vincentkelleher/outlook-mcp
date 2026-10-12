import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { loadConfig, loadEnvFile } from "../src/core/config.ts";

const clientId = "00000000-0000-0000-0000-000000000000";

test("the .env file fills in what the environment is missing", (t) => {
  const dir = mkdtempSync(join(tmpdir(), "outlook-skill-config-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  writeFileSync(
    join(dir, ".env"),
    [
      "# a comment",
      `OUTLOOK_OAUTH_CLIENT_ID=${clientId}`,
      'OUTLOOK_OAUTH_TENANT="organizations"',
      "  export OUTLOOK_MCP_LOG_LEVEL = DEBUG ",
      "OUTLOOK_MCP_READONLY=1",
      "",
      "not a pair",
    ].join("\n"),
  );

  const env = loadEnvFile(dir, { OUTLOOK_MCP_STATE_DIR: "/srv/state" });
  assert.equal(env.OUTLOOK_OAUTH_CLIENT_ID, clientId);
  assert.equal(env.OUTLOOK_OAUTH_TENANT, "organizations");
  assert.equal(env.OUTLOOK_MCP_LOG_LEVEL, "DEBUG");
  assert.equal(env.OUTLOOK_MCP_READONLY, "1");

  const config = loadConfig(env);
  assert.equal(config.clientId, clientId);
  assert.equal(
    config.authority,
    "https://login.microsoftonline.com/organizations/",
  );
  assert.equal(config.stateDir, "/srv/state");
  assert.equal(config.logLevel, "debug");
  assert.equal(config.readonly, true);

  // A variable already set in the environment is never replaced by the file.
  assert.equal(
    loadEnvFile(dir, { OUTLOOK_OAUTH_TENANT: "common" }).OUTLOOK_OAUTH_TENANT,
    "common",
  );
});

test("defaults keep an existing install on its current cache", () => {
  const config = loadConfig({ HOME: "/home/someone" });
  assert.equal(config.clientId, undefined);
  assert.equal(
    config.authority,
    "https://login.microsoftonline.com/consumers/",
  );
  assert.equal(
    config.cachePath,
    "/home/someone/.outlook-skill/msal-cache.json",
  );
  assert.equal(config.stateDir, "/home/someone/.outlook-skill/state");
  assert.equal(config.attachmentDir, join(tmpdir(), "outlook-mcp"));
  assert.equal(config.readonly, false);
  assert.equal(config.logLevel, "info");

  const moved = loadConfig({
    HOME: "/home/someone",
    OUTLOOK_SKILL_HOME: "/srv/outlook",
    OUTLOOK_MCP_STATE_DIR: "/srv/state",
    OUTLOOK_MCP_ATTACHMENT_DIR: "/srv/attachments",
    OUTLOOK_MCP_LOG_LEVEL: "not-a-level",
  });
  assert.equal(moved.cachePath, "/srv/outlook/msal-cache.json");
  assert.equal(moved.stateDir, "/srv/state");
  assert.equal(moved.attachmentDir, "/srv/attachments");
  assert.equal(moved.logLevel, "info");
});
