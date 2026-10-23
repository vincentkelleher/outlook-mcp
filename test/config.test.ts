import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { loadConfig, loadEnvFile } from "../src/core/config.ts";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));

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
      "  export OUTLOOK_MCP_ATTACHMENT_DIR = /srv/attachments ",
      "OUTLOOK_MCP_READONLY=1",
      "",
      "not a pair",
    ].join("\n"),
  );

  const env = loadEnvFile(dir, { OUTLOOK_MCP_STATE_DIR: "/srv/state" });
  assert.equal(env.OUTLOOK_OAUTH_CLIENT_ID, clientId);
  assert.equal(env.OUTLOOK_OAUTH_TENANT, "organizations");
  assert.equal(env.OUTLOOK_MCP_ATTACHMENT_DIR, "/srv/attachments");
  assert.equal(env.OUTLOOK_MCP_READONLY, "1");

  const config = loadConfig(env);
  assert.equal(config.clientId, clientId);
  assert.equal(
    config.authority,
    "https://login.microsoftonline.com/organizations/",
  );
  assert.equal(config.stateDir, "/srv/state");
  assert.equal(config.attachmentDir, "/srv/attachments");
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

  const moved = loadConfig({
    HOME: "/home/someone",
    OUTLOOK_SKILL_HOME: "/srv/outlook",
    OUTLOOK_MCP_STATE_DIR: "/srv/state",
    OUTLOOK_MCP_ATTACHMENT_DIR: "/srv/attachments",
  });
  assert.equal(moved.cachePath, "/srv/outlook/msal-cache.json");
  assert.equal(moved.stateDir, "/srv/state");
  assert.equal(moved.attachmentDir, "/srv/attachments");
});

test("the HTTP surface has a home, a port, and no invented knobs", () => {
  const plain = loadConfig({ HOME: "/home/someone" });
  // Loopback by default: nothing on that port authenticates.
  assert.equal(plain.httpHost, "127.0.0.1");
  assert.equal(plain.httpPort, 3939);
  assert.deepEqual(plain.allowedHosts, []);
  assert.deepEqual(plain.allowedOrigins, []);
  assert.equal(plain.sessionIdleMinutes, 24 * 60);

  const moved = loadConfig({
    HOME: "/home/someone",
    OUTLOOK_MCP_HOST: "0.0.0.0",
    OUTLOOK_MCP_PORT: "8080",
    OUTLOOK_MCP_ALLOWED_HOSTS: " mail.local , outlook.internal ,,",
    OUTLOOK_MCP_ALLOWED_ORIGINS: "http://Localhost:5173/",
    OUTLOOK_MCP_SESSION_IDLE_MINUTES: "30",
  });
  assert.equal(moved.httpHost, "0.0.0.0");
  assert.equal(moved.httpPort, 8080);
  assert.deepEqual(moved.allowedHosts, ["mail.local", "outlook.internal"]);
  assert.deepEqual(moved.allowedOrigins, ["http://localhost:5173"]);
  assert.equal(moved.sessionIdleMinutes, 30);
});

test("a port that is not a port is refused before anything binds", () => {
  // `server.listen(NaN)` would blame the range, not the variable.
  assert.throws(
    () => loadConfig({ HOME: "/home/someone", OUTLOOK_MCP_PORT: "393a" }),
    /OUTLOOK_MCP_PORT must be an integer from 1 to 65535, not "393a"\./,
  );
  assert.throws(
    () => loadConfig({ HOME: "/home/someone", OUTLOOK_MCP_PORT: "0" }),
    /OUTLOOK_MCP_PORT/,
  );
  assert.throws(
    () =>
      loadConfig({
        HOME: "/home/someone",
        OUTLOOK_MCP_SESSION_IDLE_MINUTES: "x",
      }),
    /OUTLOOK_MCP_SESSION_IDLE_MINUTES must be an integer/,
  );
  // An empty value is "unset", not an error: that is how an unset variable arrives.
  assert.equal(
    loadConfig({ HOME: "/home/someone", OUTLOOK_MCP_PORT: "" }).httpPort,
    3939,
  );
});

test("a log knob is either read by code or not advertised", () => {
  // This test exists because OUTLOOK_MCP_LOG_LEVEL once set a knob no code
  // looked at. There is a logger now, and it answers to OUTLOOK_LOG_LEVEL — so
  // the dead name still must not be documented, still must not be carried, and
  // must not silently be the same control under a second name.
  const readme = readFileSync(join(repo, "README.md"), "utf8");
  assert.equal(readme.includes("OUTLOOK_MCP_LOG_LEVEL"), false);
  assert.equal(
    loadConfig({ HOME: "/home/someone", OUTLOOK_MCP_LOG_LEVEL: "debug" })
      .logLevel,
    "info",
  );

  // The live pair is read, and the README that advertises it names both.
  const config = loadConfig({
    HOME: "/home/someone",
    OUTLOOK_LOG_LEVEL: "debug",
    OUTLOOK_LOG_FORMAT: "json",
  });
  assert.equal(config.logLevel, "debug");
  assert.equal(config.logFormat, "json");
  assert.match(readme, /OUTLOOK_LOG_LEVEL/);
  assert.match(readme, /OUTLOOK_LOG_FORMAT/);
});
