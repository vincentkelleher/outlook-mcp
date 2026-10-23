import assert from "node:assert/strict";
import { test } from "node:test";
import { captureLogs } from "./support/log-capture.ts";
import { loadConfig } from "../src/core/config.ts";
import {
  address,
  configureLogging,
  currentLogFormat,
  currentLogLevel,
  defaultLogFormat,
  defaultLogLevel,
  errorFields,
  fieldsForLogging,
  identifier,
  logFormatFor,
  logLevelFor,
  logger,
  textOf,
  verificationCode,
} from "../src/core/log.ts";

/**
 * The logging itself. The suite runs with the level pinned to `silent` in
 * `test/setup.ts`, so each test here takes the level, the format and a sink it
 * can read — and hands them back afterwards.
 */

test("the level names are exactly the four plus silence", () => {
  for (const level of ["debug", "info", "warn", "error", "silent"] as const)
    assert.equal(logLevelFor("OUTLOOK_LOG_LEVEL", level), level);
  // Case and stray space are a person typing, not a mistake to argue with.
  assert.equal(logLevelFor("OUTLOOK_LOG_LEVEL", "  DEBUG "), "debug");
  assert.equal(logLevelFor("OUTLOOK_LOG_LEVEL", ""), defaultLogLevel);
  assert.equal(logLevelFor("OUTLOOK_LOG_LEVEL", undefined), defaultLogLevel);
  assert.equal(defaultLogLevel, "info");
});

test("an unknown level says which variable and which names are valid", () => {
  assert.throws(
    () => logLevelFor("OUTLOOK_LOG_LEVEL", "verbose"),
    (error: Error) =>
      error.message.includes("OUTLOOK_LOG_LEVEL") &&
      error.message.includes("debug, info, warn, error, silent") &&
      error.message.includes('"verbose"'),
  );
});

test("the format is text for a person, json for a collector", () => {
  assert.equal(logFormatFor("OUTLOOK_LOG_FORMAT", "JSON"), "json");
  assert.equal(logFormatFor("OUTLOOK_LOG_FORMAT", undefined), defaultLogFormat);
  assert.equal(defaultLogFormat, "text");
  assert.throws(
    () => logFormatFor("OUTLOOK_LOG_FORMAT", "yaml"),
    /OUTLOOK_LOG_FORMAT must be one of text, json/,
  );
});

test("configuration carries both, so .env reaches them too", () => {
  const config = loadConfig({
    HOME: "/tmp/home",
    OUTLOOK_LOG_LEVEL: "warn",
    OUTLOOK_LOG_FORMAT: "json",
  });
  assert.equal(config.logLevel, "warn");
  assert.equal(config.logFormat, "json");
  assert.throws(
    () => loadConfig({ HOME: "/tmp/home", OUTLOOK_LOG_LEVEL: "everything" }),
    /OUTLOOK_LOG_LEVEL must be one of/,
  );
});

test("a text line says when, how loud, who, what, and with which fields", () => {
  const out = captureLogs({ level: "info", format: "text" });
  try {
    logger("auth").info("session check confirmed a session", {
      account: "a***@example.com",
      ms: 412,
      unread: true,
    });
    const line = out.text();
    assert.match(
      line,
      /^\d{4}-\d{2}-\d{2}T[\d:.]+Z info \d+ auth: session check confirmed a session account=a\*\*\*@example\.com ms=412 unread=true\n$/,
    );
  } finally {
    out.restore();
  }
});

test("a json line is one object, with pino's own level numbers", () => {
  const out = captureLogs({ level: "info", format: "json" });
  try {
    logger("mcp").warn("tool call failed", { tool: "list_emails", ms: 7 });
    const record = JSON.parse(out.text()) as Record<string, unknown>;
    assert.equal(record.level, 40);
    assert.equal(record.scope, "mcp");
    assert.equal(record.msg, "tool call failed");
    assert.equal(record.tool, "list_emails");
    assert.equal(typeof record.time, "string");
    assert.equal(typeof record.pid, "number");
  } finally {
    out.restore();
  }
});

test("nothing a logger writes reaches stdout, in either format", () => {
  const written: string[] = [];
  const stdout = process.stdout.write;
  const out = captureLogs({ level: "debug", format: "text" });
  process.stdout.write = ((chunk: string) => {
    written.push(String(chunk));
    return true;
  }) as typeof process.stdout.write;
  try {
    for (const format of ["text", "json"] as const) {
      configureLogging({ format, force: true });
      const log = logger("graph").child({ sessionId: "session-1" });
      log.debug("graph request", { method: "GET", path: "/me/messages" });
      log.info("page read", { items: 10 });
      log.warn("silent token refresh failed", { problem: "expired" });
      log.error("request failed", { reason: "boom" });
    }
    assert.deepEqual(written, []);
    assert.ok(out.lines.length >= 8, "every line went to the sink");
  } finally {
    process.stdout.write = stdout;
    out.restore();
  }
});

test("a level below the configured one writes nothing at all", () => {
  const out = captureLogs({ level: "warn" });
  try {
    const log = logger("auth");
    log.debug("access token requested");
    log.info("session check confirmed a session");
    log.warn("silent token refresh failed");
    log.error("sign-in failed");
    assert.equal(out.lines.length, 2);
    assert.match(out.text(), / warn \d+ auth: silent token refresh failed/);
    assert.match(out.text(), / error \d+ auth: sign-in failed/);
  } finally {
    out.restore();
  }
});

test("silence is a level, and it writes nothing", () => {
  const out = captureLogs({ level: "silent" });
  try {
    logger("cli").error("outlook stopped with an error");
    assert.deepEqual(out.lines, []);
    assert.equal(currentLogLevel(), "silent");
  } finally {
    out.restore();
  }
});

test("a value that must never be logged is gone from both formats", () => {
  const secret = "EwBkA1q8jt6Ctoken-material";
  for (const format of ["text", "json"] as const) {
    const out = captureLogs({ level: "debug", format });
    try {
      logger("auth").debug("token cache rewritten", {
        cachePath: "/home/ana/.outlook-skill/msal-cache.json",
        accessToken: secret,
        refresh_token: secret,
        authorization: `Bearer ${secret}`,
        contentBase64: "c2VjcmV0IG1haWw=",
        body: "the whole message",
        nested: { untouched: true },
      });
      const line = out.text();
      assert.ok(!line.includes(secret), `${format} leaked the token`);
      assert.ok(!line.includes("c2VjcmV0IG1haWw="), `${format} leaked content`);
      assert.ok(!line.includes("the whole message"), `${format} leaked a body`);
      // The line is still useful: what happened and where it happened survive.
      assert.match(line, /token cache rewritten/);
      assert.match(line, /msal-cache\.json/);
    } finally {
      out.restore();
    }
  }
});

test("at the default level an id is shortened, a code masked, a body counted", () => {
  const messageId = "AAMkAGI2TGV0MDE2OTY4NTUwMDAwAAAAAAAA=";
  assert.equal(identifier(messageId, true), messageId);
  assert.equal(identifier(messageId, false), `${messageId.slice(0, 8)}…37`);
  assert.equal(identifier("short", false), "short");
  assert.equal(
    address("vincent.kelleher@example.com", false),
    "v***@example.com",
  );
  assert.equal(address("ana@example.com", true), "ana@example.com");
  assert.equal(address(undefined, false), undefined);
  assert.equal(verificationCode("HKDFQPGB", false), "HKDF***");
  assert.equal(verificationCode("HKDFQPGB", true), "HKDFQPGB");
});

test("tool arguments read as what was asked, not as what the mailbox holds", () => {
  const quiet = fieldsForLogging(
    {
      query: 'from:"ana" subject:"quarterly report"',
      messageId: "AAMkAGI2TGV0MDE2OTY4NTUwMDAwAAAAAAAA=",
      to: ["ana@example.com", "bob@example.com"],
      body: "the message text",
      contentBase64: "c2VjcmV0",
      folder: "inbox",
      unread: true,
      limit: 25,
      cursor: undefined,
    },
    false,
  );
  assert.deepEqual(quiet, {
    queryChars: 37,
    messageId: "AAMkAGI2…37",
    toCount: 2,
    bodyChars: 16,
    contentBase64Chars: 8,
    folder: "inbox",
    unread: true,
    limit: 25,
  });

  const loud = fieldsForLogging(
    { query: "invoice", messageId: "AAMkAGI2", to: ["ana@example.com"] },
    true,
  );
  assert.deepEqual(loud, {
    query: "invoice",
    messageId: "AAMkAGI2",
    to: ["ana@example.com"],
  });
  // Not an object at all — a tool called with nothing, or a malformed client.
  assert.deepEqual(fieldsForLogging(undefined, false), {});
});

test("an error gives its reason always and its stack only under debug", () => {
  const cause = new Error("getaddrinfo ENOTFOUND login.microsoftonline.com");
  const failure = new Error("Could not confirm the Outlook session", {
    cause,
  });
  failure.name = "AuthRequiredError";

  const quiet = errorFields(failure, false);
  assert.equal(quiet.reason, "Could not confirm the Outlook session");
  assert.equal(quiet.errorName, "AuthRequiredError");
  assert.equal(quiet.stack, undefined);
  assert.equal(quiet.causedBy, undefined);

  const loud = errorFields(failure, true);
  assert.match(String(loud.stack), /AuthRequiredError/);
  assert.equal(
    loud.causedBy,
    "Error: getaddrinfo ENOTFOUND login.microsoftonline.com",
  );

  assert.deepEqual(errorFields("a string thrown", false), {
    reason: "a string thrown",
  });
});

test("a child line carries what its parent was bound to", () => {
  const out = captureLogs({ level: "info", format: "text" });
  try {
    const session = logger("http").child({ sessionId: "0f0f0f0f" });
    session.info("session opened", { sessions: 1 });
    session.child({ tool: "search_emails" }).info("tool call", {
      queryChars: 11,
    });
    const lines = out.text();
    // A binding belongs to the logger, so it precedes whatever the call added.
    assert.match(lines, /http: session opened sessionId=0f0f0f0f sessions=1/);
    assert.match(
      lines,
      /http: tool call sessionId=0f0f0f0f tool=search_emails queryChars=11/,
    );
  } finally {
    out.restore();
  }
});

test("a record that will not parse is written as it arrived", () => {
  // The renderer is the last thing in front of a diagnostic; a line it cannot
  // read is a line it must still pass on.
  assert.equal(textOf("not json at all\n"), "not json at all\n");
});

test("the level and format in force are readable, for status and for tests", () => {
  const out = captureLogs({ level: "debug", format: "json" });
  try {
    assert.equal(currentLogLevel(), "debug");
    assert.equal(currentLogFormat(), "json");
    assert.equal(logger("graph").detailed, true);
    configureLogging({ level: "warn", force: true });
    assert.equal(logger("graph").detailed, false);
  } finally {
    out.restore();
  }
});
