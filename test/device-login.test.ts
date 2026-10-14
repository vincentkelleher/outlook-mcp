import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { loadConfig, loginStatePath } from "../src/core/config.ts";
import type { Config } from "../src/core/config.ts";
import {
  beginLogin,
  readLoginState,
  runLogin,
  startLogin,
  writeLoginState,
} from "../src/core/device-login.ts";

const clientId = "00000000-0000-0000-0000-000000000000";
const verificationUri = "https://microsoft.com/devicelogin";

function testConfig(t: { after: (fn: () => void) => void }): Config {
  const home = mkdtempSync(join(tmpdir(), "outlook-login-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  return loadConfig({
    HOME: home,
    OUTLOOK_OAUTH_CLIENT_ID: clientId,
    OUTLOOK_MCP_STATE_DIR: join(home, "state"),
  });
}

const rawState = (config: Config) =>
  JSON.parse(readFileSync(loginStatePath(config), "utf8")) as {
    state?: string;
    pid?: number;
    expiresAt?: string;
    userCode?: string;
  };

/** A process number that is certainly gone again: a child run to completion. */
function deadPid(): number {
  const finished = spawnSync(process.execPath, ["-e", ""], {
    stdio: "ignore",
  });
  assert.throws(() => process.kill(finished.pid, 0), { code: "ESRCH" });
  return finished.pid;
}

/** A process that is certainly alive and certainly not this one. */
function liveForeignPid(): { pid: number; stop: () => void } {
  const sleeper = spawn(
    process.execPath,
    ["-e", "setTimeout(() => {}, 30000)"],
    { stdio: "ignore" },
  );
  const pid = sleeper.pid;
  assert.ok(pid !== undefined);
  assert.doesNotThrow(() => process.kill(pid, 0));
  return { pid, stop: () => sleeper.kill() };
}

/**
 * A fake launcher that records what it was asked to run and reports the
 * caller's own pid, so the liveness check has a living process to find. `fail`
 * stands in for a spawn that never happened, which a real child reports on the
 * event loop rather than by throwing.
 */
function fakeSpawn() {
  const launched: string[][] = [];
  const listeners: ((error: Error) => void)[] = [];
  return {
    launched,
    fail: (error: Error) => {
      for (const listener of listeners.splice(0)) listener(error);
    },
    spawnChild: (command: string, args: string[]) => {
      launched.push([command, ...args]);
      return {
        pid: process.pid,
        unref: () => {},
        on: (event: "error", listener: (error: unknown) => unknown) => {
          if (event === "error")
            listeners.push(listener as (error: Error) => void);
        },
      };
    },
  };
}

const oneLaunch = (entry: string) => [[process.execPath, entry, "login"]];

test("a pending sign-in whose deadline has passed reads as no sign-in at all", (t) => {
  const config = testConfig(t);

  // A child that died after publishing the code leaves this behind, and leaving
  // it standing makes every client tell the user to enter a dead code for ever.
  writeLoginState(config, {
    state: "pending",
    verificationUri,
    userCode: "DEAD-CODE",
    expiresAt: new Date(Date.now() - 1000).toISOString(),
  });
  assert.equal(readLoginState(config), undefined);

  writeLoginState(config, {
    state: "pending",
    verificationUri,
    userCode: "LIVE-CODE",
    expiresAt: new Date(Date.now() + 600_000).toISOString(),
  });
  assert.equal(readLoginState(config)?.userCode, "LIVE-CODE");

  // An outcome is not a flow, so it carries no deadline and outlives none.
  writeLoginState(config, { state: "error", message: "no client id" });
  assert.equal(readLoginState(config)?.state, "error");
  writeLoginState(config, { state: "authorized" });
  assert.equal(readLoginState(config)?.state, "authorized");
});

test("sign-in that starts again publishes a new code, not the expired one", async (t) => {
  const config = testConfig(t);
  writeLoginState(config, {
    state: "pending",
    verificationUri,
    userCode: "DEAD-CODE",
    expiresAt: new Date(Date.now() - 60_000).toISOString(),
  });

  let timer: ReturnType<typeof setTimeout> | undefined;
  t.after(() => clearTimeout(timer));
  const launch = await beginLogin(config, (started) => {
    // Stands in for the child reaching Microsoft: the state is replaced by a
    // flow the user can still complete.
    timer = setTimeout(
      () =>
        writeLoginState(started, {
          state: "pending",
          verificationUri,
          userCode: "FRESH-CODE",
          expiresAt: new Date(Date.now() + 600_000).toISOString(),
        }),
      120,
    );
  });

  assert.equal(launch.userCode, "FRESH-CODE");
  assert.match(launch.instructions, /FRESH-CODE/);
  assert.equal(launch.instructions.includes("DEAD-CODE"), false);
});

test("startLogin records the flow it started and will not start a second one", (t) => {
  const config = testConfig(t);
  const { launched, spawnChild } = fakeSpawn();
  const deps = { spawnChild, entry: join("some", "index.js") };

  startLogin(config, deps);

  assert.deepEqual(launched, oneLaunch(join("some", "index.js")));
  // The state names a living process and a deadline, which is what makes a
  // second caller's `login` a poll instead of a competing device-code request.
  const lease = rawState(config);
  assert.equal(lease.state, "pending");
  assert.equal(lease.pid, process.pid);
  assert.equal(Date.parse(lease.expiresAt ?? "") > Date.now(), true);

  // The child publishes its code; another client calls login while it is open.
  writeLoginState(config, {
    state: "pending",
    pid: process.pid,
    verificationUri,
    userCode: "ABC-XYZ",
    expiresAt: new Date(Date.now() + 600_000).toISOString(),
  });
  startLogin(config, deps);

  assert.deepEqual(
    launched,
    oneLaunch(join("some", "index.js")),
    "a second flow was started",
  );
  assert.equal(
    rawState(config).userCode,
    "ABC-XYZ",
    "the open flow was replaced, which is how its result was lost before",
  );
});

test("a sign-in that never started says so instead of looking slow", async (t) => {
  const config = testConfig(t);
  const { spawnChild, fail } = fakeSpawn();
  const start = (started: Config) =>
    startLogin(started, { spawnChild, entry: join("some", "index.js") });

  const launch = beginLogin(config, start);
  // The failure arrives after the launch returned, as a real one does.
  await new Promise((resolve) => setTimeout(resolve, 0));
  fail(new Error("spawn /nope/node EACCES"));

  const reported = await launch;
  assert.equal(reported.state, "error");
  assert.match(
    reported.instructions,
    /Sign-in could not start \(spawn \/nope\/node EACCES\)/,
  );
  assert.equal(readLoginState(config)?.message, "spawn /nope/node EACCES");
});

test("the state directory and its file are owner-only", (t) => {
  const config = testConfig(t);
  // A directory that already exists keeps whatever mode it was made with, and
  // mkdir's own mode is filtered by the umask, so both need stating again.
  mkdirSync(config.stateDir, { recursive: true, mode: 0o755 });
  writeLoginState(config, { state: "pending", userCode: "ABC-XYZ" });

  assert.equal(statSync(config.stateDir).mode & 0o777, 0o700);
  assert.equal(statSync(loginStatePath(config)).mode & 0o777, 0o600);
});

test("a flow no process owns does not block sign-in", (t) => {
  const config = testConfig(t);
  const future = new Date(Date.now() + 600_000).toISOString();

  // Owned by a process that is gone: nothing is polling Microsoft for this one.
  writeLoginState(config, {
    state: "pending",
    pid: deadPid(),
    verificationUri,
    userCode: "ORPHAN-CODE",
    expiresAt: future,
  });
  const { launched, spawnChild } = fakeSpawn();
  startLogin(config, { spawnChild, entry: "index.js" });

  assert.deepEqual(
    launched,
    oneLaunch("index.js"),
    "an orphaned flow was guarded",
  );

  // A pending state that names no owner, as an older version left behind, has
  // nothing to prove either, so it does not hold sign-in hostage.
  writeLoginState(config, {
    state: "pending",
    verificationUri,
    userCode: "OWNERLESS",
    expiresAt: future,
  });
  startLogin(config, { spawnChild, entry: "index.js" });

  assert.equal(launched.length, 2);
});

test("two clients that disagree about readonly run two flows", (t) => {
  // One machine, one state directory, two servers: a read-only client and a
  // read-write client ask Microsoft for different scopes, so the code one of
  // them publishes cannot answer for the other. Reusing it would put the user
  // through a consent that leaves the second server still asking for login.
  const home = mkdtempSync(join(tmpdir(), "outlook-login-scopes-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const env = { HOME: home, OUTLOOK_OAUTH_CLIENT_ID: clientId };
  const readOnly = loadConfig({ ...env, OUTLOOK_MCP_READONLY: "1" });
  const readWrite = loadConfig({ ...env, OUTLOOK_MCP_READONLY: "0" });

  assert.notEqual(
    loginStatePath(readOnly),
    loginStatePath(readWrite),
    "a flow for one scope set is a different file from a flow for the other",
  );

  const first = fakeSpawn();
  startLogin(readOnly, {
    spawnChild: first.spawnChild,
    entry: "index.js",
  });
  writeLoginState(readOnly, {
    state: "pending",
    pid: process.pid,
    verificationUri,
    userCode: "READ-ONLY-CODE",
    expiresAt: new Date(Date.now() + 600_000).toISOString(),
  });

  // The read-write server must not be told "already starting" with a code that
  // cannot grant what it needs, and must not overwrite the code its own
  // user may already be typing into the other window.
  const second = fakeSpawn();
  startLogin(readWrite, {
    spawnChild: second.spawnChild,
    entry: "index.js",
  });

  assert.equal(
    second.launched.length,
    1,
    "the second scope set started its own",
  );
  assert.equal(readLoginState(readWrite)?.userCode, undefined);
  assert.equal(readLoginState(readOnly)?.userCode, "READ-ONLY-CODE");

  // Same scope set, live flow: now it is genuinely the same flow, so no second
  // request goes to Microsoft and the published code is handed back.
  startLogin(readWrite, {
    spawnChild: second.spawnChild,
    entry: "index.js",
  });
  assert.equal(second.launched.length, 1, "one flow per scope set");
});

test("an outcome left in the state file does not block a new flow", (t) => {
  // The claim on the state file is now exclusive, which must not mean that a
  // finished attempt - or a flow whose process died - holds sign-in for ever.
  const config = testConfig(t);
  const future = new Date(Date.now() + 600_000).toISOString();
  const { spawnChild, launched } = fakeSpawn();

  for (const leftover of [
    { state: "error", message: "the last attempt failed" },
    { state: "authorized" },
    {
      state: "pending",
      pid: deadPid(),
      verificationUri,
      userCode: "DEAD-FLOW",
      expiresAt: future,
    },
  ] as const) {
    writeLoginState(config, leftover as never);
    startLogin(config, { spawnChild, entry: "index.js" });
  }

  assert.equal(launched.length, 3, "each leftover was claimed past");
});

test("a login in a terminal does not hijack another process's code", async (t) => {
  const config = testConfig(t);
  const future = new Date(Date.now() + 600_000).toISOString();

  // Somebody else's live flow: refusing here is the whole point, because taking
  // the state would replace a code the user may already be typing.
  const other = liveForeignPid();
  t.after(() => other.stop());
  writeLoginState(config, {
    state: "pending",
    pid: other.pid,
    verificationUri,
    userCode: "THEIR-CODE",
    expiresAt: future,
  });

  const stderr = process.stderr.write;
  const written: string[] = [];
  process.stderr.write = (chunk: string | Uint8Array) => {
    written.push(chunk.toString());
    return true;
  };
  t.after(() => {
    process.stderr.write = stderr;
  });

  await runLogin(config);

  process.stderr.write = stderr;
  assert.equal(
    readLoginState(config)?.userCode,
    "THEIR-CODE",
    "the other flow's code survived",
  );
  assert.match(written.join(""), /already running as process/);

  // The detached child that owns its own lease is still allowed to take it,
  // which is the normal path and must not be blocked by the guard above. No
  // client id is configured here, so sign-in fails on the way to Microsoft
  // rather than reaching Microsoft: a unit test has no business dialling out.
  const ownedHome = mkdtempSync(join(tmpdir(), "outlook-login-owner-"));
  t.after(() => rmSync(ownedHome, { recursive: true, force: true }));
  const owned = loadConfig({
    HOME: ownedHome,
    OUTLOOK_MCP_STATE_DIR: join(ownedHome, "state"),
  });
  writeLoginState(owned, {
    state: "pending",
    pid: process.pid,
    verificationUri,
    userCode: "OUR-CODE",
    expiresAt: future,
  });

  const stderr2 = process.stderr.write;
  const again: string[] = [];
  process.stderr.write = (chunk: string | Uint8Array) => {
    again.push(chunk.toString());
    return true;
  };
  // Sign-in reports failure by rethrowing, so the CLI exits non-zero; the guard
  // is what is under test, not the reason sign-in could not proceed.
  let thrown: Error | undefined;
  try {
    await runLogin(owned);
  } catch (error) {
    thrown = error as Error;
  } finally {
    process.stderr.write = stderr2;
  }

  assert.doesNotMatch(
    again.join(""),
    /already running as process/,
    "the owner of a live flow was refused by the guard",
  );
  assert.match(
    thrown?.message ?? "",
    /OUTLOOK_OAUTH_CLIENT_ID/,
    "sign-in stopped on the stated missing client id, not on the guard",
  );
});
