import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { loadConfig } from "../src/core/config.ts";
import type { Config } from "../src/core/config.ts";
import {
  beginLogin,
  readLoginState,
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
  JSON.parse(readFileSync(join(config.stateDir, "login.json"), "utf8")) as {
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

/**
 * A fake launcher that records what it was asked to run and reports the
 * caller's own pid, so the liveness check has a living process to find.
 */
function fakeSpawn() {
  const launched: string[][] = [];
  return {
    launched,
    spawnChild: (command: string, args: string[]) => {
      launched.push([command, ...args]);
      return { pid: process.pid, unref: () => {} };
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
