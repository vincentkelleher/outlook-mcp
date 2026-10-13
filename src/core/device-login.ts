import { spawn } from "node:child_process";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { createAuth } from "./auth.ts";
import type { Config } from "./config.ts";

export type LoginState = {
  state: "pending" | "authorized" | "error";
  verificationUri?: string;
  userCode?: string;
  expiresAt?: string;
  /** The sign-in process that owns a pending state, so no second one starts. */
  pid?: number;
  message?: string;
};

export type LoginLaunch = LoginState & {
  instructions: string;
  account?: string;
};

type Child = { unref(): void; pid?: number };
type SpawnChild = (
  command: string,
  args: string[],
  options: { stdio: "ignore"; detached: true },
) => Child;

// The state file is the only channel between the detached sign-in process and
// whoever started it, so it carries a URL and a code, never a token.
const stateTimeoutMs = 3000;
const statePollMs = 100;
/**
 * How long a sign-in process may hold the state before it has published a code.
 * Deliberately far shorter than the 15 minutes a device-code flow can run: a
 * child that dies before Microsoft answers must not block sign-in for that long.
 */
const launchLeaseMs = 30_000;

function statePath(config: Config) {
  return join(config.stateDir, "login.json");
}

export function writeLoginState(config: Config, state: LoginState): void {
  const path = statePath(config);
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  writeFileSync(path, JSON.stringify(state));
}

/**
 * A pending state carries the flow's own deadline. Past it nobody is polling
 * Microsoft any more - the child may have died after publishing the code - and
 * leaving the file in place makes every client on the machine tell the user to
 * enter a dead code for ever. Read as absent, the next login starts a fresh flow.
 */
function expired(state: LoginState): boolean {
  if (state.state !== "pending" || state.expiresAt === undefined) return false;
  const deadline = Date.parse(state.expiresAt);
  return !Number.isNaN(deadline) && deadline <= Date.now();
}

/** An absent, half-written or expired file simply means "not ready yet". */
export function readLoginState(config: Config): LoginState | undefined {
  let state: LoginState;
  try {
    state = JSON.parse(readFileSync(statePath(config), "utf8")) as LoginState;
  } catch {
    return undefined;
  }
  return expired(state) ? undefined : state;
}

/** Signal 0 asks whether a process exists without disturbing it. */
function processAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    // EPERM means it exists but belongs to someone else; anything else is gone.
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}

/**
 * A flow somebody else already started: still inside its own deadline and still
 * owned by a living process. `pid` absent means no owner to check, so nothing
 * is being guarded.
 */
function inFlight(state: LoginState | undefined): boolean {
  if (state?.state !== "pending") return false;
  if (state.pid === undefined || state.expiresAt === undefined) return false;
  return Date.parse(state.expiresAt) > Date.now() && processAlive(state.pid);
}

/** The claim a sign-in process holds on the state before Microsoft has answered. */
function launchLease(pid: number): LoginState {
  return {
    state: "pending",
    pid,
    expiresAt: new Date(Date.now() + launchLeaseMs).toISOString(),
  };
}

function clearLoginState(config: Config): void {
  rmSync(statePath(config), { force: true });
}

/**
 * Start sign-in in a process that survives the caller. The device-code flow
 * polls Microsoft for up to 15 minutes, far past any MCP request timeout, so
 * the child is detached, unref'd and left without stdio.
 */
export function startLogin(
  config: Config,
  deps: { spawnChild?: SpawnChild; entry?: string } = {},
): void {
  const entry = deps.entry ?? process.argv[1];
  if (!entry)
    throw new Error("Cannot start sign-in: there is no entry module to run.");
  // One flow at a time. Two clients that both spawned would each ask Microsoft
  // for a code, show the user two of them, and the loser would later overwrite
  // the winner's outcome with an error - while Microsoft rate-limits the retry.
  if (inFlight(readLoginState(config))) return;
  clearLoginState(config);
  const launch = deps.spawnChild ?? spawn;
  const child = launch(process.execPath, [entry, "login"], {
    stdio: "ignore",
    detached: true,
  });
  // The lease names the child, not us: this process may be gone by the time a
  // second client asks, and the flow outlives it.
  writeLoginState(config, launchLease(child.pid ?? process.pid));
  child.unref();
}

/** A code the user can still enter is an answer; a launch lease is not one yet. */
function publishable(state: LoginState): boolean {
  return state.state !== "pending" || state.userCode !== undefined;
}

async function waitForLoginState(config: Config): Promise<LoginState> {
  const deadline = Date.now() + stateTimeoutMs;
  for (;;) {
    const state = readLoginState(config);
    if (state && publishable(state)) return state;
    if (Date.now() >= deadline) break;
    await new Promise((resolve) => setTimeout(resolve, statePollMs));
  }
  return {
    state: "pending",
    message:
      "Sign-in has started, but Microsoft has not returned a code within 3 seconds.",
  };
}

function instructionsFor(state: LoginState): string {
  if (state.state === "authorized")
    return "Already signed in: nothing more to do, retry the call that failed.";
  if (state.state === "error")
    return `Sign-in could not start (${state.message ?? "unknown reason"}). Fix the cause, then call login again.`;
  if (state.userCode && state.verificationUri)
    return `Show the user ${state.verificationUri} and the code ${state.userCode}, ask them to approve it, then call auth_status until signedIn is true and retry the call that failed.`;
  return "Sign-in is starting. Call auth_status to poll: it carries the verification URL and user code as soon as Microsoft issues them, and reports signedIn once the user approves. Call login again if no code appears before the flow expires.";
}

/** The answer when a session already exists, so no flow is started. */
export function alreadySignedIn(account?: string): LoginLaunch {
  return {
    state: "authorized",
    account,
    instructions: instructionsFor({ state: "authorized" }),
  };
}

/**
 * Start the detached flow and return as soon as the child has published a
 * code. Never waits for the user.
 */
export async function beginLogin(
  config: Config,
  start: (config: Config) => void = startLogin,
): Promise<LoginLaunch> {
  start(config);
  const state = await waitForLoginState(config);
  return { ...state, instructions: instructionsFor(state) };
}

/**
 * The foreground half of sign-in, run by the detached child: records the code
 * the moment Microsoft issues it, then the outcome.
 */
export async function runLogin(config: Config): Promise<void> {
  // Take the lease for this process rather than just deleting the old state: it
  // replaces whatever the last attempt left, and it is what stops a second
  // client spawning a competing flow while Microsoft has not answered yet.
  writeLoginState(config, launchLease(process.pid));
  try {
    await createAuth({ config }).accessToken({
      interactive: true,
      onDeviceCode: (info) => {
        writeLoginState(config, {
          state: "pending",
          pid: process.pid,
          verificationUri: info.verificationUri,
          userCode: info.userCode,
          expiresAt:
            info.expiresIn === undefined
              ? undefined
              : new Date(Date.now() + info.expiresIn * 1000).toISOString(),
        });
        console.error(info.message);
      },
    });
    writeLoginState(config, { state: "authorized" });
  } catch (error) {
    writeLoginState(config, {
      state: "error",
      message: error instanceof Error ? error.message : String(error),
    });
    throw error;
  }
}
