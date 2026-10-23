import { spawn } from "node:child_process";
import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { basename, dirname } from "node:path";
import { createAuth } from "./auth.ts";
import { loginStatePath } from "./config.ts";
import type { Config } from "./config.ts";
import { errorFields, logger, verificationCode } from "./log.ts";

const log = logger("login");

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

/**
 * What this module needs of the sign-in process: it must be detachable, and it
 * must be able to report that it never started, because a spawn that fails emits
 * `error` rather than throwing.
 */
type Child = {
  unref(): void;
  pid?: number;
  on(event: "error", listener: (error: unknown) => unknown): unknown;
};
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

function stateDirectory(config: Config): string {
  const dir = dirname(loginStatePath(config));
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  // mkdir's mode is filtered by the umask and ignored when the directory is
  // already there, so state it again: a user code is redeemable by whoever reads
  // it, and a shared state directory must not make it world-readable.
  chmodSync(dir, 0o700);
  return dir;
}

export function writeLoginState(config: Config, state: LoginState): void {
  stateDirectory(config);
  const path = loginStatePath(config);
  writeFileSync(path, JSON.stringify(state), { mode: 0o600 });
  // The state names a flow, so this says which one moved to which state, with
  // the code masked until somebody asks for `debug` — a code is only useful to
  // the person being told it, and a log file is not that person.
  log.debug("sign-in state written", {
    file: basename(path),
    state: state.state,
    owner: state.pid,
    userCode: verificationCode(state.userCode, log.detailed),
  });
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
  const path = loginStatePath(config);
  let state: LoginState;
  try {
    state = JSON.parse(readFileSync(path, "utf8")) as LoginState;
  } catch (error) {
    // An absent file is the ordinary case — nobody has signed in yet. A file
    // that is there and will not parse is not: it is half a write, or something
    // else living at this name, and every client on the machine then behaves as
    // though no flow exists.
    if (existsSync(path))
      log.warn("sign-in state is unreadable and is treated as absent", {
        file: basename(path),
        ...errorFields(error, log.detailed),
      });
    return undefined;
  }
  if (expired(state)) {
    // Explains the one question this file raises: why a code the user was shown
    // never becomes a session.
    log.debug("sign-in state is past its deadline", {
      file: basename(path),
      expiresAt: state.expiresAt,
    });
    return undefined;
  }
  return state;
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
  rmSync(loginStatePath(config), { force: true });
}

/**
 * Take the state file for a new flow, atomically. Checking `inFlight` and then
 * writing is two operations, so two processes can both decide they are first;
 * creating the file with `wx` makes that decision one, and only one of them
 * creates it. Whatever the check leaves behind that is not a live flow - a
 * finished outcome, or a flow whose process died - is replaced.
 */
function claimLaunch(config: Config, lease: LoginState): boolean {
  for (let attempt = 0; attempt < 2; attempt += 1) {
    stateDirectory(config);
    try {
      writeFileSync(loginStatePath(config), JSON.stringify(lease), {
        mode: 0o600,
        flag: "wx",
      });
      return true;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      if (inFlight(readLoginState(config))) return false;
      clearLoginState(config);
    }
  }
  return false;
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
  // The claim is made before spawning, under our own pid, so no second client
  // can slip in while the spawn is in flight.
  if (!claimLaunch(config, launchLease(process.pid))) {
    const held = readLoginState(config);
    log.info("another sign-in owns the state, so no second flow was started", {
      owner: held?.pid,
      userCode: verificationCode(held?.userCode, log.detailed),
    });
    return;
  }
  const launch = deps.spawnChild ?? spawn;
  log.debug("starting detached sign-in process", { entry });
  const child = launch(process.execPath, [entry, "login"], {
    stdio: "ignore",
    detached: true,
  });
  // The lease then names the child, not us: this process may be gone by the time
  // a second client asks, and the flow outlives it.
  writeLoginState(config, launchLease(child.pid ?? process.pid));
  // A child that cannot be started writes nothing at all, so without this the
  // caller waits out the whole poll and reports a slow sign-in instead of the
  // reason. An `error` event also has to be handled or it throws out of the
  // server process.
  child.on("error", (error) => {
    log.error("the detached sign-in process could not be started", {
      ...errorFields(error, log.detailed),
    });
    writeLoginState(config, {
      state: "error",
      message: error instanceof Error ? error.message : String(error),
    });
  });
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
  log.warn("sign-in started but Microsoft issued no code within 3 seconds", {
    timeoutMs: stateTimeoutMs,
  });
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
  // The detached child owns the state it was spawned for and reclaims it under
  // its own pid. But a login started in a terminal while another process holds a
  // live flow must not take that state: it would replace a code the user may
  // already be typing with a second one they were never shown.
  const existing = readLoginState(config);
  if (
    existing?.state === "pending" &&
    inFlight(existing) &&
    existing.pid !== process.pid
  ) {
    // Printed, not only logged: a person who ran `login` in a terminal is owed
    // the sentence, whatever the log level says.
    const notice = `A sign-in is already running as process ${existing.pid}, so this one did not start and no second code was issued. Ask the agent for auth_status to see the code that is live, or sign out to clear it.`;
    console.error(notice);
    log.warn("this sign-in stood down for one already running", {
      owner: existing.pid,
    });
    return;
  }
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
        log.info("Microsoft issued a device code", {
          verificationUri: info.verificationUri,
          userCode: verificationCode(info.userCode, log.detailed),
          expiresIn: info.expiresIn,
        });
      },
    });
    writeLoginState(config, { state: "authorized" });
    log.info("sign-in completed");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    // The state file outlives this process, and `auth_status` reports it, so a
    // failure here is visible twice by design: once in the log that says when it
    // happened, once in the status a polling agent reads afterwards.
    log.error("sign-in failed", { ...errorFields(error, log.detailed) });
    writeLoginState(config, { state: "error", message });
    throw error;
  }
}
