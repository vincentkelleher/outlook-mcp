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
  message?: string;
};

export type LoginLaunch = LoginState & {
  instructions: string;
  account?: string;
};

type Child = { unref(): void };
type SpawnChild = (
  command: string,
  args: string[],
  options: { stdio: "ignore"; detached: true },
) => Child;

// The state file is the only channel between the detached sign-in process and
// whoever started it, so it carries a URL and a code, never a token.
const stateTimeoutMs = 3000;
const statePollMs = 100;

function statePath(config: Config) {
  return join(config.stateDir, "login.json");
}

export function writeLoginState(config: Config, state: LoginState): void {
  const path = statePath(config);
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  writeFileSync(path, JSON.stringify(state));
}

/** An absent or half-written file simply means "not ready yet". */
export function readLoginState(config: Config): LoginState | undefined {
  try {
    return JSON.parse(readFileSync(statePath(config), "utf8")) as LoginState;
  } catch {
    return undefined;
  }
}

export function clearLoginState(config: Config): void {
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
  clearLoginState(config);
  const launch = deps.spawnChild ?? spawn;
  launch(process.execPath, [entry, "login"], {
    stdio: "ignore",
    detached: true,
  }).unref();
}

async function waitForLoginState(config: Config): Promise<LoginState> {
  const deadline = Date.now() + stateTimeoutMs;
  for (;;) {
    const state = readLoginState(config);
    if (state) return state;
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
  return "Sign-in is starting. Call auth_status to poll, and call login again if no code appears before the flow expires.";
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
  clearLoginState(config);
  try {
    await createAuth({ config }).accessToken({
      interactive: true,
      onDeviceCode: (info) => {
        writeLoginState(config, {
          state: "pending",
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
