import { existsSync, readdirSync, readFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { logFormatFor, logLevelFor } from "./log.ts";
import type { LogFormat, LogLevel } from "./log.ts";

export type Env = Record<string, string | undefined>;

export type Config = {
  clientId: string | undefined;
  authority: string;
  cachePath: string;
  stateDir: string;
  attachmentDir: string;
  readonly: boolean;
  /** Interface `serve --http` binds: a hostname or an address. */
  httpHost: string;
  httpPort: number;
  /** Host header names accepted beyond the loopback ones. */
  allowedHosts: string[];
  /** Browser origins accepted beyond the same origin. */
  allowedOrigins: string[];
  /** How long an HTTP session with nothing connected lives, in minutes. */
  sessionIdleMinutes: number;
  /** stderr verbosity for every surface: CLI, stdio server and HTTP server. */
  logLevel: LogLevel;
  /** `text` for a person reading a terminal, `json` for a log collector. */
  logFormat: LogFormat;
};

/**
 * Defaults for `serve --http`. Nothing on that port authenticates, so the host
 * defaults to loopback; the port is this server's own, chosen to be unlikely to
 * clash.
 */
export const defaultHttpHost = "127.0.0.1";
export const defaultHttpPort = 3939;

/** The names a request may carry, from a comma-separated variable. */
function list(value: string | undefined): string[] {
  return (value ?? "")
    .split(",")
    .map((entry) => entry.trim().replace(/\/+$/, "").toLowerCase())
    .filter((entry) => entry !== "");
}

/**
 * A number from the environment is read once, here, and fails loudly:
 * `server.listen(NaN)` answers something about a valid range, which tells
 * nobody that `OUTLOOK_MCP_PORT=393a` is the cause.
 */
function numberFor(
  name: string,
  value: string | undefined,
  fallback: number,
  maximum: number,
): number {
  const raw = (value ?? "").trim();
  if (raw === "") return fallback;
  const parsed = Number(raw);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum)
    throw new Error(
      `${name} must be an integer from 1 to ${maximum}, not "${value}".`,
    );
  return parsed;
}

/** One week of minutes: past that a session is a leak, not a setting. */
const maxIdleMinutes = 7 * 24 * 60;
/** A day without reconnecting is long enough for any client on a laptop. */
export const defaultSessionIdleMinutes = 24 * 60;

/** The package's own version: one read, shared by all that name this release. */
export function packageVersion(): string {
  try {
    const manifest = JSON.parse(
      readFileSync(new URL("../../package.json", import.meta.url), "utf8"),
    ) as { version?: string };
    return manifest.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

/**
 * Node does not load `.env`, so every reader starts here: the file is where a
 * client id belongs and the environment must still win over it.
 */
export function loadEnvFile(dir = process.cwd(), env: Env = process.env): Env {
  const merged: Env = { ...env };
  const path = join(dir, ".env");
  if (!existsSync(path)) return merged;
  for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
    const entry =
      /^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/.exec(line);
    if (!entry) continue;
    const [, key, raw] = entry;
    if (merged[key] !== undefined) continue;
    merged[key] = raw.replace(/^(['"])(.*)\1$/s, "$2");
  }
  return merged;
}

// An injected env must be able to move the home directory, which os.homedir()
// cannot do on its own. Client configuration files hang off the same root.
export function homeDir(env: Env) {
  return env.HOME || homedir();
}

/**
 * The sign-in state file, named here because two modules have to agree on it:
 * `device-login.ts` writes it and `logout` removes it, so that signing out also
 * clears the last sign-in failure. It carries a URL, a code and an error, never
 * a token.
 */
export function loginStatePath(config: Config): string {
  // A flow is only interchangeable with a flow that asked Microsoft for the same
  // permissions: a `Mail.Read` code the user approves cannot widen a session to
  // `Mail.ReadWrite`. So the mail scope is part of the file name, and a
  // read-only server and a read-write server sharing one machine keep separate
  // flows and can neither answer for, nor block, each other.
  const mailScope =
    scopesFor(config.readonly).find((scope) => scope.startsWith("Mail.")) ??
    "no-mail-scope";
  return join(config.stateDir, `login-${mailScope}.json`);
}

/**
 * Every sign-in state file present, whatever scope set it belongs to. Signing
 * out clears them all, including any left behind by an earlier name.
 */
export function loginStatePaths(config: Config): string[] {
  try {
    return readdirSync(config.stateDir)
      .filter((name) => name.startsWith("login") && name.endsWith(".json"))
      .map((name) => join(config.stateDir, name));
  } catch {
    return [];
  }
}

/**
 * Delegated Graph permissions. Creating or changing a message - which is all a
 * draft is - needs `Mail.ReadWrite`, so a readonly server asks for `Mail.Read`
 * instead and gets no write tool. The consent prompt shows these names, so they
 * must match what the code is actually able to do.
 */
export function scopesFor(readonly: boolean): string[] {
  return [
    "User.Read",
    readonly ? "Mail.Read" : "Mail.ReadWrite",
    "email",
    "profile",
  ];
}

export function loadConfig(env: Env = loadEnvFile()): Config {
  const home = env.OUTLOOK_MCP_HOME || join(homeDir(env), ".outlook-mcp");
  const tenant = env.OUTLOOK_OAUTH_TENANT || "consumers";
  return {
    clientId: env.OUTLOOK_OAUTH_CLIENT_ID || undefined,
    authority: `https://login.microsoftonline.com/${tenant.trim()}/`,
    cachePath: join(home, "msal-cache.json"),
    stateDir: env.OUTLOOK_MCP_STATE_DIR || join(home, "state"),
    attachmentDir:
      env.OUTLOOK_MCP_ATTACHMENT_DIR || join(tmpdir(), "outlook-mcp"),
    readonly: !["", "0", "false"].includes(
      (env.OUTLOOK_MCP_READONLY ?? "").toLowerCase(),
    ),
    httpHost: (env.OUTLOOK_MCP_HOST ?? "").trim() || defaultHttpHost,
    httpPort: numberFor(
      "OUTLOOK_MCP_PORT",
      env.OUTLOOK_MCP_PORT,
      defaultHttpPort,
      65535,
    ),
    allowedHosts: list(env.OUTLOOK_MCP_ALLOWED_HOSTS),
    allowedOrigins: list(env.OUTLOOK_MCP_ALLOWED_ORIGINS),
    sessionIdleMinutes: numberFor(
      "OUTLOOK_MCP_SESSION_IDLE_MINUTES",
      env.OUTLOOK_MCP_SESSION_IDLE_MINUTES,
      defaultSessionIdleMinutes,
      maxIdleMinutes,
    ),
    // Not `OUTLOOK_MCP_*`: the CLI is diagnosed the same way as the server, and
    // a person running `emails list` to find out what is wrong needs the level
    // that covers both commands.
    logLevel: logLevelFor("OUTLOOK_LOG_LEVEL", env.OUTLOOK_LOG_LEVEL),
    logFormat: logFormatFor("OUTLOOK_LOG_FORMAT", env.OUTLOOK_LOG_FORMAT),
  };
}
