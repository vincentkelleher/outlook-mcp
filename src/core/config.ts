import { existsSync, readFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";

export type Env = Record<string, string | undefined>;

const logLevels = ["debug", "info", "warn", "error"] as const;
export type LogLevel = (typeof logLevels)[number];

export type Config = {
  clientId: string | undefined;
  authority: string;
  cachePath: string;
  stateDir: string;
  attachmentDir: string;
  readonly: boolean;
  logLevel: LogLevel;
};

/**
 * Bun loaded `.env` automatically and Node does not, so every reader starts
 * here. Variables already set in the environment win over the file.
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
// cannot do on its own.
function homeDir(env: Env) {
  return env.HOME || homedir();
}

function readLogLevel(env: Env): LogLevel {
  const value = env.OUTLOOK_MCP_LOG_LEVEL?.toLowerCase();
  return logLevels.find((level) => level === value) ?? "info";
}

export function loadConfig(env: Env = loadEnvFile()): Config {
  const home = env.OUTLOOK_SKILL_HOME || join(homeDir(env), ".outlook-skill");
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
    logLevel: readLogLevel(env),
  };
}
