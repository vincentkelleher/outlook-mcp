import { existsSync, readFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";

export type Env = Record<string, string | undefined>;

export type Config = {
  clientId: string | undefined;
  authority: string;
  cachePath: string;
  stateDir: string;
  attachmentDir: string;
  readonly: boolean;
};

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
  return join(config.stateDir, "login.json");
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
  };
}
