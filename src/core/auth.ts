import {
  chmodSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { dirname, join } from "node:path";
import { randomUUID } from "node:crypto";
import { PublicClientApplication } from "@azure/msal-node";
import type { ICachePlugin } from "@azure/msal-node";
import { loadConfig, loginStatePaths, scopesFor } from "./config.ts";
import type { Config } from "./config.ts";
import {
  AuthRequiredError,
  sessionProblem,
  noSessionMessage,
  sessionProblemMessage,
} from "./errors.ts";

export type MsalAccount = { username?: string };
export type MsalResult = { accessToken?: string };

/** What MSAL hands the device-code callback; the code and URL are user-facing. */
export type DeviceCodeInfo = {
  message: string;
  userCode?: string;
  verificationUri?: string;
  expiresIn?: number;
};

/** The slice of MSAL this project uses, so tests can pass a fake application. */
export type MsalApplication = {
  getTokenCache(): { getAllAccounts(): Promise<MsalAccount[]> };
  acquireTokenSilent(request: {
    account: MsalAccount;
    scopes: string[];
  }): Promise<MsalResult>;
  acquireTokenByDeviceCode(options: {
    scopes: string[];
    deviceCodeCallback: (response: DeviceCodeInfo) => void;
  }): Promise<MsalResult | null>;
};

export type MsalFactory = (config: Config) => MsalApplication;

/** Why an account that exists cannot be used as it stands. */
export type AuthProblem = "needsConsent" | "expired" | "unavailable";

export type AuthStatus = {
  signedIn: boolean;
  account: string | undefined;
  /** Only set when an account is present but unusable. */
  reason?: AuthProblem;
  authority: string;
  cachePath: string;
};

// The name a cache write goes through on its way to `msal-cache.json`, and the
// two ends a sweep recognises it by. Neither the real cache nor anything else in
// the directory matches both.
const tempCachePrefix = ".msal-cache.";
const tempCacheSuffix = ".tmp";

/**
 * A process killed between writing the temp cache and renaming it leaves a
 * plaintext copy of the token cache beside the real one. Every read sweeps such
 * copies once they pass this age, because nothing can still be writing to one -
 * anything newer belongs to another client. Best effort: a sweep that cannot read
 * the directory, or cannot delete a file, is not worth failing a token refresh
 * over.
 */
const staleTempCacheAgeMs = 60 * 60 * 1000;

function tempCachePath(dir: string): string {
  return join(
    dir,
    `${tempCachePrefix}${process.pid}.${randomUUID()}${tempCacheSuffix}`,
  );
}

function sweepStaleTempCaches(dir: string): void {
  let names: string[];
  try {
    names = readdirSync(dir);
  } catch {
    return;
  }
  for (const name of names) {
    if (!name.startsWith(tempCachePrefix) || !name.endsWith(tempCacheSuffix))
      continue;
    const path = join(dir, name);
    try {
      if (Date.now() - statSync(path).mtimeMs < staleTempCacheAgeMs) continue;
      rmSync(path, { force: true });
    } catch {
      // Another client claimed it in the meantime, or the directory moved.
    }
  }
}

export function createCachePlugin(cachePath: string): ICachePlugin {
  return {
    beforeCacheAccess: async (context) => {
      // Before reading, so a copy left by a killed write is the first thing the
      // next session cleans up rather than something it never notices.
      sweepStaleTempCaches(dirname(cachePath));
      if (existsSync(cachePath))
        context.tokenCache.deserialize(readFileSync(cachePath, "utf8"));
    },
    afterCacheAccess: async (context) => {
      if (!context.cacheHasChanged) return;
      const dir = dirname(cachePath);
      mkdirSync(dir, { recursive: true, mode: 0o700 });
      chmodSync(dir, 0o700);
      // Several MCP clients share this file: write a temp file, then rename, so
      // a reader never sees a half-written cache.
      const temp = tempCachePath(dir);
      try {
        writeFileSync(temp, context.tokenCache.serialize(), { mode: 0o600 });
        chmodSync(temp, 0o600);
        renameSync(temp, cachePath);
      } finally {
        rmSync(temp, { force: true });
      }
    },
  };
}

export function createMsal(config: Config): MsalApplication {
  if (!config.clientId)
    throw new Error(
      "Set OUTLOOK_OAUTH_CLIENT_ID in the environment or .env file.",
    );
  return new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: config.authority,
    },
    cache: { cachePlugin: createCachePlugin(config.cachePath) },
  }) as unknown as MsalApplication;
}

export function createAuth(deps: { config?: Config; msal?: MsalFactory } = {}) {
  const config = deps.config ?? loadConfig();
  const scopes = scopesFor(config.readonly);
  const newMsal = deps.msal ?? createMsal;
  let application: MsalApplication | undefined;
  const msal = () => (application ??= newMsal(config));

  function token(result: MsalResult | null) {
    if (!result?.accessToken)
      throw new Error("Microsoft did not return an access token.");
    return result.accessToken;
  }

  async function firstAccount() {
    return (await msal().getTokenCache().getAllAccounts())[0];
  }

  async function silent(account: MsalAccount) {
    return token(await msal().acquireTokenSilent({ account, scopes }));
  }

  async function accessToken(
    options: {
      interactive?: boolean;
      onDeviceCode?: (info: DeviceCodeInfo) => void;
    } = {},
  ): Promise<string> {
    const account = await firstAccount();
    if (account) {
      try {
        return await silent(account);
      } catch (error) {
        if (options.interactive) return deviceToken(options.onDeviceCode);
        // Another client may have refreshed the shared cache since we read it,
        // so re-read and retry once before asking anyone to sign in again.
        const refreshed = await firstAccount();
        if (refreshed) {
          try {
            return await silent(refreshed);
          } catch {
            /* report the first failure */
          }
        }
        const reason = error instanceof Error ? error.message : String(error);
        throw new AuthRequiredError(sessionProblemMessage(reason, scopes), {
          cause: error,
        });
      }
    }
    if (!options.interactive) throw new AuthRequiredError(noSessionMessage);
    return deviceToken(options.onDeviceCode);
  }

  async function deviceToken(onDeviceCode?: (info: DeviceCodeInfo) => void) {
    const result = await msal().acquireTokenByDeviceCode({
      scopes,
      // A caller that takes the callback owns the prompt, so the flow is
      // reported once, not twice.
      deviceCodeCallback:
        onDeviceCode ?? ((response) => console.error(response.message)),
    });
    return token(result);
  }

  /**
   * An account in the cache is not yet a session: the grant behind it can be
   * narrower than the scopes this process asks for - a mailbox signed in before
   * drafts existed is signed in for `Mail.Read` - and then every call that needs
   * one fails while status claims the user is signed in, so the agent is sent
   * round in circles. Status therefore asks the same question a call would, and
   * reports signed in only when a token really comes back.
   */
  async function authStatus(): Promise<AuthStatus> {
    const base = {
      authority: config.authority,
      cachePath: config.cachePath,
    };
    const account = await firstAccount();
    if (!account)
      return {
        signedIn: false,
        account: undefined,
        ...base,
      };
    try {
      await silent(account);
      return { signedIn: true, account: account.username, ...base };
    } catch (error) {
      // Which of the three this is has one decider, in errors.ts, so the tool
      // surface and the thrown message cannot disagree about whether the user
      // owes a consent or the network simply failed.
      const reason = error instanceof Error ? error.message : String(error);
      return {
        signedIn: false,
        account: account.username,
        reason: sessionProblem(reason),
        ...base,
      };
    }
  }

  return { accessToken, authStatus };
}

/**
 * Sign out. The sign-in state goes with the cache: an error left behind by the
 * last attempt would otherwise read as a fresh failure to whoever signs in next
 * with a different account, or to the same one after a fix.
 */
export function logout(config: Config = loadConfig()): void {
  rmSync(config.cachePath, { force: true });
  for (const state of loginStatePaths(config)) rmSync(state, { force: true });
}
