import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join } from "node:path";
import { randomUUID } from "node:crypto";
import { PublicClientApplication } from "@azure/msal-node";
import type { ICachePlugin } from "@azure/msal-node";
import { loadConfig, scopesFor } from "./config.ts";
import type { Config } from "./config.ts";
import {
  AuthRequiredError,
  isSessionExpiry,
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
export type AuthProblem = "needsConsent" | "expired";

export type AuthStatus = {
  signedIn: boolean;
  account: string | undefined;
  /** Only set when an account is present but unusable. */
  reason?: AuthProblem;
  authority: string;
  cachePath: string;
};

export function createCachePlugin(cachePath: string): ICachePlugin {
  return {
    beforeCacheAccess: async (context) => {
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
      const temp = join(dir, `.msal-cache.${process.pid}.${randomUUID()}.tmp`);
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
      // A silent failure that is not expiry means the grant has to be widened,
      // which only a fresh consent flow can do.
      const reason = error instanceof Error ? error.message : String(error);
      return {
        signedIn: false,
        account: account.username,
        reason: isSessionExpiry(reason) ? "expired" : "needsConsent",
        ...base,
      };
    }
  }

  return { accessToken, authStatus, config };
}

export function logout(cachePath = loadConfig().cachePath) {
  rmSync(cachePath, { force: true });
}
