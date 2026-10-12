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
import { loadConfig, loadEnvFile } from "./config.ts";
import type { Config, Env } from "./config.ts";
import {
  AuthRequiredError,
  expiredSessionMessage,
  noSessionMessage,
} from "./errors.ts";

const scopes = ["User.Read", "Mail.Read", "email", "profile"];

export type MsalAccount = { username?: string };
export type MsalResult = { accessToken?: string };

/** The slice of MSAL this project uses, so tests can pass a fake application. */
export type MsalApplication = {
  getTokenCache(): { getAllAccounts(): Promise<MsalAccount[]> };
  acquireTokenSilent(request: {
    account: MsalAccount;
    scopes: string[];
  }): Promise<MsalResult>;
  acquireTokenByDeviceCode(options: {
    scopes: string[];
    deviceCodeCallback: (response: { message: string }) => void;
  }): Promise<MsalResult | null>;
};

export type MsalFactory = (config: Config) => MsalApplication;

export type AuthStatus = {
  signedIn: boolean;
  account: string | undefined;
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
    options: { interactive?: boolean } = {},
  ): Promise<string> {
    const account = await firstAccount();
    if (account) {
      try {
        return await silent(account);
      } catch (error) {
        if (options.interactive) return deviceToken();
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
        throw new AuthRequiredError(expiredSessionMessage(reason), {
          cause: error,
        });
      }
    }
    if (!options.interactive) throw new AuthRequiredError(noSessionMessage);
    return deviceToken();
  }

  async function deviceToken() {
    const result = await msal().acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => console.error(response.message),
    });
    return token(result);
  }

  async function authStatus(): Promise<AuthStatus> {
    const account = await firstAccount();
    return {
      signedIn: Boolean(account),
      account: account?.username,
      authority: config.authority,
      cachePath: config.cachePath,
    };
  }

  return { accessToken, authStatus, config };
}

export function accessToken(
  env: Env = loadEnvFile(),
  options: { interactive?: boolean } = {},
): Promise<string> {
  return createAuth({ config: loadConfig(env) }).accessToken(options);
}

export function authStatus(env: Env = loadEnvFile()): Promise<AuthStatus> {
  return createAuth({ config: loadConfig(env) }).authStatus();
}

export function logout(cachePath = loadConfig().cachePath) {
  rmSync(cachePath, { force: true });
}
