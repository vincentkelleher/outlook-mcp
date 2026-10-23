import assert from "node:assert/strict";
import { test } from "node:test";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  statSync,
  utimesSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import {
  createAuth,
  createCachePlugin,
  createMsal,
  logout,
} from "../src/core/auth.ts";
import type {
  MsalAccount,
  MsalApplication,
  MsalFactory,
  MsalResult,
} from "../src/core/auth.ts";
import { loadConfig } from "../src/core/config.ts";
import type { Env } from "../src/core/config.ts";
import { readLoginState, writeLoginState } from "../src/core/device-login.ts";
import { AuthRequiredError, sessionProblem } from "../src/core/errors.ts";

const clientId = "00000000-0000-0000-0000-000000000000";
const consumers = "https://login.microsoftonline.com/consumers/";

const tempDir = () => mkdtempSync(join(tmpdir(), "outlook-mcp-"));
const tempCache = () => join(tempDir(), "msal-cache.json");
const testConfig = (env: Env = {}) =>
  loadConfig({ HOME: tempDir(), OUTLOOK_OAUTH_CLIENT_ID: clientId, ...env });

type Calls = {
  accountReads: number;
  silent: MsalAccount[];
  scopes: string[][];
  device: number;
};

function fakeMsal(behaviour: {
  accounts: MsalAccount[][];
  /** Given the scopes this process asks for, so a narrow grant can be replayed. */
  silent?: (account: MsalAccount, scopes: string[]) => Promise<MsalResult>;
  device?: () => Promise<MsalResult | null>;
}) {
  const calls: Calls = {
    accountReads: 0,
    silent: [],
    scopes: [],
    device: 0,
  };
  const application: MsalApplication = {
    getTokenCache() {
      return {
        getAllAccounts: async () => {
          const snapshot =
            behaviour.accounts[
              Math.min(calls.accountReads, behaviour.accounts.length - 1)
            ] ?? [];
          calls.accountReads += 1;
          return snapshot;
        },
      };
    },
    async acquireTokenSilent({ account, scopes }): Promise<MsalResult> {
      calls.silent.push(account);
      calls.scopes.push(scopes);
      if (behaviour.silent) return behaviour.silent(account, scopes);
      return { accessToken: "silent-token" };
    },
    async acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback,
    }): Promise<MsalResult | null> {
      calls.device += 1;
      calls.scopes.push(scopes);
      deviceCodeCallback({ message: "To sign in, open the verification URL" });
      if (behaviour.device) return behaviour.device();
      return { accessToken: "device-token" };
    },
  };
  const factory: MsalFactory = () => application;
  return { calls, factory };
}

async function captureStderr(run: () => Promise<void>) {
  const lines: string[] = [];
  const original = console.error;
  console.error = (...args: unknown[]) => lines.push(args.join(" "));
  try {
    await run();
  } finally {
    console.error = original;
  }
  return lines;
}

test("cache plugin restores and persists changed MSAL caches", async () => {
  const cache = tempCache();
  const plugin = createCachePlugin(cache);
  const restored: string[] = [];
  mkdirSync(dirname(cache), { recursive: true });
  writeFileSync(cache, "saved");

  await plugin.beforeCacheAccess({
    tokenCache: { deserialize: (value: string) => restored.push(value) },
  } as never);
  await plugin.afterCacheAccess({
    cacheHasChanged: true,
    tokenCache: { serialize: () => "updated" },
  } as never);

  assert.deepEqual(restored, ["saved"]);
  assert.equal(readFileSync(cache, "utf8"), "updated");
  assert.equal(statSync(dirname(cache)).mode & 0o777, 0o700);
  assert.equal(statSync(cache).mode & 0o777, 0o600);
});

test("cache plugin does not write unchanged caches", async () => {
  const cache = tempCache();
  const plugin = createCachePlugin(cache);
  let touched = 0;

  await plugin.beforeCacheAccess({
    tokenCache: {
      deserialize: () => {
        touched += 1;
      },
    },
  } as never);
  await plugin.afterCacheAccess({
    cacheHasChanged: false,
    tokenCache: {
      serialize: () => {
        touched += 1;
        return "";
      },
    },
  } as never);

  assert.equal(touched, 0);
  assert.equal(existsSync(cache), false);
});

test("cache writes replace the file by rename and leave no temp files", async () => {
  const cache = tempCache();
  mkdirSync(dirname(cache), { recursive: true });
  writeFileSync(cache, "before", { mode: 0o600 });
  const before = statSync(cache).ino;

  await createCachePlugin(cache).afterCacheAccess({
    cacheHasChanged: true,
    tokenCache: { serialize: () => "after" },
  } as never);

  assert.equal(readFileSync(cache, "utf8"), "after");
  // A rename swaps the inode. An in-place rewrite would keep it, and would show
  // a truncated cache to a second client reading at the same moment.
  assert.notEqual(statSync(cache).ino, before);
  assert.deepEqual(readdirSync(dirname(cache)), ["msal-cache.json"]);
  assert.equal(statSync(cache).mode & 0o777, 0o600);
});

test("silent refresh returns a token without prompting", async () => {
  const account = { username: "someone@example.com" };
  const { calls, factory } = fakeMsal({ accounts: [[account]] });

  const token = await createAuth({
    config: testConfig(),
    msal: factory,
  }).accessToken();

  assert.equal(token, "silent-token");
  assert.deepEqual(calls.silent, [account]);
  assert.equal(calls.device, 0);
  assert.equal(calls.accountReads, 1);
});

test("a silent failure re-reads the shared cache and retries once", async () => {
  const stale = { username: "stale@example.com" };
  const refreshed = { username: "refreshed@example.com" };
  const { calls, factory } = fakeMsal({
    accounts: [[stale], [refreshed]],
    silent: async (account) => {
      if (account === refreshed) return { accessToken: "refreshed-token" };
      throw new Error("cache_unavailable");
    },
  });

  const token = await createAuth({
    config: testConfig(),
    msal: factory,
  }).accessToken();

  assert.equal(token, "refreshed-token");
  assert.deepEqual(calls.silent, [stale, refreshed]);
  assert.equal(calls.accountReads, 2);
  assert.equal(calls.device, 0);
});

test("an exhausted retry reports the expired session", async () => {
  const { calls, factory } = fakeMsal({
    accounts: [[{ username: "someone@example.com" }]],
    silent: async () => {
      throw "expired";
    },
  });

  await assert.rejects(
    createAuth({ config: testConfig(), msal: factory }).accessToken(),
    (error: unknown) => {
      assert.ok(error instanceof AuthRequiredError);
      assert.equal(
        error.message,
        "Outlook session expired (expired). Run the login command again.",
      );
      assert.equal(error.cause, "expired");
      return true;
    },
  );
  assert.equal(calls.silent.length, 2);
  assert.equal(calls.accountReads, 2);
  assert.equal(calls.device, 0);
});

test("a grant that is merely too narrow is reported as needing re-authorisation", async () => {
  const { calls, factory } = fakeMsal({
    accounts: [[{ username: "someone@example.com" }]],
    silent: async (_account, scopes) => {
      if (scopes.includes("Mail.ReadWrite"))
        throw new Error(
          "AADSTS65001: The user has not requested the applicable OAuth scopes.",
        );
      return { accessToken: "readonly-token" };
    },
  });

  await assert.rejects(
    createAuth({ config: testConfig(), msal: factory }).accessToken(),
    (error: unknown) => {
      assert.ok(error instanceof AuthRequiredError);
      assert.match(error.message, /needs re-authorisation/);
      // The permissions wanted are named, because that is the whole diagnosis.
      assert.match(error.message, /Mail\.ReadWrite/);
      assert.match(error.message, /AADSTS65001/);
      assert.equal(
        error.message.toLowerCase().includes("expired"),
        false,
        "a session that is too narrow is not an expired one",
      );
      return true;
    },
  );
  assert.equal(calls.device, 0);
});

test("auth status asks for a token, because a cached account is not a session", async () => {
  const home = tempDir();
  const account = { username: "someone@example.com" };
  // What a mailbox signed in before `Mail.Read` became `Mail.ReadWrite` really
  // looks like: the account is cached, a silent token for the wider scope is
  // refused, and a status that read only the account called that signed in - so
  // login said "already signed in", the call that needed the scope failed
  // again, and the agent looped.
  const { calls, factory } = fakeMsal({
    accounts: [[account]],
    silent: async (_account, scopes) => {
      if (scopes.includes("Mail.ReadWrite"))
        throw new Error(
          "AADSTS65001: The user has not requested the applicable OAuth scopes.",
        );
      return { accessToken: "readonly-token" };
    },
  });
  const drafts = loadConfig({ HOME: home, OUTLOOK_OAUTH_CLIENT_ID: clientId });
  const readonly = loadConfig({
    HOME: home,
    OUTLOOK_OAUTH_CLIENT_ID: clientId,
    OUTLOOK_MCP_READONLY: "1",
  });

  assert.deepEqual(
    await createAuth({ config: drafts, msal: factory }).authStatus(),
    {
      signedIn: false,
      account: "someone@example.com",
      reason: "needsConsent",
      authority: consumers,
      cachePath: drafts.cachePath,
    },
  );

  // The same account against the scope it does hold is a session after all, so
  // the answer comes from the token, not from a guess.
  assert.deepEqual(
    await createAuth({ config: readonly, msal: factory }).authStatus(),
    {
      signedIn: true,
      account: "someone@example.com",
      authority: consumers,
      cachePath: readonly.cachePath,
    },
  );
  assert.deepEqual(calls.scopes, [
    ["User.Read", "Mail.ReadWrite", "email", "profile"],
    ["User.Read", "Mail.Read", "email", "profile"],
  ]);
});

test("auth status calls a refresh that expired what it is", async () => {
  const { factory } = fakeMsal({
    accounts: [[{ username: "someone@example.com" }]],
    silent: async () => {
      throw new Error(
        "AADSTS700082: The refresh token has been due to inactivity.",
      );
    },
  });

  const status = await createAuth({
    config: testConfig(),
    msal: factory,
  }).authStatus();
  assert.equal(status.signedIn, false);
  assert.equal(status.reason, "expired");
  assert.equal(status.account, "someone@example.com");
});

test("interactive login falls back to the device code flow", async () => {
  const { calls, factory } = fakeMsal({
    accounts: [[{ username: "someone@example.com" }]],
    silent: async () => {
      throw new Error("interaction_required");
    },
  });
  const auth = createAuth({ config: testConfig(), msal: factory });

  const lines = await captureStderr(async () => {
    assert.equal(await auth.accessToken({ interactive: true }), "device-token");
  });

  assert.equal(calls.device, 1);
  assert.equal(calls.accountReads, 1);
  assert.deepEqual(lines, ["To sign in, open the verification URL"]);
});

test("a missing session is reported without prompting", async () => {
  const { calls, factory } = fakeMsal({ accounts: [[]] });

  await assert.rejects(
    createAuth({ config: testConfig(), msal: factory }).accessToken(),
    (error: unknown) => {
      assert.ok(error instanceof AuthRequiredError);
      assert.equal(
        error.message,
        "No Outlook session. Run the login command first.",
      );
      return true;
    },
  );
  assert.equal(calls.device, 0);
  assert.equal(calls.silent.length, 0);
});

test("a device flow without a token fails clearly", async () => {
  const { factory } = fakeMsal({ accounts: [[]], device: async () => null });

  await assert.rejects(
    createAuth({ config: testConfig(), msal: factory }).accessToken({
      interactive: true,
    }),
    /Microsoft did not return an access token\./,
  );
});

test("the client id is required before Microsoft is contacted", async () => {
  assert.throws(
    () => createMsal(loadConfig({})),
    /Set OUTLOOK_OAUTH_CLIENT_ID in the environment or \.env file\./,
  );
  await assert.rejects(
    createAuth({ config: loadConfig({}) }).accessToken(),
    /Set OUTLOOK_OAUTH_CLIENT_ID in the environment or \.env file\./,
  );
});

test("auth status reports the cached account, authority and cache path", async () => {
  const home = tempDir();
  const config = loadConfig({
    HOME: home,
    OUTLOOK_OAUTH_CLIENT_ID: clientId,
  });
  assert.equal(config.authority, consumers);
  assert.equal(config.cachePath, join(home, ".outlook-mcp", "msal-cache.json"));
  assert.equal(config.stateDir, join(home, ".outlook-mcp", "state"));

  const signedIn = fakeMsal({
    accounts: [[{ username: "someone@example.com" }]],
  });
  assert.deepEqual(
    await createAuth({ config, msal: signedIn.factory }).authStatus(),
    {
      signedIn: true,
      account: "someone@example.com",
      authority: consumers,
      cachePath: config.cachePath,
    },
  );

  const work = loadConfig({
    HOME: tempDir(),
    OUTLOOK_OAUTH_CLIENT_ID: clientId,
    OUTLOOK_OAUTH_TENANT: "organizations",
  });
  const signedOut = fakeMsal({ accounts: [[]] });
  assert.deepEqual(
    await createAuth({ config: work, msal: signedOut.factory }).authStatus(),
    {
      signedIn: false,
      account: undefined,
      authority: "https://login.microsoftonline.com/organizations/",
      cachePath: work.cachePath,
    },
  );
});

test("drafts need Mail.ReadWrite and readonly asks for Mail.Read only", async () => {
  const account = { username: "someone@example.com" };

  const readWrite = fakeMsal({ accounts: [[account]] });
  await createAuth({
    config: testConfig(),
    msal: readWrite.factory,
  }).accessToken();
  assert.deepEqual(readWrite.calls.scopes, [
    ["User.Read", "Mail.ReadWrite", "email", "profile"],
  ]);

  const readonly = fakeMsal({ accounts: [[account]] });
  await createAuth({
    config: testConfig({ OUTLOOK_MCP_READONLY: "1" }),
    msal: readonly.factory,
  }).accessToken();
  assert.deepEqual(readonly.calls.scopes, [
    ["User.Read", "Mail.Read", "email", "profile"],
  ]);
});

test("logout is safe when the cache is present or absent", () => {
  const config = testConfig();
  logout(config);
  mkdirSync(dirname(config.cachePath), { recursive: true });
  writeFileSync(config.cachePath, "session");
  logout(config);
  assert.equal(existsSync(config.cachePath), false);
});

test("logout takes the record of a failed sign-in with the session", () => {
  // A lastLoginError outlives the session it belongs to otherwise, and the next
  // agent to ask auth_status reads an old failure as a fresh one.
  const config = testConfig();
  writeLoginState(config, { state: "error", message: "no client id" });
  assert.equal(readLoginState(config)?.state, "error");

  logout(config);

  assert.equal(readLoginState(config), undefined);
});

test("a plaintext cache copy left by a killed write is swept on the next read", async () => {
  const cache = tempCache();
  mkdirSync(dirname(cache), { recursive: true });
  const stale = join(dirname(cache), ".msal-cache.4242.dead.tmp");
  const fresh = join(dirname(cache), ".msal-cache.9999.alive.tmp");
  const other = join(dirname(cache), "notes.txt");
  for (const path of [stale, fresh, other]) writeFileSync(path, "token cache");
  const anHourAgo = new Date(Date.now() - 2 * 60 * 60 * 1000);
  utimesSync(stale, anHourAgo, anHourAgo);

  await createCachePlugin(cache).beforeCacheAccess({
    tokenCache: { deserialize: () => {} },
  } as never);

  assert.equal(existsSync(stale), false, "the abandoned copy is gone");
  assert.equal(existsSync(fresh), true, "another client's write is left alone");
  assert.equal(
    existsSync(other),
    true,
    "nothing else in the directory is gone",
  );
});

test("a session that could not be checked is not called a consent problem", async () => {
  // Labelling every non-expiry failure needsConsent sent agents to re-consent on
  // a machine whose session was fine - a laptop asleep, a DNS failure, a Microsoft
  // incident - and each of those consent attempts spends a rate-limited device
  // code. needsConsent now requires Microsoft to have said so.
  const home = tempDir();
  const account = { username: "someone@example.com" };
  const { factory } = fakeMsal({
    accounts: [[account]],
    silent: async () => {
      throw new Error(
        "getaddrinfo ENOTFOUND login.microsoftonline.com while acquiring a token",
      );
    },
  });
  const config = loadConfig({ HOME: home, OUTLOOK_OAUTH_CLIENT_ID: clientId });

  const status = await createAuth({ config, msal: factory }).authStatus();
  assert.equal(status.signedIn, false);
  assert.equal(status.reason, "unavailable");

  // And what a caller is told must not advise the thing it must not do.
  await assert.rejects(
    createAuth({ config, msal: factory }).accessToken({ interactive: false }),
    (error: Error) =>
      /Could not confirm the Outlook session/.test(error.message) &&
      /no re-consent is needed/.test(error.message),
  );
});

test("the three reasons separate an outage from expiry from a narrow grant", async () => {
  const reasons = [
    ["invalid_grant Error(s): 70000", "expired"],
    ["AADSTS700082: The refresh token has expired", "expired"],
    ["AADSTS65001: the user has not consented", "needsConsent"],
    ["interaction_required: consent required", "needsConsent"],
    ["consent_required", "needsConsent"],
    ["getaddrinfo ENOTFOUND login.microsoftonline.com", "unavailable"],
    ["ECONNRESET while contacting the token endpoint", "unavailable"],
    ["server_error: something happened", "unavailable"],
  ] as const;
  for (const [failure, expected] of reasons) {
    assert.equal(
      sessionProblem(failure),
      expected,
      `${failure} was labelled wrongly`,
    );
  }
});

test("createAuth hands back only what a caller uses", () => {
  // The config is the caller's own: a copy on the return value was never read.
  const { factory } = fakeMsal({ accounts: [[]] });
  const auth = createAuth({ config: testConfig(), msal: factory });
  assert.deepEqual(Object.keys(auth).sort(), ["accessToken", "authStatus"]);
});
