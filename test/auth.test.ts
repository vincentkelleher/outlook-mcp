import assert from "node:assert/strict";
import { test } from "node:test";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import {
  accessToken,
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
import { AuthRequiredError } from "../src/core/errors.ts";

const clientId = "00000000-0000-0000-0000-000000000000";
const consumers = "https://login.microsoftonline.com/consumers/";

const tempDir = () => mkdtempSync(join(tmpdir(), "outlook-skill-"));
const tempCache = () => join(tempDir(), "msal-cache.json");
const testConfig = (env: Env = {}) =>
  loadConfig({ HOME: tempDir(), OUTLOOK_OAUTH_CLIENT_ID: clientId, ...env });

type Calls = {
  accountReads: number;
  silent: MsalAccount[];
  device: number;
};

function fakeMsal(behaviour: {
  accounts: MsalAccount[][];
  silent?: (account: MsalAccount) => Promise<MsalResult>;
  device?: () => Promise<MsalResult | null>;
}) {
  const calls: Calls = { accountReads: 0, silent: [], device: 0 };
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
    async acquireTokenSilent({ account }): Promise<MsalResult> {
      calls.silent.push(account);
      if (behaviour.silent) return behaviour.silent(account);
      return { accessToken: "silent-token" };
    },
    async acquireTokenByDeviceCode({
      deviceCodeCallback,
    }): Promise<MsalResult | null> {
      calls.device += 1;
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
    accessToken({}),
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
  assert.equal(
    config.cachePath,
    join(home, ".outlook-skill", "msal-cache.json"),
  );
  assert.equal(config.stateDir, join(home, ".outlook-skill", "state"));

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

test("logout is safe when the cache is present or absent", () => {
  const cache = tempCache();
  logout(cache);
  mkdirSync(dirname(cache), { recursive: true });
  writeFileSync(cache, "session");
  logout(cache);
  assert.equal(existsSync(cache), false);
});
