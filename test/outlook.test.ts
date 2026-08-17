import assert from "node:assert/strict";
import { expect, mock, test } from "bun:test";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

const state = {
  accounts: [] as unknown[],
  silent: async () => ({ accessToken: "silent-token" }),
  device: async () => ({ accessToken: "device-token" }),
  config: undefined as unknown,
  authProvider: undefined as
    | undefined
    | ((done: (error: Error | null, token: string | null) => void) => void),
  requests: [] as Array<{
    path: string;
    headers: string[][];
    search?: string;
    filter?: string;
    top?: number;
    select?: string[];
  }>,
};

class PublicClientApplication {
  constructor(config: unknown) {
    state.config = config;
  }
  getTokenCache() {
    return { getAllAccounts: async () => state.accounts };
  }
  acquireTokenSilent() {
    return state.silent();
  }
  acquireTokenByDeviceCode({
    deviceCodeCallback,
  }: {
    deviceCodeCallback: (response: { message: string }) => void;
  }) {
    deviceCodeCallback({ message: "device code" });
    return state.device();
  }
}

mock.module("@azure/msal-node", () => ({ PublicClientApplication }));
mock.module("@microsoft/microsoft-graph-client", () => ({
  Client: {
    init: ({ authProvider }: { authProvider: typeof state.authProvider }) => {
      state.authProvider = authProvider;
      return {
        api: (path: string) => {
          const request: {
            path: string;
            headers: string[][];
            search?: string;
            filter?: string;
            top?: number;
            select?: string[];
          } = { path, headers: [] };
          state.requests.push(request);
          const chain = {
            header: (name: string, value: string) => (
              request.headers.push([name, value]),
              chain
            ),
            select: (fields: string[]) => ((request.select = fields), chain),
            orderby: () => chain,
            top: (value: number) => ((request.top = value), chain),
            filter: (value: string) => ((request.filter = value), chain),
            search: (value: string) => ((request.search = value), chain),
            get: async () =>
              path.includes("/messages/")
                ? { id: "message" }
                : { value: ["message"] },
          };
          return chain;
        },
      };
    },
  },
}));

const {
  accessToken,
  createCachePlugin,
  graphClient,
  limit,
  listEmails,
  logout,
  main,
  msalClient,
  searchEmails,
  getEmail,
} = await import("../scripts/outlook.ts");
const tempCache = () =>
  join(
    mkdtempSync(join(tmpdir(), "outlook-skill-")),
    "cache",
    "msal-cache.json",
  );

function reset() {
  state.accounts = [];
  state.silent = async () => ({ accessToken: "silent-token" });
  state.device = async () => ({ accessToken: "device-token" });
  state.requests = [];
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

  expect(restored).toEqual(["saved"]);
  expect(readFileSync(cache, "utf8")).toBe("updated");
  expect(statSync(dirname(cache)).mode & 0o777).toBe(0o700);
  expect(statSync(cache).mode & 0o777).toBe(0o600);
});

test("cache plugin does not write unchanged caches", async () => {
  const cache = tempCache();
  const plugin = createCachePlugin(cache);
  await plugin.beforeCacheAccess({
    tokenCache: { deserialize: () => expect.unreachable() },
  } as never);
  await plugin.afterCacheAccess({
    cacheHasChanged: false,
    tokenCache: { serialize: () => expect.unreachable() },
  } as never);
  expect(existsSync(cache)).toBe(false);
});

test("logout is safe when the cache is present or absent", () => {
  const cache = tempCache();
  logout(cache);
  mkdirSync(dirname(cache), { recursive: true });
  writeFileSync(cache, "session");
  logout(cache);
  expect(existsSync(cache)).toBe(false);
});

test("access tokens use silent refresh, device login, and useful errors", async () => {
  reset();
  state.accounts = [{}];
  expect(await accessToken()).toBe("silent-token");

  state.silent = async () => {
    throw "expired";
  };
  await assert.rejects(accessToken(), /Outlook session expired \(expired\)/);
  expect(await accessToken(true)).toBe("device-token");

  state.accounts = [];
  await assert.rejects(accessToken(), /No Outlook session/);
  state.device = async () => undefined as never;
  await assert.rejects(
    accessToken(true),
    /Microsoft did not return an access token\./,
  );
});

test("MSAL and Graph clients use the required configuration and auth callback", async () => {
  reset();
  const clientId = Bun.env.OUTLOOK_OAUTH_CLIENT_ID;
  try {
    delete Bun.env.OUTLOOK_OAUTH_CLIENT_ID;
    expect(() => msalClient()).toThrow("Set OUTLOOK_OAUTH_CLIENT_ID");
  } finally {
    if (clientId) Bun.env.OUTLOOK_OAUTH_CLIENT_ID = clientId;
    else delete Bun.env.OUTLOOK_OAUTH_CLIENT_ID;
  }
  msalClient();
  expect(state.config).toMatchObject({
    auth: { authority: "https://login.microsoftonline.com/consumers/" },
  });

  state.accounts = [{}];
  graphClient();
  const token = await new Promise<string>((resolve, reject) =>
    state.authProvider!((error, value) =>
      error ? reject(error) : resolve(value!),
    ),
  );
  expect(token).toBe("silent-token");

  state.accounts = [];
  graphClient();
  await assert.rejects(
    new Promise((resolve, reject) =>
      state.authProvider!((error, value) =>
        error ? reject(error) : resolve(value),
      ),
    ),
    /No Outlook session/,
  );
});

test("mail requests select, filter, escape, and limit the documented data", async () => {
  reset();
  expect(await listEmails(true, 2)).toEqual(["message"]);
  expect(await getEmail("a/b")).toEqual({ id: "message" });
  expect(await searchEmails('a"b\\c', 3)).toEqual(["message"]);

  expect(state.requests).toEqual([
    expect.objectContaining({
      path: "/me/mailFolders/inbox/messages",
      filter: "isRead eq false",
      top: 2,
    }),
    expect.objectContaining({
      path: "/me/messages/a%2Fb",
      headers: [["Prefer", 'outlook.body-content-type="text"']],
    }),
    expect.objectContaining({
      path: "/me/messages",
      headers: [["ConsistencyLevel", "eventual"]],
      search: '"a\\"b\\\\c"',
      top: 3,
    }),
  ]);
});

test("limits and CLI actions use the correct command maxima", async () => {
  expect(limit("1", 10)).toBe(1);
  for (const value of ["0", "11", "1.5", "nope"])
    expect(() => limit(value, 10)).toThrow(
      "--limit must be an integer from 1 to 10.",
    );

  reset();
  const output: string[] = [];
  const originalLog = console.log;
  console.log = (...args: unknown[]) => output.push(args.join(" "));
  try {
    await main(["bun", "outlook", "login"]);
    await main(["bun", "outlook", "emails", "list", "--limit", "100"]);
    await main(["bun", "outlook", "emails", "get", "id"]);
    await main([
      "bun",
      "outlook",
      "emails",
      "search",
      "query",
      "--limit",
      "1000",
    ]);
  } finally {
    console.log = originalLog;
  }
  expect(output).toHaveLength(3);
  expect(state.requests.map((request) => request.top)).toEqual([
    100,
    undefined,
    1000,
  ]);
});
