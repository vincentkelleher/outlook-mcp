#!/usr/bin/env bun

import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { PublicClientApplication, type ICachePlugin } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { Command } from "commander";

const scopes = ["User.Read", "Mail.Read", "email", "profile"];
const tokenCachePath = join(homedir(), ".outlook-skill", "msal-cache.json");
const messageFields =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments";
const fullMessageFields = `${messageFields},toRecipients,ccRecipients,body`;

export function createCachePlugin(cachePath = tokenCachePath): ICachePlugin {
  return {
    beforeCacheAccess: async (context) => {
      if (existsSync(cachePath))
        context.tokenCache.deserialize(readFileSync(cachePath, "utf8"));
    },
    afterCacheAccess: async (context) => {
      if (!context.cacheHasChanged) return;
      mkdirSync(dirname(cachePath), { recursive: true, mode: 0o700 });
      chmodSync(dirname(cachePath), 0o700);
      writeFileSync(cachePath, context.tokenCache.serialize(), { mode: 0o600 });
      chmodSync(cachePath, 0o600);
    },
  };
}

export function logout(cachePath = tokenCachePath) {
  rmSync(cachePath, { force: true });
}

export function msalClient(): PublicClientApplication {
  const clientId = Bun.env.OUTLOOK_OAUTH_CLIENT_ID;
  if (!clientId)
    throw new Error(
      "Set OUTLOOK_OAUTH_CLIENT_ID in the environment or .env file.",
    );

  return new PublicClientApplication({
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/consumers/",
    },
    cache: { cachePlugin: createCachePlugin() },
  });
}

export async function accessToken(interactive = false) {
  const client = msalClient();
  const account = (await client.getTokenCache().getAllAccounts())[0];

  if (account) {
    try {
      return (await client.acquireTokenSilent({ account, scopes })).accessToken;
    } catch (error) {
      if (!interactive) {
        const reason = error instanceof Error ? error.message : String(error);
        throw new Error(
          `Outlook session expired (${reason}). Run the login command again.`,
          { cause: error },
        );
      }
    }
  }

  if (!interactive)
    throw new Error("No Outlook session. Run the login command first.");
  const result = await client.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => console.error(response.message),
  });
  if (!result?.accessToken)
    throw new Error("Microsoft did not return an access token.");
  return result.accessToken;
}

export function graphClient() {
  return Client.init({
    authProvider: async (done) => {
      try {
        done(null, await accessToken());
      } catch (error) {
        done(error instanceof Error ? error : new Error(String(error)), null);
      }
    },
  });
}

export async function listEmails(unread = false, top = 10) {
  const request = graphClient()
    .api("/me/mailFolders/inbox/messages")
    .select(messageFields.split(","))
    .orderby("receivedDateTime desc")
    .top(top);
  return (await (unread ? request.filter("isRead eq false") : request).get())
    .value;
}

export function getEmail(id: string) {
  return graphClient()
    .api(`/me/messages/${encodeURIComponent(id)}`)
    .header("Prefer", 'outlook.body-content-type="text"')
    .select(fullMessageFields.split(","))
    .get();
}

export async function searchEmails(query: string, top = 10) {
  return (
    await graphClient()
      .api("/me/messages")
      .header("ConsistencyLevel", "eventual")
      .search(`"${query.replace(/[\\"]/g, "\\$&")}"`)
      .top(top)
      .select(messageFields.split(","))
      .get()
  ).value;
}

export function limit(value: string, maximum: number) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum) {
    throw new Error(`--limit must be an integer from 1 to ${maximum}.`);
  }
  return parsed;
}

export async function main(argv = Bun.argv) {
  const program = new Command()
    .name("outlook")
    .description("Read Outlook email");
  const emails = program.command("emails").description("List or search email");

  program
    .command("login")
    .description("Sign in and save the Outlook session")
    .action(async () => {
      await accessToken(true);
      console.error("Signed in.");
    });

  program
    .command("logout")
    .description("Delete the saved Outlook session")
    .action(() => {
      logout();
      console.error("Signed out.");
    });

  emails
    .command("list")
    .description("List Inbox emails")
    .option("--unread", "only unread emails")
    .option("--limit <number>", "maximum emails to return", "10")
    .action(async ({ unread, limit: value }) => {
      console.log(
        JSON.stringify(await listEmails(unread, limit(value, 100)), null, 2),
      );
    });

  emails
    .command("get <id>")
    .description("Get one email with its full text content")
    .action(async (id) =>
      console.log(JSON.stringify(await getEmail(id), null, 2)),
    );

  emails
    .command("search <query>")
    .description("Search the mailbox")
    .option("--limit <number>", "maximum emails to return", "10")
    .action(async (query, { limit: value }) => {
      console.log(
        JSON.stringify(await searchEmails(query, limit(value, 1000)), null, 2),
      );
    });

  await program.parseAsync(argv);
}

if (import.meta.main) {
  main().catch((error) => {
    console.error(`Error: ${error.message}`);
    process.exit(1);
  });
}
