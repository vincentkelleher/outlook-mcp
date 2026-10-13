import { Command } from "commander";
import { createAuth, logout } from "./core/auth.ts";
import {
  assertClientId,
  clientIds,
  installClient,
  renderClient,
  resolveEntry,
} from "./core/clients.ts";
import { loadConfig, loadEnvFile } from "./core/config.ts";
import {
  alreadySignedIn,
  beginLogin,
  runLogin,
  startLogin,
} from "./core/device-login.ts";
import { createGraph } from "./core/graph.ts";
import type { Graph } from "./core/graph.ts";
import type { AuthStatus } from "./core/auth.ts";

/**
 * Everything the CLI would otherwise take from the machine. `authStatus` is
 * injectable because the background sign-in has to answer "is there already a
 * session?" before it starts a detached process.
 */
export type CliDeps = {
  graph?: Graph;
  authStatus?: () => Promise<AuthStatus>;
};

function parseLimit(value: string, maximum: number) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum)
    throw new Error(`--limit must be an integer from 1 to ${maximum}.`);
  return parsed;
}

export async function main(
  argv: string[] = process.argv,
  deps: CliDeps = {},
): Promise<void> {
  const env = loadEnvFile();
  const config = loadConfig(env);
  const auth = createAuth({ config });
  const graph = deps.graph ?? createGraph({ token: () => auth.accessToken() });

  const program = new Command()
    .name("outlook")
    .description("Read Outlook email; with no arguments, serve MCP over stdio");
  const emails = program.command("emails").description("List or search email");

  program
    .command("login")
    .description("Sign in and save the Outlook session")
    .option(
      "--background",
      "sign in from a detached process and return as soon as the code exists",
    )
    .option("--json", "with --background, print the sign-in state as JSON")
    .action(async ({ background, json }) => {
      if (background) {
        // Same rule as the login tool: a session that exists needs no flow,
        // so a second `login --background` is a poll, not a second sign-in.
        const status = await (deps.authStatus ?? (() => auth.authStatus()))();
        const launch = status.signedIn
          ? alreadySignedIn(status.account)
          : await beginLogin(config, startLogin);
        if (json) console.log(JSON.stringify(launch));
        else console.error(launch.instructions);
        return;
      }
      await runLogin(config);
      console.error("Signed in.");
    });

  program
    .command("logout")
    .description("Delete the saved Outlook session")
    .action(() => {
      logout(config.cachePath);
      console.error("Signed out.");
    });

  program
    .command("config <client>")
    .description(
      `Print the MCP configuration for one of ${clientIds.join(
        ", ",
      )}, or install it with --write`,
    )
    .option(
      "--write",
      "merge the entry into the client's own configuration file",
    )
    .action((client: string, { write }: { write?: boolean }) => {
      const target = assertClientId(client);
      const entry = resolveEntry();
      if (entry.warning) console.error(`Warning: ${entry.warning}`);
      const input = { entry: entry.path, clientId: config.clientId, env };
      if (!write) {
        const rendered = renderClient(target, input);
        console.error(`${rendered.label}: ${rendered.destination}.`);
        console.log(rendered.snippet);
        for (const note of rendered.notes) console.error(note);
        return;
      }
      const outcome = installClient(target, input);
      console.log(
        outcome.status === "unchanged"
          ? `${outcome.path} already had this outlook server.`
          : `${outcome.path} ${outcome.status} with the outlook server.`,
      );
    });

  emails
    .command("list")
    .description("List Inbox emails")
    .option("--unread", "only unread emails")
    .option("--limit <number>", "maximum emails to return", "10")
    .action(async ({ unread, limit: value }) => {
      const { items } = await graph.listMessages({
        unread,
        top: parseLimit(value, 100),
      });
      console.log(JSON.stringify(items, null, 2));
    });

  emails
    .command("get <id>")
    .description("Get one email with its full text content")
    .action(async (id) =>
      console.log(JSON.stringify(await graph.getMessage(id), null, 2)),
    );

  emails
    .command("attachments <message-id>")
    .description("Get an email's attachments")
    .action(async (messageId) =>
      console.log(
        JSON.stringify(await graph.listAttachments(messageId), null, 2),
      ),
    );

  emails
    .command("search <query>")
    .description("Search the mailbox")
    .option("--limit <number>", "maximum emails to return", "10")
    .action(async (query, { limit: value }) => {
      const { items } = await graph.searchMessages({
        query,
        top: parseLimit(value, 1000),
      });
      console.log(JSON.stringify(items, null, 2));
    });

  await program.parseAsync(argv);
}
