import { Command } from "commander";
import { createAuth, logout } from "./core/auth.ts";
import {
  assertClientId,
  clientIds,
  installClient,
  renderClient,
  resolveEntry,
} from "./core/clients.ts";
import { loadConfig, loadEnvFile, scopesFor } from "./core/config.ts";
import { mcpPath } from "./core/http.ts";
import { serve, startHttp } from "./mcp.ts";
import { createToolDeps } from "./tools.ts";
import {
  alreadySignedIn,
  beginLogin,
  runLogin,
  startLogin,
} from "./core/device-login.ts";
import {
  createGraph,
  defaultPageSize,
  maxPageSize,
  maxSearchPageSize,
} from "./core/graph.ts";
import type { Graph } from "./core/graph.ts";
import type { AuthStatus } from "./core/auth.ts";
import {
  configureLogging,
  fieldsForLogging,
  logger,
  stopwatch,
  verificationCode,
} from "./core/log.ts";

const log = logger("cli");

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

/**
 * The port `serve --http` starts on. `0` is accepted here and not in
 * `OUTLOOK_MCP_PORT`: a service must name the port people connect to, while a
 * flag that prints the address it actually bound cannot mislead anybody — and it
 * is how a test, or a second server on a busy machine, gets a free port.
 */
function parsePort(value: string): number {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 0 || parsed > 65535)
    throw new Error(
      `--port must be an integer from 0 to 65535, 0 meaning any free port, not "${value}".`,
    );
  return parsed;
}

/** `emails list`, `serve`, `config docker`: what the person actually typed. */
function commandPath(command: Command): string {
  const names: string[] = [];
  for (let at: Command | null = command; at; at = at.parent) {
    names.unshift(at.name());
  }
  return names.filter((name) => name !== "outlook").join(" ");
}

/**
 * What a command was asked for: its flags, and its positional arguments under
 * the names they were declared with, so `emails get AAMk…` reads as `id=` and
 * `fieldsForLogging` can shorten it at the default level.
 */
function requestFields(command: Command) {
  const named: Record<string, unknown> = { ...command.opts() };
  command.registeredArguments.forEach((argument, index) => {
    const value = command.processedArgs?.[index];
    if (value !== undefined) named[argument.name()] = value;
  });
  return fieldsForLogging(named, log.detailed);
}

export async function main(
  argv: string[] = process.argv,
  deps: CliDeps = {},
): Promise<void> {
  const env = loadEnvFile();
  const config = loadConfig(env);
  // Before the first command runs: `outlook emails list` is itself a way to find
  // out what is wrong, and its level has to be settled before it says anything.
  configureLogging({ level: config.logLevel, format: config.logFormat });
  const auth = createAuth({ config });
  const graph = deps.graph ?? createGraph({ token: () => auth.accessToken() });

  // The hook pair that starts and stops a command's clock. A command runs one at
  // a time in this process, so one variable is enough to time it.
  let started: ReturnType<typeof stopwatch> | undefined;
  const program = new Command()
    .name("outlook")
    .description("Read Outlook email; with no arguments, serve MCP over stdio")
    .hook("preAction", (_thisCommand, actionCommand) => {
      started = stopwatch();
      log.info(`command ${commandPath(actionCommand)}`, {
        ...requestFields(actionCommand),
      });
    })
    .hook("postAction", (_thisCommand, actionCommand) => {
      // `serve` never reaches this hook — it resolves only when the listener is
      // closed — which is why the listener logs its own stopping line.
      log.info(`command ${commandPath(actionCommand)} finished`, {
        ms: started?.() ?? 0,
      });
    });
  const emails = program.command("emails").description("List or search email");

  program
    .command("serve")
    .description("Serve MCP over stdio, or over Streamable HTTP with --http")
    .option("--http", `serve Streamable HTTP on ${mcpPath} instead of stdio`)
    .option(
      "--host <host>",
      "with --http: the interface to bind",
      config.httpHost,
    )
    .option(
      "--port <number>",
      `with --http: the port to bind, 0 for any free one (default ${config.httpPort})`,
      `${config.httpPort}`,
    )
    .action(
      async ({
        http,
        host,
        port,
      }: {
        http?: boolean;
        host: string;
        port: string;
      }) => {
        if (!http) {
          // The stdio server logs its own start line: here it would only repeat
          // the command line the hook already wrote, and stdout is already spoken
          // for by the transport.
          await serve(createToolDeps(config));
          return;
        }
        const wanted = parsePort(port);
        const listener = await startHttp(createToolDeps(config), {
          host,
          port: wanted,
        });
        // stderr: a container reads these logs and a person piping JSON does not
        // want the address in the stream.
        process.stderr.write(`Outlook MCP serving on ${listener.url}\n`);
        if (config.readonly)
          process.stderr.write(
            `Read-only mode: draft tools are not offered (${scopesFor(true).join(", ")}).\n`,
          );
        process.stderr.write(
          "Nothing on that port authenticates: keep it on loopback, or name every " +
            "Host you reach it by in OUTLOOK_MCP_ALLOWED_HOSTS.\n",
        );
        // No line here: the HTTP server logs the bind it actually has, with the
        // session limits around it, and `command serve` above already says which
        // flags this process was started with.
        // dumb-init, or any supervisor, sends SIGTERM to stop the container; the
        // sessions and the listener go down together rather than at the kernel's
        // convenience.
        for (const signal of ["SIGINT", "SIGTERM"] as const)
          process.on(signal, () => {
            void listener.close();
          });
        await listener.closed;
      },
    );

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
        // The flow itself runs in a detached child and logs there, so this is
        // the seam that says what the caller was told: a code, an existing
        // session, or a reason nothing started.
        log.info("background sign-in answered", {
          state: launch.state,
          userCode: verificationCode(launch.userCode, log.detailed),
        });
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
      logout(config);
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
      log.debug("writing a client's own configuration file", { client });
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
    .option(
      "--limit <number>",
      `maximum emails to return, 1 to ${maxPageSize}`,
      `${defaultPageSize}`,
    )
    .action(async ({ unread, limit: value }) => {
      const { items } = await graph.listMessages({
        unread,
        top: parseLimit(value, maxPageSize),
      });
      // stdout carries the JSON a person pipes on; the count is the same fact in
      // a line that does not disturb it.
      log.info("inbox read", { items: items.length });
      log.debug("inbox message ids", {
        ids: items.map((item) => (item as { id?: string }).id),
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
    .option(
      "--limit <number>",
      `maximum emails to return, 1 to ${maxSearchPageSize}`,
      `${defaultPageSize}`,
    )
    .action(async (query, { limit: value }) => {
      const { items } = await graph.searchMessages({
        query,
        top: parseLimit(value, maxSearchPageSize),
      });
      log.info("mailbox searched", { items: items.length });
      console.log(JSON.stringify(items, null, 2));
    });

  // A failure is reported once, by the entry point that turns it into a message
  // and an exit code, with the command named there: reporting it here too would
  // put the same sentence in the log twice, and the second one is the one a
  // person reads as a separate problem.
  await program.parseAsync(argv);
}
