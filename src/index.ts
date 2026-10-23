#!/usr/bin/env node
import { realpathSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { main } from "./cli.ts";
import { loadConfig } from "./core/config.ts";
import { errorFields, logger } from "./core/log.ts";
import { serve } from "./mcp.ts";
import { createToolDeps } from "./tools.ts";

const log = logger("outlook");

/**
 * Which command died, named as far as a log line should name it: `emails search`
 * says what was running, and stops there rather than copying a search string, a
 * message id or an attachment name into a file that outlives the run.
 */
function commandOf(argv: string[]): string {
  return argv
    .filter((argument) => !argument.startsWith("-"))
    .slice(0, 2)
    .join(" ");
}

// Node has no import.meta.main, and either path can arrive through a symlinked
// directory, so compare realpaths.
function isMainModule() {
  if (!process.argv[1]) return false;
  try {
    return (
      realpathSync(process.argv[1]) ===
      realpathSync(fileURLToPath(import.meta.url))
    );
  } catch {
    return false;
  }
}

// No arguments is the MCP server over stdio: stdout is then JSON-RPC and must
// stay clean. `serve`, with or without --http, is a commander command, because
// only the CLI can read its flags; anything else is the human-facing CLI too.
function servesMcp(argv: string[]) {
  return argv.length === 0;
}

if (isMainModule()) {
  const run = servesMcp(process.argv.slice(2))
    ? serve(createToolDeps(loadConfig()))
    : main();
  run.catch((error: unknown) => {
    // The sentence is for whoever is standing at the terminal and gets the exit
    // code; the line is for whoever reads the log afterwards and wants to know
    // which command produced it, at what version, and why.
    log.error("stopped with an error", {
      command: commandOf(process.argv.slice(2)),
      ...errorFields(error, log.detailed),
    });
    process.stderr.write(
      `Error: ${error instanceof Error ? error.message : String(error)}\n`,
    );
    process.exit(1);
  });
}
