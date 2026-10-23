#!/usr/bin/env node
import { realpathSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { main } from "./cli.ts";
import { loadConfig } from "./core/config.ts";
import { serve } from "./mcp.ts";
import { createToolDeps } from "./tools.ts";

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

// No arguments preserves the stdio default for existing client configs.
// Explicit `serve` goes through the CLI to validate transport options.
function servesMcp(argv: string[]) {
  return argv.length === 0;
}

if (isMainModule()) {
  const run = servesMcp(process.argv.slice(2))
    ? serve(createToolDeps(loadConfig()))
    : main();
  run.catch((error: unknown) => {
    process.stderr.write(
      `Error: ${error instanceof Error ? error.message : String(error)}\n`,
    );
    process.exit(1);
  });
}
