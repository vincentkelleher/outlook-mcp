import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { homeDir } from "./config.ts";
import type { Env } from "./config.ts";

/**
 * Per-client MCP configuration: the single place that knows each client's file
 * format, so the printed snippet and the written file cannot drift.
 * `docs/clients.md` is the prose version of this table.
 *
 * Provenance is stated directly, per entry: a shape is presented as fact only
 * when it was read in that client's own official documentation, what was not
 * read there travels in `unverified`, and `config` says the difference out loud
 * rather than presenting convention as fact. No design record stands behind any
 * of it - what the client documents today is the only authority here.
 */

/** One server name everywhere: Pi and Claude Code enforce this character set. */
const serverName = "outlook";
const envName = "OUTLOOK_OAUTH_CLIENT_ID";
/** Stands in for a client id the process does not have. Never a secret. */
const clientIdPlaceholder = "your-application-client-id";

export const clientIds = [
  "pi",
  "claude-code",
  "claude-desktop",
  "codex",
  "vscode",
  "cursor",
  "gemini-cli",
  "opencode",
  "generic",
] as const;

export type ClientId = (typeof clientIds)[number];

/** A resolved `RenderInput`: nothing optional is left. */
type Resolved = {
  entry: string;
  clientId: string | undefined;
  home: string;
  appData: string;
  platform: NodeJS.Platform;
};

export type RenderInput = {
  /** Absolute path of the entry the client should launch. */
  entry: string;
  clientId?: string;
  env?: Env;
  platform?: NodeJS.Platform;
};

type Rendered = {
  label: string;
  /** Where the snippet belongs, in words. */
  destination: string;
  /** Exactly what `--write` merges, which is why print and write agree. */
  snippet: string;
  notes: string[];
  /** The user configuration file `--write` merges into, when one can be named. */
  path?: string;
  kind: "json" | "toml";
  /** Key holding the server map in a JSON file. */
  rootKey?: string;
  /** The object placed under the server name, kept for the merge. */
  entry?: Record<string, unknown>;
};

export type WriteOutcome = {
  path: string;
  status: "created" | "updated" | "unchanged";
};

type Definition = {
  label: string;
  destination: string;
  notes: string[];
  /**
   * What was not checked against this client's own official documentation, and
   * so must be said out loud. Absence is the record of a verification: the shape
   * was read there, and re-reading it is what changes this entry.
   */
  unverified?: string;
  /** Absolute user configuration file, when the client has one we can name. */
  path?: (target: Resolved) => string;
  /** How that file stores a server: Codex is TOML, the rest JSON. */
  file?: { kind: "json"; rootKey: string } | { kind: "toml" };
  /** The object placed under the server name in a JSON file. */
  entry?: (target: Resolved) => Record<string, unknown>;
  /** For clients whose documented surface is a shell command (Gemini CLI). */
  command?: (target: Resolved) => string;
};

/** The client id as stored, per what the client documents. */
function envValue(target: Resolved, expand: boolean): string {
  if (expand) return "${" + envName + "}";
  return target.clientId || clientIdPlaceholder;
}

function stdioEntry(
  target: Resolved,
  expand: boolean,
): Record<string, unknown> {
  return {
    command: "node",
    args: [target.entry],
    env: { [envName]: envValue(target, expand) },
  };
}

/**
 * Where a desktop client keeps its user configuration: under the roaming app
 * data on Windows, under `Library/Application Support` on the Mac, under
 * `.config` everywhere else. The client contributes only the names below it, so
 * one platform rule is written once.
 */
function appConfigPath(
  directory: string,
  ...tail: string[]
): (target: Resolved) => string {
  return (target) =>
    target.platform === "win32"
      ? join(target.appData, directory, ...tail)
      : target.platform === "darwin"
        ? join(
            target.home,
            "Library",
            "Application Support",
            directory,
            ...tail,
          )
        : join(target.home, ".config", directory, ...tail);
}

/** Paths arrive with backslashes on Windows: TOML basic strings must keep them. */
function tomlString(value: string): string {
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/** Codex documents no interpolation in TOML, so the variable is forwarded. */
function codexTable(target: Resolved): string {
  return [
    `[mcp_servers.${serverName}]`,
    `command = ${tomlString("node")}`,
    `args = [${tomlString(target.entry)}]`,
    `env_vars = [${tomlString(envName)}]`,
    // Codex gives a server 10 seconds to start, which a first cold start can
    // spend; the 60 second tool default is left alone because a read or a
    // draft finishes well inside it.
    "startup_timeout_sec = 30",
  ].join("\n");
}

const definitions: Record<ClientId, Definition> = {
  pi: {
    label: "Pi",
    destination: "Pi's MCP servers",
    notes: [
      'Pi defaults MCP tools to codemode, so this entry keeps "exposure": "direct".',
      "Pi's timeout option counts seconds and defaults to 60.",
      "Pi is normally configured with pi mcp add, recorded in docs/clients.md, so --write has no Pi file to merge into.",
    ],
    file: { kind: "json", rootKey: "mcpServers" },
    entry: (target) => ({
      ...stdioEntry(target, true),
      description: "Read, search, and draft Outlook mail",
      exposure: "direct",
    }),
  },
  "claude-code": {
    label: "Claude Code",
    destination: "Claude Code's user-scope configuration",
    notes: [
      "Claude Code documents ${VAR} expansion in a project .mcp.json but not at user scope, so the client id is written as a literal value.",
    ],
    path: (target) => join(target.home, ".claude.json"),
    file: { kind: "json", rootKey: "mcpServers" },
    entry: (target) => stdioEntry(target, false),
  },
  "claude-desktop": {
    label: "Claude Desktop",
    destination: "Claude Desktop's configuration file",
    notes: [
      "Claude Desktop only reads this file when it starts, so reload the app afterwards.",
    ],
    unverified:
      "Unverified: this Claude Desktop file location was not checked against its official documentation.",
    path: appConfigPath("Claude", "claude_desktop_config.json"),
    file: { kind: "json", rootKey: "mcpServers" },
    entry: (target) => stdioEntry(target, false),
  },
  codex: {
    label: "Codex",
    destination: "Codex's global configuration file",
    notes: [
      `Codex documents no interpolation for TOML env, so ${envName} is forwarded with env_vars: export it before Codex starts, or write the literal value into an env table.`,
    ],
    path: (target) => join(target.home, ".codex", "config.toml"),
    file: { kind: "toml" },
  },
  vscode: {
    label: "VS Code / Copilot",
    destination: "the mcp.json in your VS Code user profile",
    notes: [
      'VS Code marks "type" required, so it is always written.',
      "Workspace-scoped servers need Workspace Trust, and Agent Host skips servers that use ${input:...}: this one asks for nothing.",
    ],
    unverified:
      "Unverified: the exact path of the user-profile mcp.json is not published, so check it in VS Code with the MCP: Open User Config command.",
    path: appConfigPath("Code", "User", "mcp.json"),
    file: { kind: "json", rootKey: "servers" },
    entry: (target) => ({ type: "stdio", ...stdioEntry(target, false) }),
  },
  cursor: {
    label: "Cursor",
    destination: "Cursor's user-level mcp.json",
    notes: [],
    unverified:
      "Unverified: this Cursor file location and entry shape were not checked against official documentation.",
    path: (target) => join(target.home, ".cursor", "mcp.json"),
    file: { kind: "json", rootKey: "mcpServers" },
    entry: (target) => stdioEntry(target, false),
  },
  "gemini-cli": {
    label: "Gemini CLI",
    destination: "a terminal, at user scope",
    notes: [
      "Gemini CLI keeps user-scope servers in .gemini/settings.json, which --write merges instead.",
    ],
    unverified:
      "Unverified: this gemini mcp add form was not checked against official documentation.",
    path: (target) => join(target.home, ".gemini", "settings.json"),
    file: { kind: "json", rootKey: "mcpServers" },
    entry: (target) => stdioEntry(target, false),
    command: (target) =>
      `gemini mcp add ${serverName} node ${target.entry} --scope user --env ${envName}=${envValue(target, false)}`,
  },
  opencode: {
    label: "OpenCode",
    destination:
      "OpenCode's configuration file, where a local server carries its command as a list",
    notes: [],
    unverified:
      "Unverified: this OpenCode entry shape and file location were not checked against official documentation.",
    path: (target) => join(target.home, ".config", "opencode", "opencode.json"),
    file: { kind: "json", rootKey: "mcp" },
    entry: (target) => ({
      type: "local",
      command: ["node", target.entry],
      environment: { [envName]: envValue(target, false) },
      enabled: true,
    }),
  },
  generic: {
    label: "Any other MCP client",
    destination:
      "any client that starts a stdio server, such as an in-house agent harness",
    notes: [
      `Replace ${envName} with the literal client id if the client cannot expand \${VAR}.`,
    ],
    unverified:
      "Unverified: this is the plain stdio shape, checked against no single client's documentation.",
    file: { kind: "json", rootKey: "mcpServers" },
    entry: (target) => stdioEntry(target, true),
  },
};

function isClientId(value: string): value is ClientId {
  return (clientIds as readonly string[]).includes(value);
}

/** The error names every client, so a typo is answerable without --help. */
export function assertClientId(value: string): ClientId {
  if (!isClientId(value))
    throw new Error(
      `Unknown client "${value}". Choose one of: ${clientIds.join(", ")}.`,
    );
  return value;
}

// The built entry is what every client should launch, and `dist` sits two
// levels below the package root from either entry point: src/core or dist/core.
const packageRoot = fileURLToPath(new URL("../../", import.meta.url));

export function resolveEntry(root = packageRoot): {
  path: string;
  warning?: string;
} {
  const built = join(root, "dist", "index.js");
  if (existsSync(built)) return { path: built };
  return {
    path: join(root, "src", "index.ts"),
    warning: `${built} does not exist, so the configuration points at the TypeScript sources, which need Node 22.18 or newer. Run npm run build for a path that works on Node 20.`,
  };
}

function resolveInput(input: RenderInput): Resolved {
  const env = input.env ?? process.env;
  return {
    entry: input.entry,
    clientId: input.clientId,
    home: homeDir(env),
    appData: env.APPDATA || join(homeDir(env), "AppData", "Roaming"),
    platform: input.platform ?? process.platform,
  };
}

export function renderClient(client: ClientId, input: RenderInput): Rendered {
  const definition = definitions[client];
  const target = resolveInput(input);
  const path = definition.file ? definition.path?.(target) : undefined;
  const rootKey =
    definition.file?.kind === "toml"
      ? undefined
      : (definition.file?.rootKey ?? "mcpServers");
  const entry = definition.entry?.(target);
  const command = definition.command?.(target);
  const notes = [...definition.notes];
  if (definition.unverified) notes.push(definition.unverified);
  if (path)
    notes.push(
      command
        ? `The same entry can be merged into ${path} with --write.`
        : `Run the same command with --write to merge it into ${path}.`,
    );
  if (!target.clientId)
    notes.push(
      `${envName} is not set, so the configuration carries the placeholder ${clientIdPlaceholder}.`,
    );
  return {
    label: definition.label,
    destination: definition.destination,
    snippet: command
      ? command
      : definition.file?.kind === "toml"
        ? codexTable(target)
        : JSON.stringify(
            { [rootKey ?? "mcpServers"]: { [serverName]: entry } },
            null,
            2,
          ),
    notes,
    path,
    kind: definition.file?.kind ?? "json",
    rootKey,
    entry,
  };
}

const unparsable =
  "This command never overwrites a configuration file it cannot parse: fix or move the file, then run it again.";

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function readFile(path: string): string | undefined {
  return existsSync(path) ? readFileSync(path, "utf8") : undefined;
}

function saveFile(path: string, text: string, isNew: boolean): void {
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  writeFileSync(path, text, isNew ? { mode: 0o600 } : undefined);
}

/** Preserves every unrelated key, and rewrites nothing when the entry matches. */
function mergeJson(
  path: string,
  rootKey: string,
  entry: Record<string, unknown>,
): WriteOutcome {
  const existing = readFile(path);
  let file: Record<string, unknown> = {};
  if (existing !== undefined) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(existing);
    } catch (error) {
      throw new Error(
        `${path} is not valid JSON (${reasonOf(error)}). ${unparsable}`,
        { cause: error },
      );
    }
    if (!isRecord(parsed))
      throw new Error(
        `${path} holds ${typeof parsed} rather than a JSON object. ${unparsable}`,
      );
    file = parsed;
  }
  let servers: Record<string, unknown> = {};
  const stored = file[rootKey];
  if (stored !== undefined) {
    if (!isRecord(stored))
      throw new Error(
        `${path} has a "${rootKey}" entry that is not an object, so nothing can be merged into it. ${unparsable}`,
      );
    servers = stored;
  }
  if (JSON.stringify(servers[serverName]) === JSON.stringify(entry))
    return { path, status: "unchanged" };
  saveFile(
    path,
    `${JSON.stringify({ ...file, [rootKey]: { ...servers, [serverName]: entry } }, null, 2)}\n`,
    existing === undefined,
  );
  return { path, status: existing === undefined ? "created" : "updated" };
}

const tomlHeader = /^\[\[?[^\]]+\]\]?$/;
const tomlAssignment = /^[A-Za-z0-9_.-]+\s*=/;

function quotesBalanced(line: string, quote: string): boolean {
  let count = 0;
  for (let index = 0; index < line.length; index++) {
    if (line[index] !== quote || line[index - 1] === "\\") continue;
    count++;
  }
  return count % 2 === 0;
}

/**
 * There is no TOML parser in the dependency set, so the merge accepts only what
 * it can recognise line by line: blanks, comments, table headers and
 * assignments. Anything else - a multi-line string included - stops the write.
 */
function assertMergeableToml(text: string, path: string): void {
  for (const [index, raw] of text.split(/\r?\n/).entries()) {
    const line = raw.trim();
    if (line === "" || line.startsWith("#")) continue;
    if (line.includes('"""') || line.includes("'''"))
      throw new Error(
        `${path} line ${index + 1} uses a multi-line string. ${unparsable}`,
      );
    if (tomlHeader.test(line)) continue;
    if (
      tomlAssignment.test(line) &&
      quotesBalanced(line, '"') &&
      quotesBalanced(line, "'")
    )
      continue;
    throw new Error(
      `${path} line ${index + 1} is neither a TOML table header nor an assignment. ${unparsable}`,
    );
  }
}

/** The own table is replaced in place, or appended; other tables are kept. */
function mergeToml(path: string, table: string): WriteOutcome {
  const existing = readFile(path);
  if (existing === undefined) {
    saveFile(path, `${table}\n`, true);
    return { path, status: "created" };
  }
  assertMergeableToml(existing, path);
  const header = table.split("\n")[0];
  const lines = existing.split("\n");
  const start = lines.findIndex((line) => line.trim() === header);
  let end = lines.length;
  if (start >= 0)
    for (let index = start + 1; index < lines.length; index++)
      if (tomlHeader.test(lines[index].trim())) {
        end = index;
        break;
      }
  const kept =
    start === -1
      ? existing.trimEnd()
      : [...lines.slice(0, start), ...lines.slice(end)].join("\n").trimEnd();
  const next = kept ? `${kept}\n\n${table}\n` : `${table}\n`;
  if (next === existing) return { path, status: "unchanged" };
  saveFile(path, next, false);
  return { path, status: "updated" };
}

/**
 * `--write` is idempotent: an entry that already matches is left alone, so a
 * second run cannot reshuffle the file. A file that does not parse is refused.
 */
export function installClient(
  client: ClientId,
  input: RenderInput,
): WriteOutcome {
  const rendered = renderClient(client, input);
  if (!rendered.path)
    throw new Error(
      `${rendered.label} has no configuration file this command can write to. Run the command without --write and paste the snippet.`,
    );
  if (rendered.kind === "toml")
    return mergeToml(rendered.path, rendered.snippet);
  return mergeJson(
    rendered.path,
    rendered.rootKey ?? "mcpServers",
    rendered.entry ?? {},
  );
}
