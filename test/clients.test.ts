import assert from "node:assert/strict";
import {
  existsSync,
  mkdtempSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import {
  assertClientId,
  clientIds,
  installClient,
  renderClient,
  resolveEntry,
} from "../src/core/clients.ts";
import type { ClientId } from "../src/core/clients.ts";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const entry = "/opt/outlook-mcp/dist/index.js";
const clientId = "00000000-0000-0000-0000-000000000000";
const home = () => mkdtempSync(join(tmpdir(), "outlook-clients-"));

/** A throwaway home, so nothing here can touch a real client configuration. */
function target(overrides: { HOME?: string } = {}) {
  return {
    entry,
    clientId,
    env: { HOME: overrides.HOME ?? home() },
    platform: "darwin" as const,
  };
}

function snippet(client: ClientId, input = target()) {
  return renderClient(client, input).snippet;
}

/** A configuration file is a map of server maps, three levels deep. */
function parsed(
  client: ClientId,
): Record<string, Record<string, Record<string, unknown>>> {
  return JSON.parse(snippet(client));
}

function noteText(client: ClientId) {
  return renderClient(client, target()).notes.join("\n");
}

function messageOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
  assert.fail("the call was expected to throw");
}

/** Every client with a file `--write` can merge into. */
const writable = clientIds.filter(
  (client) => renderClient(client, target()).path !== undefined,
);

test("every client launches the built entry with node", () => {
  for (const client of clientIds) {
    const text = snippet(client);
    assert.ok(text.includes(entry), `${client} lost the entry path: ${text}`);
    assert.ok(
      text.includes("node"),
      `${client} lost the node command: ${text}`,
    );
  }
});

test("Pi keeps the tools declared to the model and forwards the variable", () => {
  const server = parsed("pi").mcpServers.outlook;
  assert.equal(server.exposure, "direct");
  assert.ok(server.description);
  assert.equal(
    (server.env as Record<string, string>).OUTLOOK_OAUTH_CLIENT_ID,
    "${OUTLOOK_OAUTH_CLIENT_ID}",
  );
});

test("Codex writes a TOML table that forwards the variable instead of expanding it", () => {
  const text = snippet("codex");
  assert.ok(text.startsWith("[mcp_servers.outlook]"), text);
  assert.ok(text.includes(`args = ["${entry}"]`), text);
  assert.ok(text.includes('env_vars = ["OUTLOOK_OAUTH_CLIENT_ID"]'), text);
  assert.ok(text.includes("startup_timeout_sec = 30"), text);
  // No interpolation is documented for Codex, so none may be emitted.
  assert.equal(text.includes("${"), false);
});

test("VS Code gets a top-level servers map with the required type", () => {
  const file = parsed("vscode");
  assert.equal(file.mcpServers, undefined);
  assert.equal(file.servers.outlook.type, "stdio");
});

test("a desktop client's file is found on every platform", () => {
  // The one platform rule the desktop entries share, checked at each root rather
  // than per client: Windows keeps app data elsewhere, the Mac under Library.
  const env = { HOME: "/home/someone", APPDATA: "/AppData/Roaming" };
  const pathOn = (client: ClientId, platform: NodeJS.Platform) =>
    renderClient(client, { entry, clientId, env, platform }).path;
  const support = join("Library", "Application Support");

  assert.equal(
    pathOn("claude-desktop", "win32"),
    join("/AppData/Roaming", "Claude", "claude_desktop_config.json"),
  );
  assert.equal(
    pathOn("claude-desktop", "darwin"),
    join("/home/someone", support, "Claude", "claude_desktop_config.json"),
  );
  assert.equal(
    pathOn("claude-desktop", "linux"),
    join("/home/someone", ".config", "Claude", "claude_desktop_config.json"),
  );
  assert.equal(
    pathOn("vscode", "win32"),
    join("/AppData/Roaming", "Code", "User", "mcp.json"),
  );
  assert.equal(
    pathOn("vscode", "darwin"),
    join("/home/someone", support, "Code", "User", "mcp.json"),
  );
  assert.equal(
    pathOn("vscode", "linux"),
    join("/home/someone", ".config", "Code", "User", "mcp.json"),
  );
});

test("OpenCode carries a local command as a list", () => {
  const server = parsed("opencode").mcp.outlook;
  assert.equal(server.type, "local");
  assert.deepEqual(server.command, ["node", entry]);
});

test("Gemini CLI prints its add form rather than a file", () => {
  assert.match(
    snippet("gemini-cli"),
    /^gemini mcp add outlook node \/opt\/outlook-mcp\/dist\/index\.js --scope user/,
  );
});

test("a client id the process does not have becomes a stated placeholder", () => {
  const rendered = renderClient("claude-code", {
    ...target(),
    clientId: undefined,
  });
  assert.ok(rendered.snippet.includes("your-application-client-id"));
  assert.match(rendered.notes.join("\n"), /placeholder/);
});

test("a shape nobody checked against official documentation says so", () => {
  for (const client of [
    "claude-desktop",
    "cursor",
    "gemini-cli",
    "opencode",
  ] as const)
    assert.match(noteText(client), /^Unverified:/m, client);
  // These shapes were read in the client's own official documentation, which is
  // why the entry carries no `unverified` note.
  for (const client of ["pi", "claude-code", "codex"] as const)
    assert.equal(noteText(client).includes("Unverified"), false, client);
  // VS Code's shape is documented, the absolute path of its user file is not.
  assert.match(noteText("vscode"), /^Unverified: the exact path/m);
});

/** Every TypeScript and Markdown file under `dir`. */
function textFiles(dir: string, found: string[] = []): string[] {
  for (const item of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, item.name);
    if (item.isDirectory()) textFiles(path, found);
    else if (/\.(?:ts|md)$/.test(item.name)) found.push(path);
  }
  return found;
}

test("the migration plan is history: not shipped, and cited by nothing", () => {
  // The plan records the design as it was decided, and it disagrees with the
  // code in places. Shipping it and pointing a reader at it for what is verified
  // would hand that drift to whoever trusted the citation, so provenance is
  // stated where it is decided instead: per entry in `src/core/clients.ts`.
  const manifest = JSON.parse(
    readFileSync(join(repo, "package.json"), "utf8"),
  ) as { files?: string[] };
  assert.deepEqual(manifest.files, [
    "dist",
    "skills",
    "README.md",
    "SECURITY.md",
    "CHANGELOG.md",
    "LICENSE",
  ]);

  // `test/` is out of the scan because this very test has to name the file.
  const readable = [
    ...textFiles(join(repo, "src")),
    ...textFiles(join(repo, "docs")).filter(
      (file) => !file.endsWith("plan-mcp-skill.md"),
    ),
    ...textFiles(join(repo, "skills")),
    join(repo, "README.md"),
    join(repo, "SECURITY.md"),
    join(repo, "CHANGELOG.md"),
  ];
  assert.equal(readable.length >= 15, true, "the shipped surface is read");
  assert.deepEqual(
    readable.filter((file) =>
      readFileSync(file, "utf8").includes("plan-mcp-skill.md"),
    ),
    [],
  );
});

test("the built entry wins, and its absence says the sources need a newer Node", () => {
  const root = home();
  const missing = resolveEntry(root);
  assert.equal(missing.path, join(root, "src", "index.ts"));
  assert.match(missing.warning ?? "", /npm run build/);

  mkdirSync(join(root, "dist"), { recursive: true });
  writeFileSync(join(root, "dist", "index.js"), "");
  assert.deepEqual(resolveEntry(root), {
    path: join(root, "dist", "index.js"),
  });
});

test("--write is idempotent for every client that has a file", () => {
  assert.deepEqual(writable, [
    "claude-code",
    "claude-desktop",
    "codex",
    "vscode",
    "cursor",
    "gemini-cli",
    "opencode",
  ]);
  for (const client of writable) {
    const input = target();
    const first = installClient(client, input);
    assert.equal(first.status, "created", client);
    assert.equal(existsSync(first.path), true, client);
    const written = readFileSync(first.path, "utf8");
    assert.ok(written.includes(entry), client);
    assert.equal(installClient(client, input).status, "unchanged", client);
    assert.equal(readFileSync(first.path, "utf8"), written, client);
  }
});

test("--write merges beside an entry it does not own", () => {
  const jsonInput = target();
  const jsonPath = renderClient("vscode", jsonInput).path!;
  mkdirSync(jsonPath.slice(0, jsonPath.lastIndexOf("/")), { recursive: true });
  writeFileSync(
    jsonPath,
    '{\n  "servers": {\n    "other": {\n      "x": 1\n    }\n  }\n}\n',
  );
  const outcome = installClient("vscode", jsonInput);
  const merged = JSON.parse(readFileSync(jsonPath, "utf8"));
  assert.deepEqual(Object.keys(merged.servers).sort(), ["other", "outlook"]);
  assert.equal(outcome.status, "updated");

  const tomlInput = target();
  const tomlPath = renderClient("codex", tomlInput).path!;
  mkdirSync(join(tomlInput.env.HOME!, ".codex"), { recursive: true });
  writeFileSync(
    tomlPath,
    '[mcp_servers.other]\ncommand = "y"\n\n[other]\nkey = 1\n',
  );
  installClient("codex", tomlInput);
  const toml = readFileSync(tomlPath, "utf8");
  assert.match(toml, /\[mcp_servers\.other\]\ncommand = "y"/);
  assert.match(toml, /\[other\]\nkey = 1/);

  // A changed entry rewrites its own table, once, and nothing else.
  installClient("codex", { ...tomlInput, entry: "/elsewhere/index.js" });
  const updated = readFileSync(tomlPath, "utf8");
  assert.equal(updated.includes(entry), false);
  assert.match(updated, /args = \["\/elsewhere\/index\.js"\]/);
  assert.equal(updated.match(/\[mcp_servers\.outlook\]/g)?.length, 1);
});

test("--write refuses a configuration file it cannot parse and leaves it alone", () => {
  const input = target();
  const created = installClient("claude-code", input);
  writeFileSync(created.path, '{ "mcpServers": {oops');
  assert.throws(
    () => installClient("claude-code", input),
    /never overwrites a configuration file it cannot parse/,
  );
  assert.equal(readFileSync(created.path, "utf8"), '{ "mcpServers": {oops');

  // A root key that is not a map cannot be merged into either.
  const clash = target();
  const clashPath = renderClient("cursor", clash).path!;
  mkdirSync(clashPath.slice(0, clashPath.lastIndexOf("/")), {
    recursive: true,
  });
  writeFileSync(clashPath, '{\n  "mcpServers": []\n}\n');
  assert.throws(() => installClient("cursor", clash), /is not an object/);
});

test("--write refuses a Codex file it cannot read as TOML and leaves it alone", () => {
  const input = target();
  mkdirSync(join(input.env.HOME!, ".codex"), { recursive: true });
  const multiline = 'name = """\nblock\n"""\n';
  const path = join(input.env.HOME!, ".codex", "config.toml");
  writeFileSync(path, multiline);
  assert.throws(() => installClient("codex", input), /multi-line string/);
  assert.equal(readFileSync(path, "utf8"), multiline);

  writeFileSync(path, "not toml at all\n");
  assert.throws(
    () => installClient("codex", input),
    /neither a TOML table header nor an assignment/,
  );
  assert.equal(readFileSync(path, "utf8"), "not toml at all\n");
});

test("--write refuses the clients that have no file to merge into", () => {
  for (const client of ["pi", "generic"] as const)
    assert.throws(
      () => installClient(client, target()),
      /no configuration file this command can write to/,
      client,
    );
});

test("an unknown client names every choice", () => {
  const message = messageOf(() => assertClientId("notepad"));
  assert.match(message, /^Unknown client "notepad"\./);
  for (const client of clientIds)
    assert.ok(message.includes(client), `${client} missing from: ${message}`);
  assert.equal(assertClientId("gemini-cli"), "gemini-cli");
});
