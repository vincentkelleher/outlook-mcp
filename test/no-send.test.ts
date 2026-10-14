import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/sdk/client";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { loadConfig, scopesFor } from "../src/core/config.ts";
import type { Config, Env } from "../src/core/config.ts";
import { createGraph } from "../src/core/graph.ts";
import { createServer } from "../src/mcp.ts";
import {
  draftToolNames,
  readToolNames,
  tools,
  toolsFor,
} from "../src/tools.ts";
import type { ToolDeps } from "../src/tools.ts";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const clientId = "00000000-0000-0000-0000-000000000000";

/**
 * A tool name is send-capable when it offers to transmit mail: send, reply or
 * forward, with or without a target. A name ending in `draft` only writes one,
 * which is the entire feature, so `create_reply_draft` is fine and `reply` is
 * not.
 */
const sendCapableName = /(^|_)(send|reply|forward)(_|$)/;
const draftOnlyName = /_draft$/;

/**
 * Graph's transmitting actions: `/send`, `/reply`, `/replyAll`, `/forward`. The
 * segment must start after a `/` and end there, so the create actions that keep
 * a thread - `/createReply`, `/createForward` - are not matched.
 */
const sendEndpoint =
  /\/(?:microsoft\.graph\.)?(?:replyAll|forwardAll|forward|reply|send)(?![A-Za-z0-9])/;

/** Every quoted or templated string in a source file that names a Graph path. */
function graphPathsOf(source: string): string[] {
  return [...source.matchAll(/(["'`])(\/[^"'`\n]*)\1/g)].map(
    ([, , path]) => path,
  );
}

function sourceFiles(dir: string, found: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) sourceFiles(path, found);
    else if (entry.name.endsWith(".ts")) found.push(path);
  }
  return found.sort();
}

function testConfig(t: { after: (fn: () => void) => void }, env: Env = {}) {
  const home = mkdtempSync(join(tmpdir(), "outlook-no-send-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  return loadConfig({ HOME: home, OUTLOOK_OAUTH_CLIENT_ID: clientId, ...env });
}

/**
 * A real graph object over no network: these tests read what the server
 * advertises, so a token must never be asked for.
 */
function depsFor(config: Config): ToolDeps {
  return {
    graph: createGraph({
      token: async () => {
        throw new Error("no Microsoft request belongs in this test");
      },
    }),
    config,
    authStatus: async () => ({
      signedIn: false,
      account: undefined,
      authority: config.authority,
      cachePath: config.cachePath,
    }),
    startLogin: () => {},
    logout: () => {},
  };
}

async function advertised(
  t: { after: (fn: () => void) => void },
  config: Config,
) {
  const server = createServer(depsFor(config));
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "no-send-test", version: "0.0.0" });
  await Promise.all([
    server.connect(serverTransport),
    client.connect(clientTransport),
  ]);
  t.after(() => void client.close());
  const listed = await client.listTools();
  return { client, names: listed.tools.map((tool) => tool.name).sort() };
}

test("the registry is the read tools plus seven draft tools, nothing else", () => {
  assert.equal(draftToolNames.length, 7);
  assert.deepEqual(
    tools.map((tool) => tool.name).sort(),
    [...readToolNames, ...draftToolNames].sort(),
  );
});

test("no registered tool name offers to send, reply or forward mail", () => {
  for (const name of tools.map((tool) => tool.name))
    assert.equal(
      sendCapableName.test(name) && !draftOnlyName.test(name),
      false,
      `${name} looks send-capable`,
    );

  for (const forbidden of [
    "send",
    "send_draft",
    "send_message",
    "reply",
    "reply_all",
    "forward",
    "forward_message",
  ])
    assert.equal(
      tools.some((tool) => tool.name === forbidden),
      false,
      `${forbidden} is registered`,
    );
});

test("no source file addresses a sending endpoint", () => {
  const files = sourceFiles(join(repo, "src"));
  assert.equal(files.length >= 7, true, "the whole of src is scanned");

  for (const file of files) {
    const source = readFileSync(file, "utf8");
    // Paths read as a whole first, so the report names what it found.
    for (const path of graphPathsOf(source))
      assert.equal(
        sendEndpoint.test(path),
        false,
        `${file} targets the sending endpoint ${path}`,
      );
    // Then the raw text, in case a path is assembled in pieces.
    assert.equal(
      sendEndpoint.test(source),
      false,
      `${file} builds a Graph request that transmits mail`,
    );
  }
});

test("the scan really distinguishes the create actions from the sending ones", () => {
  // A guard that cannot fail proves nothing, so it is aimed at known shapes.
  const allowed = [
    "/me/messages",
    "/me/messages/a%2Fb/createReply",
    "/me/messages/a%2Fb/createForward",
    "/me/messages/a%2Fb/attachments",
    "/me/mailFolders/drafts/messages",
  ];
  const forbidden = [
    "/me/messages/a%2Fb/send",
    "/me/messages/a%2Fb/microsoft.graph.send",
    "/me/messages/a%2Fb/reply",
    "/me/messages/a%2Fb/replyAll",
    "/me/messages/a%2Fb/forward",
  ];
  for (const path of allowed)
    assert.equal(sendEndpoint.test(path), false, `${path} was flagged`);
  for (const path of forbidden)
    assert.equal(sendEndpoint.test(path), true, `${path} was missed`);

  // The thread-preserving actions are what the draft tools actually call.
  const graphSource = readFileSync(join(repo, "src/core/graph.ts"), "utf8");
  assert.equal(graphPathsOf(graphSource).includes("/me/messages"), true);
  assert.match(graphSource, /\/createReply/);
  assert.match(graphSource, /\/createForward/);
});

test("the mutating Graph calls are exactly the six reviewed draft writes", () => {
  // A name regex cannot see a path assembled in pieces, so this closes the
  // gap from the other side: every request that changes the mailbox goes
  // through one of six call sites in src/core/graph.ts. A new one is not
  // automatically wrong, but it must be looked at, and this test is where
  // that conversation starts.
  const source = readFileSync(join(repo, "src/core/graph.ts"), "utf8");
  const verbs = [...source.matchAll(/\.(post|put|patch|delete)\(/g)]
    .map((match) => match[1])
    .sort();
  assert.deepEqual(
    verbs,
    ["delete", "patch", "post", "post", "post", "post"],
    "a new mutating Graph call appeared: a draft write may be fine, anything that transmits mail is not",
  );
});

test("readonly mode advertises no draft tool at all", async (t) => {
  const config = testConfig(t, { OUTLOOK_MCP_READONLY: "1" });
  const { client, names } = await advertised(t, config);

  // Pinned against the read half of the registry: a draft tool left in this mode
  // cannot match it, so there is nothing to re-check name by name.
  assert.deepEqual(names, [...readToolNames].sort());

  // What is left is read-only mail plus sign-in, which touches the local cache
  // and not the mailbox; without it a read-only server could not sign in.
  const listed = await client.listTools();
  assert.deepEqual(
    listed.tools
      .filter((tool) => tool.annotations?.readOnlyHint !== true)
      .map((tool) => tool.name)
      .sort(),
    ["login", "logout"],
  );
  assert.match(
    client.getInstructions() ?? "",
    /readonly mode.*exposes none of .*create_draft/s,
  );
  assert.deepEqual(
    toolsFor(config)
      .map((tool) => tool.name)
      .sort(),
    names,
    "the CLI and the server read the same filtered registry",
  );
});

test("the advertised tools and the advertised permissions are one decision", async (t) => {
  // A server must not offer create_draft on a token that only carries
  // Mail.Read, nor hold Mail.ReadWrite while hiding every draft tool. Both
  // sides read scopesFor(config.readonly); this says so out loud. The scope set
  // actually sent to Microsoft is asserted in test/auth.test.ts.
  for (const readonly of [true, false]) {
    const config = testConfig(t, readonly ? { OUTLOOK_MCP_READONLY: "1" } : {});
    const { client, names } = await advertised(t, config);
    const status = (await client.callTool({
      name: "auth_status",
      arguments: {},
    })) as unknown as { content: { text?: string }[] };
    const reported = JSON.parse(status.content[0]?.text ?? "{}") as {
      readonly?: boolean;
      scopes?: string[];
    };

    assert.equal(reported.readonly, readonly);
    assert.deepEqual(reported.scopes, scopesFor(readonly));
    assert.equal(reported.scopes?.includes("Mail.ReadWrite"), !readonly);
    assert.equal(
      names.some((name) => draftToolNames.includes(name)),
      !readonly,
      `draft tools with readonly=${readonly}`,
    );
  }
});

test("the instructions name every tool the server has, in both modes", async (t) => {
  // An agent reads the instructions instead of the registry, so prose that falls
  // behind it is a claim the server cannot keep: in read-only mode the draft
  // tools are named as the ones that are missing, in the other mode as the ones
  // that write.
  for (const readonly of [true, false]) {
    const config = testConfig(t, readonly ? { OUTLOOK_MCP_READONLY: "1" } : {});
    const { client } = await advertised(t, config);
    const instructions = client.getInstructions() ?? "";
    for (const name of tools.map((tool) => tool.name))
      assert.match(
        instructions,
        new RegExp(`\\b${name}\\b`),
        `${name} missing from the instructions (readonly=${readonly})`,
      );
    // And it promises nothing the registry does not hold, least of all a way to
    // send: `draftToolNames` is the whole write surface.
    for (const promised of ["send_message", "send_draft", "reply_all"])
      assert.equal(
        instructions.includes(promised),
        false,
        `the instructions name ${promised}, which is not registered`,
      );
  }
});

test("without readonly the seven draft tools are advertised with their limits", async (t) => {
  const config = testConfig(t);
  const { client, names } = await advertised(t, config);

  assert.deepEqual(names, [...readToolNames, ...draftToolNames].sort());
  assert.match(
    client.getInstructions() ?? "",
    /create_draft, create_reply_draft, create_forward_draft, update_draft, list_drafts, delete_draft and add_attachment/,
  );

  const create = (await client.listTools()).tools.find(
    (tool) => tool.name === "create_draft",
  );
  const schema = create?.inputSchema as {
    required?: string[];
    properties?: {
      to?: { type?: string; minItems?: number; items?: { format?: string } };
      bodyFormat?: { default?: string };
    };
  };
  // Recipients are enforced addresses, and the address list cannot be empty.
  assert.deepEqual(schema?.required, ["to"]);
  assert.equal(schema?.properties?.to?.type, "array");
  assert.equal(schema?.properties?.to?.minItems, 1);
  assert.equal(schema?.properties?.to?.items?.format, "email");
  assert.equal(schema?.properties?.bodyFormat?.default, "text");
});
