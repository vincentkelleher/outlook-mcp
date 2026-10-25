import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { loadConfig, scopesFor } from "../src/core/config.ts";
import type { Config, Env } from "../src/core/config.ts";
import { createGraph } from "../src/core/graph.ts";
import { createServer } from "../src/mcp.ts";
import {
  draftToolNames,
  flagToolNames,
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

/**
 * Written out in full rather than read from the registry, on purpose. If the
 * write surface is derived from `draftTools`, moving a draft tool into the read
 * list moves both sides of every assertion below and the suite stays green while
 * read-only mode starts advertising a write. This is the one copy that cannot
 * move with the code.
 */
const draftWriteToolNames = [
  "add_attachment",
  "create_draft",
  "create_forward_draft",
  "create_reply_draft",
  "delete_draft",
  "list_drafts",
  "remove_attachment",
  "update_draft",
];

/**
 * The other mailbox write: a message's read flag. Nothing here transmits, but a
 * flag needs `Mail.ReadWrite`, so it belongs to the half read-only mode hides —
 * and a flag tool moved into the read list would be a write advertised by a
 * `Mail.Read` server, which is exactly what the assertions below catch.
 */
const flagWriteToolNames = ["mark_read", "mark_unread"];

test("the registry is the read tools, two flag tools and eight draft tools, nothing else", () => {
  assert.deepEqual(
    [...draftToolNames].sort(),
    draftWriteToolNames,
    "the draft surface changed: read-only mode and every assertion below depends on this list",
  );
  assert.deepEqual(
    [...flagToolNames].sort(),
    flagWriteToolNames,
    "the flag surface changed: read-only mode and every assertion below depends on this list",
  );
  assert.equal(draftToolNames.length, 8);
  assert.equal(flagToolNames.length, 2);
  assert.deepEqual(
    tools.map((tool) => tool.name).sort(),
    [...readToolNames, ...flagToolNames, ...draftToolNames].sort(),
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

test("the mutating Graph calls are exactly the nine reviewed mailbox writes", () => {
  // A name regex cannot see a path assembled in pieces, so this closes the
  // gap from the other side: every request that changes the mailbox goes
  // through one of nine call sites in src/core/graph.ts — eight on a draft (two
  // on its attachments) and one PATCH of `isRead`. A new one is not
  // automatically wrong, but it must be looked at, and this test is where that
  // conversation starts. The tenth write, an upload session chunk, is not a
  // builder call at all and is pinned on its own below.
  //
  // A type argument on the call is matched, not skipped, because that is how a
  // write hides from this count: `.patch<T>({ isRead: read })` was six calls for
  // as long as the pattern asked for a parenthesis straight after the verb.
  const source = readFileSync(join(repo, "src/core/graph.ts"), "utf8");
  const verbs = [
    ...source.matchAll(/\.(post|put|patch|delete)\s*(?:<[^()]*>)?\s*\(/g),
  ]
    .map((match) => match[1])
    .sort();
  assert.deepEqual(
    verbs,
    [
      "delete",
      "delete",
      "patch",
      "patch",
      "post",
      "post",
      "post",
      "post",
      "post",
    ],
    "a new mutating Graph call appeared: a draft write or a read flag may be fine, anything that transmits mail is not",
  );
});

test("the write the verb scan cannot see is one chunk PUT with no Authorization header", () => {
  // An upload session URL is pre-authenticated — it carries its own bearer token
  // — so the chunk cannot go through the SDK builder, which adds an
  // `Authorization` header to everything it sends. It goes out over `fetch`,
  // which the `.post|.put|.patch|.delete` scan above cannot see. So the one
  // mailbox write that is not a builder call is named here: exactly one, and no
  // Authorization header anywhere near it, because Graph says to send none.
  const source = readFileSync(join(repo, "src/core/graph.ts"), "utf8");
  const chunks = [...source.matchAll(/await put\(/g)];
  assert.equal(
    chunks.length,
    1,
    "the upload session chunk write moved, and this is where it is accounted for",
  );
  const chunkCall = source.slice(chunks[0].index, chunks[0].index + 600);
  assert.match(
    chunkCall,
    /method:\s*"PUT"/,
    "the one chunk write is not a PUT",
  );
  assert.equal(
    /Authorization/.test(chunkCall),
    false,
    "a chunk PUT would carry the user's token to a URL that already carries one",
  );
  // The same URL is a credential, so the one response body handed straight back
  // to a caller strikes it out, and the logger never sees it at all.
  assert.match(source, /authtoken=/);
});

test("the read-flag write patches isRead on one message and nothing else", () => {
  // `patch` is the verb that could rewrite a message, so the flag write is
  // pinned to the one property it is allowed to send, on a path built from the
  // id the caller named.
  const source = readFileSync(join(repo, "src/core/graph.ts"), "utf8");
  const start = source.indexOf("async function setMessageRead");
  const flagWrite = source.slice(
    start,
    start === -1
      ? undefined
      : source.indexOf("return { updated, failed }", start),
  );
  assert.ok(
    flagWrite.length > 0,
    "setMessageRead is the flag write this test reads",
  );
  assert.match(flagWrite, /\.api\(messagePath\(id\)\)/);
  assert.match(flagWrite, /\.patch[\s\S]{0,80}isRead: read/);
  for (const content of [
    "draftRequestBody",
    "subject",
    "Recipients",
    ".post(",
    ".delete(",
  ]) {
    assert.equal(
      flagWrite.includes(content),
      false,
      `the flag write touches ${content}, so it writes more than a flag`,
    );
  }
});

test("readonly mode advertises no mailbox write at all", async (t) => {
  const config = testConfig(t, { OUTLOOK_MCP_READONLY: "1" });
  const { client, names } = await advertised(t, config);

  // Pinned against the read half of the registry: a draft tool or a flag tool
  // left in this mode cannot match it, so there is nothing to re-check name by
  // name. Both are named anyway, because the two groups can move independently.
  assert.deepEqual(names, [...readToolNames].sort());
  for (const write of [...draftToolNames, ...flagToolNames])
    assert.equal(
      names.includes(write),
      false,
      `${write} is advertised where the token is only Mail.Read`,
    );

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
    for (const write of [...draftToolNames, ...flagToolNames])
      assert.equal(
        names.includes(write),
        !readonly,
        `${write} with readonly=${readonly}`,
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

test("without readonly the eight draft tools are advertised with their limits", async (t) => {
  const config = testConfig(t);
  const { client, names } = await advertised(t, config);

  assert.deepEqual(
    names,
    [...readToolNames, ...flagToolNames, ...draftToolNames].sort(),
  );
  assert.match(
    client.getInstructions() ?? "",
    /create_draft, create_reply_draft, create_forward_draft, update_draft, list_drafts, delete_draft, add_attachment and remove_attachment/,
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
