import assert from "node:assert/strict";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import {
  Client,
  StreamableHTTPClientTransport,
} from "@modelcontextprotocol/client";
import { loadConfig } from "../src/core/config.ts";
import { attachmentResources } from "../src/core/attachments.ts";
import type { Graph } from "../src/core/graph.ts";
import type { ToolDeps } from "../src/tools.ts";
import { startHttp } from "../src/mcp.ts";

test("attachment size configuration defaults to 20 MiB and validates overrides", () => {
  assert.equal(loadConfig({}).maxAttachmentBytes, 20 * 1024 * 1024);
  assert.equal(
    loadConfig({ OUTLOOK_MCP_MAX_ATTACHMENT_BYTES: "42" }).maxAttachmentBytes,
    42,
  );
  for (const value of ["0", "-1", "1.5", "NaN", "Infinity", "9007199254740992"])
    assert.throws(
      () => loadConfig({ OUTLOOK_MCP_MAX_ATTACHMENT_BYTES: value }),
      /OUTLOOK_MCP_MAX_ATTACHMENT_BYTES/,
    );
});

function mailbox() {
  let body = "aGVsbG8=";
  let size: number | undefined = 5;
  let type = "#microsoft.graph.fileAttachment";
  let deleted = false;
  const calls: boolean[] = [];
  const graph = {
    async getAttachment(
      _message: string,
      _attachment: string,
      metadataOnly = false,
    ) {
      calls.push(metadataOnly);
      if (deleted) throw new Error("Attachment deleted");
      return {
        name: "../note.txt",
        contentType: "text/plain",
        size,
        "@odata.type": type,
        ...(metadataOnly ? {} : { contentBytes: body }),
      };
    },
  } as unknown as Graph;
  return {
    graph,
    calls,
    setBody(value: string) {
      body = value;
    },
    setSize(value: number | undefined) {
      size = value;
    },
    setType(value: string) {
      type = value;
    },
    remove() {
      deleted = true;
    },
  };
}

test("resources retain only handles, re-fetch content, and expire with their registry", async () => {
  const fake = mailbox();
  const registry = attachmentResources(fake.graph, loadConfig({}));
  const link = await registry.create("message", "attachment");
  assert.equal(link.name, "note.txt");
  assert.deepEqual(
    fake.calls,
    [true],
    "creating a link must not fetch content",
  );
  assert.equal((await registry.create("message", "attachment")).uri, link.uri);
  assert.equal((await registry.read(link.uri)).contents[0].blob, "aGVsbG8=");
  fake.setBody("d29ybGQ=");
  assert.equal(
    (await registry.read(link.uri)).contents[0].blob,
    "d29ybGQ=",
    "no stale content cache",
  );
  const other = attachmentResources(fake.graph, loadConfig({}));
  await assert.rejects(other.read(link.uri), /unknown or expired/);
  fake.remove();
  await assert.rejects(registry.read(link.uri), /deleted/);
  registry.clear();
  await assert.rejects(registry.read(link.uri), /unknown or expired/);
});

test("size limits are enforced before body fetch and against actual decoded content", async () => {
  const fake = mailbox();
  const registry = attachmentResources(
    fake.graph,
    loadConfig({ OUTLOOK_MCP_MAX_ATTACHMENT_BYTES: "5" }),
  );
  fake.setSize(6);
  await assert.rejects(registry.create("m", "a"), /maximum is 5/);
  assert.deepEqual(fake.calls, [true]);
  fake.setSize(5);
  const link = await registry.create("m", "a");
  assert.equal(
    (await registry.read(link.uri)).contents[0].blob,
    "aGVsbG8=",
    "exact limit is accepted",
  );
  fake.setSize(undefined);
  fake.setBody(Buffer.alloc(6).toString("base64"));
  await assert.rejects(registry.read(link.uri), /maximum is 5/);
  fake.setBody("bad content");
  await assert.rejects(registry.read(link.uri), /invalid attachment base64/);
  fake.setType("#microsoft.graph.itemAttachment");
  await assert.rejects(registry.create("m", "item"), /Only file attachments/);
});

test("the default limit accepts a 20 MiB decoded resource", async () => {
  const fake = mailbox();
  const size = 20 * 1024 * 1024;
  fake.setSize(size);
  fake.setBody(Buffer.alloc(size, 97).toString("base64"));
  const registry = attachmentResources(fake.graph, loadConfig({}));
  const link = await registry.create("m", "a");
  const resource = await registry.read(link.uri);
  assert.equal(
    Buffer.from(resource.contents[0].blob, "base64").byteLength,
    size,
  );
  fake.setSize(size + 1);
  await assert.rejects(registry.read(link.uri), /maximum is 20971520/);
});

test("HTTP clients read attachments without shared files and cannot reuse another session's handle", async (t) => {
  const home = mkdtempSync(join(tmpdir(), "outlook-resources-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const config = loadConfig({
    HOME: home,
    OUTLOOK_MCP_ATTACHMENT_DIR: join(home, "attachments"),
  });
  const fake = mailbox();
  const handle = await startHttp({ config, graph: fake.graph } as ToolDeps, {
    host: "127.0.0.1",
    port: 0,
  });
  t.after(() => handle.close());
  const open = async () => {
    const client = new Client({ name: "attachments-test", version: "0" });
    const transport = new StreamableHTTPClientTransport(new URL(handle.url));
    await client.connect(transport);
    t.after(() => client.close());
    return { client, transport };
  };
  const first = await open();
  const result = await first.client.callTool({
    name: "download_attachment",
    arguments: { messageId: "m", attachmentId: "a" },
  });
  const link = result.content.find((block) => block.type === "resource_link");
  assert.ok(link && link.type === "resource_link");
  assert.equal(link.size, 5);
  const resource = await first.client.readResource({ uri: link.uri });
  assert.ok("blob" in resource.contents[0]);
  assert.equal(resource.contents[0].blob, "aGVsbG8=");
  assert.equal(
    existsSync(config.attachmentDir),
    false,
    "resource delivery writes nothing",
  );
  const inline = await first.client.callTool({
    name: "download_attachment",
    arguments: { messageId: "m", attachmentId: "a", encoding: "base64" },
  });
  assert.equal(
    (inline.structuredContent as Record<string, unknown>).contentBase64,
    "aGVsbG8=",
  );
  assert.equal(
    existsSync(config.attachmentDir),
    false,
    "inline delivery writes nothing",
  );
  const second = await open();
  await assert.rejects(
    second.client.readResource({ uri: link.uri }),
    /unknown or expired/,
  );
  await first.transport.terminateSession();
  const replacement = await open();
  await assert.rejects(
    replacement.client.readResource({ uri: link.uri }),
    /unknown or expired/,
  );
  const renewed = await replacement.client.callTool({
    name: "download_attachment",
    arguments: { messageId: "m", attachmentId: "a" },
  });
  const renewedLink = renewed.content.find(
    (block) => block.type === "resource_link",
  );
  assert.ok(renewedLink && renewedLink.type === "resource_link");
  assert.notEqual(renewedLink.uri, link.uri);
  const renewedResource = await replacement.client.readResource({
    uri: renewedLink.uri,
  });
  assert.ok("blob" in renewedResource.contents[0]);
  assert.equal(renewedResource.contents[0].blob, "aGVsbG8=");
});
