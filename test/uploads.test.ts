import assert from "node:assert/strict";
import {
  mkdirSync,
  mkdtempSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { loadConfig } from "../src/core/config.ts";
import type { Config } from "../src/core/config.ts";
import { maxGraphUploadBytes } from "../src/core/graph.ts";
import type { Graph } from "../src/core/graph.ts";
import { sourceOf } from "../src/core/uploads.ts";

function testConfig(
  t: { after: (fn: () => void) => void },
  env: Record<string, string> = {},
): Config {
  const home = mkdtempSync(join(tmpdir(), "outlook-uploads-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  return loadConfig({
    HOME: home,
    OUTLOOK_OAUTH_CLIENT_ID: "00000000-0000-0000-0000-000000000000",
    OUTLOOK_MCP_ATTACHMENT_DIR: join(home, "attachments"),
    OUTLOOK_MCP_STATE_DIR: join(home, "state"),
    ...env,
  });
}

test("the upload ceiling defaults to the mailbox's, and is set within Graph's", () => {
  assert.equal(loadConfig({}).maxUploadBytes, 35 * 1024 * 1024);
  assert.equal(
    loadConfig({ OUTLOOK_MCP_MAX_UPLOAD_BYTES: "4096" }).maxUploadBytes,
    4096,
  );
  // A number above what Graph takes is a mistake about Graph, not a wider
  // allowance, so it fails here rather than at the fortieth megabyte.
  assert.throws(
    () =>
      loadConfig({
        OUTLOOK_MCP_MAX_UPLOAD_BYTES: String(maxGraphUploadBytes + 1),
      }),
    /OUTLOOK_MCP_MAX_UPLOAD_BYTES/,
  );
  for (const value of ["0", "-1", "1.5", "NaN", "ten"])
    assert.throws(
      () => loadConfig({ OUTLOOK_MCP_MAX_UPLOAD_BYTES: value }),
      /OUTLOOK_MCP_MAX_UPLOAD_BYTES/,
    );
  // The download limit is a separate knob: raising one does not raise the other.
  const raised = loadConfig({ OUTLOOK_MCP_MAX_ATTACHMENT_BYTES: "1024" });
  assert.equal(raised.maxAttachmentBytes, 1024);
  assert.equal(raised.maxUploadBytes, 35 * 1024 * 1024);
});

/** One received attachment, with the calls it took named in `seen`. */
function received(
  options: { size?: number; name?: string; contentType?: string } = {},
) {
  const seen: boolean[] = [];
  const graph = {
    async getAttachment(
      _messageId: string,
      _attachmentId: string,
      metadataOnly = false,
    ) {
      seen.push(metadataOnly);
      return {
        name: options.name ?? "invoice.pdf",
        contentType: options.contentType ?? "application/pdf",
        size: options.size ?? 5,
        "@odata.type": "#microsoft.graph.fileAttachment",
        ...(metadataOnly
          ? {}
          : { contentBytes: Buffer.from("hello").toString("base64") }),
      };
    },
  } as unknown as Graph;
  return { graph, seen };
}

test("base64 content becomes a source that reads itself back", async (t) => {
  const config = testConfig(t);
  const { graph } = received();
  const source = await sourceOf(graph, config, {
    name: "note.txt",
    contentType: "text/plain",
    contentBase64: Buffer.from("hello").toString("base64"),
  });

  assert.equal(source.name, "note.txt");
  assert.equal(source.contentType, "text/plain");
  assert.equal(source.size, 5);
  assert.equal((await source.read(0, 5)).toString(), "hello");
  assert.equal(
    (await source.read(2, 2)).toString(),
    "ll",
    "reads are positional",
  );
  assert.equal(
    source.close,
    undefined,
    "nothing was opened, so nothing to close",
  );
});

test("serverPath reads a file inside the attachment directory, by any route that stays inside", async (t) => {
  const config = testConfig(t);
  const { graph } = received();
  mkdirSync(join(config.attachmentDir, "downloads"), { recursive: true });
  writeFileSync(
    join(config.attachmentDir, "downloads", "quote.pdf"),
    "the quote",
  );

  const source = await sourceOf(graph, config, {
    serverPath: "downloads/quote.pdf",
  });
  assert.equal(
    source.name,
    "quote.pdf",
    "the leaf is the name the recipient sees",
  );
  assert.equal(source.contentType, "application/octet-stream");
  assert.equal(source.size, "the quote".length);
  assert.equal((await source.read(0, 9)).toString(), "the quote");
  // An explicit name and type win over what the file says about itself.
  const named = await sourceOf(graph, config, {
    serverPath: "downloads/quote.pdf",
    name: "Q3.pdf",
    contentType: "application/pdf",
  });
  assert.equal(named.name, "Q3.pdf");
  assert.equal(named.contentType, "application/pdf");
  source.close?.();
  named.close?.();
});

test("serverPath refuses every path that is not a plain file inside the directory", async (t) => {
  const config = testConfig(t);
  const { graph } = received();
  mkdirSync(config.attachmentDir, { recursive: true });
  writeFileSync(join(config.attachmentDir, "plain.txt"), "plain");
  symlinkSync(
    join(config.attachmentDir, "plain.txt"),
    join(config.attachmentDir, "link.txt"),
  );
  symlinkSync("/etc/passwd", join(config.attachmentDir, "escape.txt"));
  mkdirSync(join(config.attachmentDir, "adir"), { recursive: true });

  const refused: [string, RegExp][] = [
    ["../outside.txt", /must be inside the attachment directory/],
    ["/etc/passwd", /must be inside the attachment directory/],
    // Normalising `..` is `resolve`'s job and it happens before the check, so a
    // path that climbs out is refused whatever it looks like on the way.
    ["sub/../../etc/passwd", /must be inside the attachment directory/],
    ["link.txt", /is a symlink/],
    ["escape.txt", /is a symlink/],
    ["adir", /regular file/],
    ["missing.txt", /no file at/],
  ];
  for (const [serverPath, reason] of refused) {
    await assert.rejects(
      () => sourceOf(graph, config, { serverPath }),
      reason,
      `${serverPath} was read`,
    );
  }
});

test("a copy takes its name and type from the attachment, and reads its size before its content", async (t) => {
  const config = testConfig(t);
  const mailbox = received({ size: 5 });

  const source = await sourceOf(mailbox.graph, config, {
    fromAttachment: { messageId: "mail-1", attachmentId: "att-1" },
  });

  assert.equal(source.name, "invoice.pdf");
  assert.equal(source.contentType, "application/pdf");
  assert.equal(source.size, 5);
  assert.equal((await source.read(0, 5)).toString(), "hello");
  // Metadata first so a copy that cannot fit costs the small request; content
  // once, for the bytes themselves.
  assert.deepEqual(mailbox.seen, [true, false]);
});

test("a copy too big for the smaller of the two limits is refused before its content is fetched", async (t) => {
  const config = { ...testConfig(t), maxAttachmentBytes: 1024 * 1024 };
  const mailbox = received({ size: 4 * 1024 * 1024, name: "deck.pptx" });

  await assert.rejects(
    () =>
      sourceOf(mailbox.graph, config, {
        fromAttachment: { messageId: "mail-1", attachmentId: "att-1" },
      }),
    (error: unknown) => {
      const message = error instanceof Error ? error.message : String(error);
      assert.match(message, /deck\.pptx is 4194304 bytes/);
      assert.match(message, /limited to 1048576 bytes/);
      assert.match(message, /OUTLOOK_MCP_MAX_ATTACHMENT_BYTES/);
      assert.match(message, /OUTLOOK_MCP_MAX_UPLOAD_BYTES/);
      assert.deepEqual(
        mailbox.seen,
        [true],
        "the file was fetched to refuse it",
      );
      return true;
    },
  );
});

test("one source, and only one: the choice is settled before anything is read", async (t) => {
  const config = testConfig(t);
  const mailbox = received();

  await assert.rejects(
    () => sourceOf(mailbox.graph, config, {}),
    /needs one of contentBase64, serverPath or fromAttachment/,
  );
  await assert.rejects(
    () =>
      sourceOf(mailbox.graph, config, {
        contentBase64: "aGVsbG8=",
        serverPath: "downloads/quote.pdf",
      }),
    /takes one source at a time, not contentBase64 and serverPath/,
  );
  assert.deepEqual(mailbox.seen, []);
});
