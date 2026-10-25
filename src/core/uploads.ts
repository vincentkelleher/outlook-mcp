import {
  attachmentContent,
  attachmentMetadata,
  safeFileName,
} from "./attachments.ts";
import type { Config } from "./config.ts";
import { openAttachmentFile } from "./files.ts";
import type { Graph, UploadSource } from "./graph.ts";

/**
 * The three places a draft's file can come from, and the one shape every upload
 * path takes them in. Nothing here touches Graph's wire: this module decides
 * *whose bytes* are being attached and how big they are, and `uploadAttachment`
 * decides *how* they travel.
 *
 * The distinction is the point. `contentBase64` is the only source that puts a
 * file in front of a model, and the two that do not are the reason this exists:
 * a 20 MB file is 27 MB of base64 nobody should have to emit token by token.
 */
export type AttachmentSources = {
  /** Already in the caller's hands, small, and in the transcript. */
  contentBase64?: string;
  /** A file inside `OUTLOOK_MCP_ATTACHMENT_DIR`, read from disk as it goes up. */
  serverPath?: string;
  /** A received attachment, copied without leaving the server. */
  fromAttachment?: { messageId: string; attachmentId: string };
};

export type AttachmentRequest = AttachmentSources & {
  /** What the recipient sees. Defaults from the source, so only base64 needs it. */
  name?: string;
  contentType?: string;
};

/**
 * The one source a request named, as a union the rest of the module can switch
 * on. The tool's schema already refuses none-or-two; this is the same rule for
 * anyone calling `sourceOf` directly, and it is what lets the branches below
 * read `request.contentBase64` without asserting that it is there.
 */
type Source =
  | { kind: "base64"; contentBase64: string }
  | { kind: "file"; serverPath: string }
  | { kind: "copy"; messageId: string; attachmentId: string };

/** What each source is called in the sentence a caller is refused with. */
const spellings = {
  base64: "contentBase64",
  file: "serverPath",
  copy: "fromAttachment",
} as const;

function oneSource(request: AttachmentRequest): Source {
  const { contentBase64, serverPath, fromAttachment } = request;
  const picked = [
    contentBase64 !== undefined && { kind: "base64", contentBase64 },
    serverPath !== undefined && { kind: "file", serverPath },
    fromAttachment !== undefined && {
      kind: "copy",
      messageId: fromAttachment.messageId,
      attachmentId: fromAttachment.attachmentId,
    },
  ].filter((source): source is Source => source !== false);
  if (picked.length !== 1)
    throw new Error(
      picked.length === 0
        ? "add_attachment needs one of contentBase64, serverPath or fromAttachment: bytes to attach, a file inside the attachment directory, or a received attachment to copy."
        : `add_attachment takes one source at a time, not ${picked
            .map((source) => spellings[source.kind])
            .join(" and ")}.`,
    );
  return picked[0];
}

/** A source that keeps something open gives it back here, however the upload ends. */
type WithClose = UploadSource & { close?: () => void };

/**
 * Resolve one attachment request into bytes Graph can send.
 *
 * The size is settled here — from the decoded base64, from the file's own
 * `stat`, or from Graph's metadata for a copy — so the ceiling is enforced
 * before any write is attempted, and the name and type can fall back to what
 * the source already says about itself. A name taken from a received mail is
 * sender data and goes through `safeFileName`; an explicit one is the caller's
 * and is used as given.
 */
export async function sourceOf(
  graph: Graph,
  config: Config,
  request: AttachmentRequest,
): Promise<WithClose> {
  const source = oneSource(request);
  const explicitName = request.name;

  if (source.kind === "base64") {
    const bytes = Buffer.from(source.contentBase64, "base64");
    return {
      // The schema requires a name here, base64 carrying none of its own.
      name: explicitName ?? "attachment",
      contentType: request.contentType ?? "application/octet-stream",
      size: bytes.byteLength,
      read: async (offset, length) => bytes.subarray(offset, offset + length),
    };
  }

  if (source.kind === "file") {
    // Containment is `openAttachmentFile`'s to enforce, and it enforces it by
    // throwing: nothing below this line has a path that is not inside the
    // attachment directory.
    const file = openAttachmentFile(config, source.serverPath);
    return {
      name: explicitName ?? safeFileName(file.name),
      contentType: request.contentType ?? "application/octet-stream",
      size: file.size,
      read: async (offset, length) => file.read(offset, length),
      close: file.close,
    };
  }

  const { messageId, attachmentId } = source;
  // A copy crosses both limits, because it is a download and an upload in one
  // request that never leaves the process. The metadata is read first so the
  // refusal says what a copy needs — which setting to move — instead of
  // borrowing the wording of a download the caller never asked for.
  const limit = Math.min(config.maxAttachmentBytes, config.maxUploadBytes);
  const metadata = await attachmentMetadata(
    graph,
    messageId,
    attachmentId,
    Number.MAX_SAFE_INTEGER,
  );
  if (metadata.size !== undefined && metadata.size > limit)
    throw new Error(
      `${metadata.name ?? attachmentId} is ${metadata.size} bytes and a copy of a received attachment is limited to ${limit} bytes — the smaller of the download limit and the upload limit, because it is both in one call. Raise OUTLOOK_MCP_MAX_ATTACHMENT_BYTES and OUTLOOK_MCP_MAX_UPLOAD_BYTES, or download it and attach it with serverPath. Nothing was attached.`,
    );
  const { attachment, bytes } = await attachmentContent(
    graph,
    messageId,
    attachmentId,
    limit,
    metadata,
  );
  return {
    name: explicitName ?? safeFileName(attachment.name),
    contentType:
      request.contentType ??
      attachment.contentType ??
      "application/octet-stream",
    size: bytes.byteLength,
    read: async (offset, length) => bytes.subarray(offset, offset + length),
  };
}
