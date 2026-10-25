import { randomUUID } from "node:crypto";
import type { Config } from "./config.ts";
import type { Graph } from "./graph.ts";

export type GraphAttachment = {
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentBytes?: string;
  "@odata.type"?: string;
};

export function safeFileName(value: string | undefined): string {
  const leaf = (value ?? "").split(/[\\/]/).pop() ?? "";
  return (
    leaf
      .replace(/[\u0000-\u001f]/g, "_")
      .replace(/^\.+/, "")
      .trim() || "attachment"
  );
}

export function checkAttachmentSize(
  size: number | undefined,
  maximum: number,
): void {
  if (size !== undefined && (!Number.isSafeInteger(size) || size < 0))
    throw new Error("Graph returned an invalid attachment size.");
  if (size !== undefined && size > maximum)
    throw new Error(
      `Attachment is ${size} bytes; maximum is ${maximum} bytes. Use resource delivery for oversized inline requests, or increase OUTLOOK_MCP_MAX_ATTACHMENT_BYTES if the configured download limit is too low.`,
    );
}

export async function attachmentMetadata(
  graph: Graph,
  messageId: string,
  attachmentId: string,
  maximum: number,
): Promise<GraphAttachment> {
  const attachment = (await graph.getAttachment(
    messageId,
    attachmentId,
    true,
  )) as GraphAttachment;
  if (
    attachment["@odata.type"] &&
    attachment["@odata.type"] !== "#microsoft.graph.fileAttachment"
  )
    throw new Error("Only file attachments can be read as binary resources.");
  checkAttachmentSize(attachment.size, maximum);
  return attachment;
}

/**
 * Preflight metadata before fetching a body, then enforce the actual decoded
 * size. `prefetched` is for a caller that read the metadata a moment ago to
 * answer something else about the attachment — a copy, which needs its name and
 * type as well as its size — and should not have to ask twice.
 */
export async function attachmentContent(
  graph: Graph,
  messageId: string,
  attachmentId: string,
  maximum: number,
  prefetched?: GraphAttachment,
) {
  if (prefetched === undefined)
    await attachmentMetadata(graph, messageId, attachmentId, maximum);
  const attachment = (await graph.getAttachment(
    messageId,
    attachmentId,
  )) as GraphAttachment;
  checkAttachmentSize(attachment.size, maximum);
  const content = attachment.contentBytes;
  if (typeof content !== "string")
    throw new Error(
      "Attachment has no file content; it may have been deleted or is not a file attachment.",
    );
  if (
    content.length % 4 !== 0 ||
    /[^A-Za-z0-9+/]/.test(content.replace(/={1,2}$/, ""))
  )
    throw new Error("Graph returned invalid attachment base64 content.");
  const size =
    (content.length / 4) * 3 -
    (content.endsWith("==") ? 2 : content.endsWith("=") ? 1 : 0);
  checkAttachmentSize(size, maximum);
  return { attachment, bytes: Buffer.from(content, "base64") };
}

/** One registry per MCP connection. Only identifiers are retained, never content. */
export function attachmentResources(graph: Graph, config: Config) {
  const handles = new Map<
    string,
    { messageId: string; attachmentId: string }
  >();
  return {
    async create(messageId: string, attachmentId: string) {
      const metadata = await attachmentMetadata(
        graph,
        messageId,
        attachmentId,
        config.maxAttachmentBytes,
      );
      let uri = [...handles].find(
        ([, value]) =>
          value.messageId === messageId && value.attachmentId === attachmentId,
      )?.[0];
      if (!uri) {
        if (handles.size >= 1024)
          throw new Error(
            "Attachment resource limit reached for this session; start a new session.",
          );
        uri = `outlook://attachments/${randomUUID()}`;
        handles.set(uri, { messageId, attachmentId });
      }
      return {
        uri,
        name: safeFileName(metadata.name),
        mimeType: metadata.contentType ?? "application/octet-stream",
        ...(metadata.size === undefined ? {} : { size: metadata.size }),
      };
    },
    async read(uri: string) {
      const handle = handles.get(uri);
      if (!handle)
        throw new Error(
          "Attachment resource is unknown or expired. Call download_attachment in this session to obtain a new resource link.",
        );
      const { attachment, bytes } = await attachmentContent(
        graph,
        handle.messageId,
        handle.attachmentId,
        config.maxAttachmentBytes,
      );
      return {
        contents: [
          {
            uri,
            mimeType: attachment.contentType ?? "application/octet-stream",
            blob: bytes.toString("base64"),
          },
        ],
      };
    },
    clear() {
      handles.clear();
    },
  };
}
