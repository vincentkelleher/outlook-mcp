import { Client } from "@microsoft/microsoft-graph-client";

export type TokenProvider = () => Promise<string>;

/** The subset of the Graph SDK request builder this project calls. */
export type GraphRequest = {
  header(name: string, value: string): GraphRequest;
  select(fields: string[]): GraphRequest;
  orderby(value: string): GraphRequest;
  top(value: number): GraphRequest;
  filter(value: string): GraphRequest;
  search(value: string): GraphRequest;
  get<T = unknown>(): Promise<T>;
  post<T = unknown>(body: unknown): Promise<T>;
  patch<T = unknown>(body: unknown): Promise<T>;
  delete(): Promise<void>;
};

export type GraphClient = {
  api(path: string): GraphRequest;
};

export type GraphClientFactory = (token: TokenProvider) => GraphClient;

export type GraphDeps = {
  token: TokenProvider;
  client?: GraphClientFactory;
};

export type Page<T> = {
  items: T[];
  nextCursor?: string;
};

type GraphCollection = {
  value: { id: string }[];
  "@odata.nextLink"?: string;
};

/** What a draft carries. Absent keys are left out of the request, not cleared. */
export type DraftContent = {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  bodyFormat?: "text" | "html";
  categories?: string[];
};

export type AttachmentUpload = {
  name: string;
  contentType: string;
  contentBase64: string;
};

/**
 * Graph accepts an attachment in one request only while it stays under 3 MB;
 * 3-150 MB needs an upload session, which this client does not implement, so
 * the ceiling is enforced before anything is sent.
 */
export const maxAttachmentBytes = 3 * 1024 * 1024;

export type Folder = {
  id: string;
  displayName: string;
  totalItemCount?: number;
  unreadItemCount?: number;
};

const messageFields =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments";
const fullMessageFields = `${messageFields},toRecipients,ccRecipients,body`;
const draftFields = `${messageFields},isDraft,webLink,toRecipients`;
const folderFields = "id,displayName,totalItemCount,unreadItemCount";
const folderPattern = /^[A-Za-z0-9][A-Za-z0-9 ._@-]*$/;

function sdkClient(token: TokenProvider): GraphClient {
  return Client.init({
    authProvider: async (done) => {
      try {
        done(null, await token());
      } catch (error) {
        done(error instanceof Error ? error : new Error(String(error)), null);
      }
    },
  }) as unknown as GraphClient;
}

function messagePath(id: string) {
  return `/me/messages/${encodeURIComponent(id)}`;
}

/** Graph takes a recipient as a nested address object, never a bare string. */
function recipientsOf(addresses: string[] | undefined) {
  return addresses?.map((address) => ({ emailAddress: { address } }));
}

/**
 * Keys with no value are dropped rather than sent as null, so a patch changes
 * only what it names and a readback claims only what Graph returned.
 */
export function withoutUndefined<T extends object>(fields: T): Partial<T> {
  return Object.fromEntries(
    Object.entries(fields).filter(([, value]) => value !== undefined),
  ) as Partial<T>;
}

function draftRequestBody(content: DraftContent) {
  return withoutUndefined({
    subject: content.subject,
    body:
      content.body === undefined
        ? undefined
        : {
            contentType: content.bodyFormat ?? "text",
            content: content.body,
          },
    toRecipients: recipientsOf(content.to),
    ccRecipients: recipientsOf(content.cc),
    bccRecipients: recipientsOf(content.bcc),
    categories: content.categories,
  });
}

export function messagesPath(folder = "inbox") {
  const name = folder.trim();
  if (name === "all") return "/me/messages";
  if (!folderPattern.test(name))
    throw new Error(`Invalid Outlook folder name: ${folder}`);
  return `/me/mailFolders/${encodeURIComponent(name)}/messages`;
}

export function createGraph(deps: GraphDeps) {
  const client = deps.client ? deps.client(deps.token) : sdkClient(deps.token);

  async function page<T>(
    target: string,
    build: (request: GraphRequest) => GraphRequest = (request) => request,
  ): Promise<Page<T>> {
    const response = (await build(client.api(target)).get()) as {
      value: T[];
      "@odata.nextLink"?: string;
    };
    return { items: response.value, nextCursor: response["@odata.nextLink"] };
  }

  async function listMessages(
    options: {
      folder?: string;
      unread?: boolean;
      top?: number;
      cursor?: string;
    } = {},
  ) {
    const { folder = "inbox", unread = false, top = 10, cursor } = options;
    // A cursor is a Graph nextLink, which already carries every query option.
    if (cursor) return page(cursor);
    return page(messagesPath(folder), (request) => {
      const ordered = request
        .select(messageFields.split(","))
        .orderby("receivedDateTime desc")
        .top(top);
      return unread ? ordered.filter("isRead eq false") : ordered;
    });
  }

  async function searchMessages(options: {
    query: string;
    top?: number;
    cursor?: string;
  }) {
    const { query, top = 10, cursor } = options;
    if (cursor) return page(cursor);
    return page("/me/messages", (request) =>
      request
        .header("ConsistencyLevel", "eventual")
        .search(`"${query.replace(/[\\"]/g, "\\$&")}"`)
        .top(top)
        .select(messageFields.split(",")),
    );
  }

  function getMessage(id: string) {
    return client
      .api(messagePath(id))
      .header("Prefer", 'outlook.body-content-type="text"')
      .select(fullMessageFields.split(","))
      .get();
  }

  function getAttachment(messageId: string, attachmentId: string) {
    return client
      .api(
        `${messagePath(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
      )
      .get();
  }

  async function listFolders() {
    return page<Folder>("/me/mailFolders", (request) =>
      request.select(folderFields.split(",")).top(100),
    );
  }

  async function listAttachments(messageId: string) {
    const path = `${messagePath(messageId)}/attachments`;
    const collected: { id: string }[] = [];
    let target: string | undefined = path;
    while (target) {
      const collection: GraphCollection = await client
        .api(target)
        .get<GraphCollection>();
      collected.push(...collection.value);
      target = collection["@odata.nextLink"];
    }
    return Promise.all(collected.map(({ id }) => getAttachment(messageId, id)));
  }

  /**
   * POST /me/messages creates a draft. Graph only transmits on its `send`
   * action, and nothing in this file calls it, or the `reply` or `forward`
   * actions - see test/no-send.test.ts.
   */
  function createDraft(content: DraftContent) {
    return client.api("/me/messages").post(draftRequestBody(content));
  }

  /** The create actions keep the thread and copy in the answered recipients. */
  function createReplyDraft(messageId: string, comment: string) {
    return client
      .api(`${messagePath(messageId)}/createReply`)
      .post({ comment });
  }

  function createForwardDraft(
    messageId: string,
    to: string[],
    comment: string,
  ) {
    return client
      .api(`${messagePath(messageId)}/createForward`)
      .post({ comment, toRecipients: recipientsOf(to) });
  }

  function updateDraft(messageId: string, patch: DraftContent) {
    return client.api(messagePath(messageId)).patch(draftRequestBody(patch));
  }

  async function listDrafts(options: { top?: number; cursor?: string } = {}) {
    const { top = 10, cursor } = options;
    if (cursor) return page(cursor);
    return page(messagesPath("drafts"), (request) =>
      request
        .select(draftFields.split(","))
        .orderby("receivedDateTime desc")
        .top(top),
    );
  }

  function deleteDraft(messageId: string) {
    return client.api(messagePath(messageId)).delete();
  }

  /** The decoded size is checked here, so an over-limit file costs no request. */
  function addAttachment(messageId: string, attachment: AttachmentUpload) {
    const bytes = Buffer.from(attachment.contentBase64, "base64").byteLength;
    if (bytes >= maxAttachmentBytes)
      throw new Error(
        `${attachment.name} is ${bytes} bytes and a single-request attachment must stay under ${maxAttachmentBytes} bytes (3 MB). Graph needs an upload session from 3 to 150 MB, which this server does not implement, so the draft was not changed.`,
      );
    return client.api(`${messagePath(messageId)}/attachments`).post({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: attachment.name,
      contentType: attachment.contentType,
      contentBase64: attachment.contentBase64,
    });
  }

  return {
    listMessages,
    searchMessages,
    getMessage,
    listFolders,
    listAttachments,
    getAttachment,
    createDraft,
    createReplyDraft,
    createForwardDraft,
    updateDraft,
    listDrafts,
    deleteDraft,
    addAttachment,
  };
}

export type Graph = ReturnType<typeof createGraph>;
