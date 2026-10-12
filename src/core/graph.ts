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

const messageFields =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments";
const fullMessageFields = `${messageFields},toRecipients,ccRecipients,body`;
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
      .api(`/me/messages/${encodeURIComponent(id)}`)
      .header("Prefer", 'outlook.body-content-type="text"')
      .select(fullMessageFields.split(","))
      .get();
  }

  function getAttachment(messageId: string, attachmentId: string) {
    return client
      .api(
        `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
      )
      .get();
  }

  async function listAttachments(messageId: string) {
    const path = `/me/messages/${encodeURIComponent(messageId)}/attachments`;
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

  return {
    listMessages,
    searchMessages,
    getMessage,
    listAttachments,
    getAttachment,
  };
}

export type Graph = ReturnType<typeof createGraph>;
