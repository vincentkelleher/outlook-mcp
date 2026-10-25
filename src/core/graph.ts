import { Client } from "@microsoft/microsoft-graph-client";
import { AuthRequiredError, noSessionMessage } from "./errors.ts";
import { errorFields, logger, stopwatch } from "./log.ts";

const log = logger("graph");

type TokenProvider = () => Promise<string>;

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

type GraphClientFactory = (token: TokenProvider) => GraphClient;

type GraphDeps = {
  token: TokenProvider;
  client?: GraphClientFactory;
  /** Ceiling for one attachment, whatever Graph would take. */
  maxUploadBytes?: number;
  /** Only the upload session chunk `PUT` uses it; everything else is the SDK. */
  fetch?: FetchLike;
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
type DraftContent = {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  bodyFormat?: "text" | "html";
  categories?: string[];
};

/**
 * One file, from wherever it lives, as the only shape the upload paths need.
 * `read` is positional rather than a stream: an upload session asks for a range,
 * retries a range, or resumes at a range the server names, and a stream can do
 * none of that. Sizes are what the caller says the file is, checked against the
 * configured ceiling before a single request is made.
 */
export type UploadSource = {
  name: string;
  contentType: string;
  size: number;
  read(offset: number, length: number): Promise<Buffer>;
  /** A source holding something open — a file descriptor — releases it here. */
  close?(): void;
};

/** What an upload ended as: the attachment Graph made, as far as it says. */
export type UploadedAttachment = {
  id?: string;
  name: string;
  contentType: string;
  bytes: number;
  /** Which of the two ways the file travelled. Worth knowing when a draft is odd. */
  transport: "single_request" | "upload_session";
};

/**
 * The draft an attachment write reads back on its way past the draft guard. The
 * fields an answer needs to carry are spelled here so the write that already has
 * them can hand them over, instead of reading the draft a second time.
 */
export type GraphDraftRecord = {
  id?: string;
  subject?: string;
  body?: { content?: string; contentType?: string };
  toRecipients?: { emailAddress?: { address?: string } }[];
  ccRecipients?: { emailAddress?: { address?: string } }[];
  bccRecipients?: { emailAddress?: { address?: string } }[];
  categories?: string[];
  webLink?: string;
  isDraft?: boolean;
};

/**
 * The chunk `PUT` cannot go through the SDK: an upload session URL is
 * pre-authenticated with a bearer token in its query string, and the Graph
 * documentation is explicit that no `Authorization` header may be sent with it,
 * which the builder would add to every request. So the raw exchange is behind
 * this one seam, injectable so no test reaches Outlook.
 */
type PutResponse = {
  status: number;
  header(name: string): string | null;
  body(): Promise<string>;
};

export type FetchLike = (
  url: string,
  init: { method: "PUT"; headers: Record<string, string>; body: Buffer },
) => Promise<PutResponse>;

const putOverFetch: FetchLike = async (url, init) => {
  // undici's `BodyInit` will not admit a `Buffer` whose backing store is generic,
  // though it accepts one at runtime and this is the only place the bytes leave.
  // A copy into a plain `Uint8Array` would duplicate a chunk of every upload, so
  // the value goes over as it stands.
  const response = await fetch(url, init as RequestInit);
  return {
    status: response.status,
    header: (name) => response.headers.get(name),
    body: () => response.text(),
  };
};

/**
 * The page sizes, in one place because an agent is told them in prose: a
 * default, the most any single page asks for - emails in a list, drafts, folders
 * - and the larger ceiling a search may ask for, since a query narrows on its
 * own. Every schema, error message and description quotes these, so a limit
 * cannot disagree with the text a caller was given.
 */
export const defaultPageSize = 10;
export const maxPageSize = 100;
export const maxSearchPageSize = 1000;

/**
 * How many messages one `mark_read` / `mark_unread` call may name. Graph flags
 * a message one at a time, so this is a ceiling on the requests a single tool
 * call is allowed to make, not a limit Graph imposes — which is why it is the
 * same size as a page of mail and not the larger one a search gets away with.
 */
export const maxFlaggedMessages = 100;

/** A Graph error says the same thing in four lines; a result needs the first. */
export const maxFlagReasonChars = 300;

/** What one flag call ended up changing, id by id. */
export type MessageFlags = {
  /** What Graph says each changed message now carries. */
  updated: { id: string; isRead: boolean }[];
  /** Ids Graph refused, each with its own reason: one refusal is not all of them. */
  failed: { id: string; reason: string }[];
};

/**
 * Graph accepts an attachment in one request only while it stays under 3 MB;
 * anything larger goes up an upload session, which is what `uploadAttachment`
 * drives. The number is therefore a routing threshold rather than a limit: it
 * decides which of the two ways a file travels, not whether it may travel.
 */
export const singleRequestAttachmentBytes = 3 * 1024 * 1024;

/**
 * The largest attachment Graph itself takes, by either route, and the default
 * ceiling this server applies. Graph will accept 150 MB and a mailbox will not
 * necessarily send it: Outlook.com refuses around 20 MB and Microsoft 365 around
 * 35 MB, at Send, after everything above has already succeeded. The default is
 * the mailbox's number rather than Graph's so the failure arrives here, before
 * the bytes, with the setting that moves it named.
 */
export const maxGraphUploadBytes = 150 * 1024 * 1024;
export const defaultMaxUploadBytes = 35 * 1024 * 1024;

/**
 * Bytes per chunk of an upload session. Graph asks for ranges under 4 MB and
 * aligns every upload session on 320 KiB, so this is 320 KiB twelve times over:
 * under the ask, on the alignment, and few enough requests that a 35 MB file is
 * ten of them.
 */
export const uploadChunkBytes = 320 * 1024 * 12;

type Folder = {
  id: string;
  displayName: string;
  totalItemCount?: number;
  unreadItemCount?: number;
};

const messageFields =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments";
const fullMessageFields = `${messageFields},toRecipients,ccRecipients,body`;
const draftFields = `${messageFields},isDraft,webLink,toRecipients`;
/** What an attachment write reads back: the draft, its body, and its categories. */
const draftWriteFields = `${draftFields},body,categories`;
const folderFields = "id,displayName,totalItemCount,unreadItemCount";
const folderPattern = /^[A-Za-z0-9][A-Za-z0-9 ._@-]*$/;

/** The Graph host this client talks to. A cursor may name no other. */
const graphHost = "graph.microsoft.com";

/**
 * The page paths Graph issues a nextLink from, as the SDK versions them. A
 * cursor that points anywhere else was not issued by a page read here.
 */
const graphPagePath = /^\/v1\.0\/me\/(?:mailFolders\/[^/]+\/)?messages$/;

/**
 * A cursor comes back from the model, not from Graph, and an absolute URL given
 * to `client.api` is fetched as it stands - carrying the user's Microsoft bearer
 * token - to whatever host it names. So a cursor is accepted only as a URL Graph
 * itself could have issued: https, our host, no port, no credentials, a page
 * path under `/me`. A genuine nextLink passes through untouched.
 */
function assertGraphCursor(cursor: string): string {
  let link: URL | undefined;
  try {
    link = new URL(cursor);
  } catch {
    link = undefined;
  }
  const issued =
    link !== undefined &&
    link.protocol === "https:" &&
    link.hostname === graphHost &&
    link.port === "" &&
    link.username === "" &&
    link.password === "" &&
    graphPagePath.test(link.pathname);
  if (!issued)
    throw new Error(
      `cursor must be the nextCursor from an earlier result, an https://${graphHost} address issued by this server, not ${cursor}.`,
    );
  return cursor;
}

/**
 * The path without its query. The path is the useful half — it says which call
 * was made, and a message id in it is what a person debugging needs — while a
 * cursor repeats itself for as long as a caller pages, and is the one string in
 * the URL that carries a signed token. These lines are `debug`, which is where
 * the id belongs; the query belongs nowhere.
 */
function pathForLogging(path: string): string {
  const query = path.indexOf("?");
  const withoutQuery = query === -1 ? path : path.slice(0, query);
  if (query === -1 && !withoutQuery.includes("://")) return withoutQuery;
  // A cursor arrives as an absolute URL. The host is one this client refuses any
  // other than, so it adds nothing to a line; the path is the page.
  try {
    return `${new URL(withoutQuery).pathname}${query === -1 ? "" : "?…"}`;
  } catch {
    return withoutQuery;
  }
}

/** Graph's words for one refusal, on one line and short enough to return. */
function reasonOfGraphError(error: unknown): string {
  const message = (error instanceof Error ? error.message : String(error))
    .replace(/\s+/g, " ")
    .trim();
  return message.length > maxFlagReasonChars
    ? `${message.slice(0, maxFlagReasonChars)}…`
    : message;
}

/**
 * Hand back the sign-in failure as it left this project.
 *
 * The SDK rebuilds every rejection as its own `GraphError`: the message carries
 * over and `code` is filled from `error.name`, but the class is gone — and the
 * class is the one thing callers test. `error instanceof AuthRequiredError` is
 * what turns a failed call into "sign in, then retry" at the tool boundary, what
 * keeps it a `warn` rather than an `error` in the log, and what makes a flag
 * batch fail once instead of reporting the same refusal per id. Left wrapped, a
 * read of a mailbox with no session reads back as an ordinary server fault with
 * the right sentence in it, and the agent gets no route to `login`.
 */
function unwrapAuthError(error: unknown): unknown {
  if (error instanceof AuthRequiredError) return error;
  const wrapped = (error ?? {}) as { code?: unknown; message?: unknown };
  if (wrapped.code !== "AuthRequiredError") return error;
  return new AuthRequiredError(
    typeof wrapped.message === "string" && wrapped.message !== ""
      ? wrapped.message
      : noSessionMessage,
    { cause: error },
  );
}

type BuilderMethod = (...args: unknown[]) => unknown;

/**
 * One Graph call, timed and counted.
 *
 * Every request this project makes goes through the chained builder, so the
 * place to see all of them is a wrapper around that builder rather than forty
 * call sites that a future method forgets to add. It records at `debug`: a
 * successful page of mail is not news, and the failure that matters is already
 * reported once, at the boundary, with the reason. What the wrapper adds at
 * `debug` is which HTTP call that failure was.
 *
 * A builder method that returns the builder is part of the chain and keeps the
 * wrapper; anything else is the terminal call, whose promise is timed. The
 * rejection also passes through `unwrapAuthError`, because this is the one place
 * every request goes through and the SDK's rewrite is the one thing a caller
 * cannot see for itself.
 */
function traced(request: GraphRequest, path: string): GraphRequest {
  const wrapper: GraphRequest = new Proxy(request, {
    get(target, property, receiver) {
      const value = Reflect.get(target, property, receiver);
      if (typeof value !== "function") return value;
      const chainable = value as BuilderMethod;
      return (...args: unknown[]) => {
        const verb = String(property).toUpperCase();
        const outcome = chainable.apply(target, args);
        if (outcome === target) return wrapper;
        if (!(outcome instanceof Promise)) return outcome;
        const stop = stopwatch();
        return outcome.then(
          (resolved: unknown) => {
            log.debug("graph request", {
              method: verb,
              path: pathForLogging(path),
              ms: stop(),
              ...countsOf(resolved),
            });
            return resolved;
          },
          (error: unknown) => {
            log.debug("graph request failed", {
              method: verb,
              path: pathForLogging(path),
              ms: stop(),
              ...errorFields(error, log.detailed),
            });
            throw unwrapAuthError(error);
          },
        );
      };
    },
  });
  return wrapper;
}

/** How many items a response carried, when it is a collection. */
function countsOf(response: unknown) {
  if (Array.isArray(response)) return { items: response.length };
  const value = (response as { value?: unknown[] })?.value;
  return Array.isArray(value) ? { items: value.length } : {};
}

/**
 * Wrap the client so `api(path)` hands back a traced builder. The HTTP verb is
 * taken from the terminal method the chain ends in, which is the only verb this
 * project's own code can choose.
 */
function tracedClient(client: GraphClient): GraphClient {
  return {
    api(path: string) {
      return traced(client.api(path), path);
    },
  };
}

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
  const client = tracedClient(
    deps.client ? deps.client(deps.token) : sdkClient(deps.token),
  );
  const maxUploadBytes = Math.min(
    deps.maxUploadBytes ?? defaultMaxUploadBytes,
    maxGraphUploadBytes,
  );
  const put = deps.fetch ?? putOverFetch;

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
    const {
      folder = "inbox",
      unread = false,
      top = defaultPageSize,
      cursor,
    } = options;
    // A cursor is a Graph nextLink, which already carries every query option.
    if (cursor) return page(assertGraphCursor(cursor));
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
    const { query, top = defaultPageSize, cursor } = options;
    if (cursor) return page(assertGraphCursor(cursor));
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

  function getAttachment(
    messageId: string,
    attachmentId: string,
    metadataOnly = false,
  ) {
    const request = client.api(
      `${messagePath(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
    );
    if (metadataOnly)
      request.select(["id", "name", "contentType", "size", "isInline"]);
    return request.get();
  }

  async function listFolders() {
    return page<Folder>("/me/mailFolders", (request) =>
      request.select(folderFields.split(",")).top(maxPageSize),
    );
  }

  async function listAttachments(messageId: string, metadataOnly = false) {
    const path = `${messagePath(messageId)}/attachments`;
    const collected: { id: string }[] = [];
    let target: string | undefined = path;
    while (target) {
      const request = client.api(target);
      if (metadataOnly)
        request.select(["id", "name", "contentType", "size", "isInline"]);
      const collection = await request.get<GraphCollection>();
      collected.push(...collection.value);
      // Followed without assertGraphCursor, unlike a cursor that arrives from a
      // model: this link came back from Graph over TLS on a request we chose, so
      // validating it would guard nothing. Never copy this into a path a caller
      // can supply.
      target = collection["@odata.nextLink"];
    }
    if (metadataOnly) return collected;
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
    const { top = defaultPageSize, cursor } = options;
    if (cursor) return page(assertGraphCursor(cursor));
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

  /**
   * Flip the read flag on one or more messages with a PATCH of `isRead`, which
   * is the only way Graph does it: there is no collection endpoint, so a batch
   * is fanned out here and every id answers for itself. A message Graph will
   * not touch — already deleted, or never in this mailbox — is reported beside
   * the ones that moved rather than unwinding the call, because flagging is
   * per message and an agent that named ninety-nine good ids and one stale one
   * should not have the ninety-nine redone.
   *
   * A missing session is the exception. Every id fails for that one reason, and
   * the remediation is a sign-in rather than a page of identical refusals, so
   * it is raised as it came from the token provider.
   */
  async function setMessageRead(
    ids: string[],
    read: boolean,
  ): Promise<MessageFlags> {
    const named = [...new Set(ids)];
    const settled = await Promise.allSettled(
      named.map((id) => {
        return client
          .api(messagePath(id))
          .patch<{ id?: string; isRead?: boolean }>({
            isRead: read,
          });
      }),
    );
    const signedOut = settled.find(
      (outcome) =>
        outcome.status === "rejected" &&
        outcome.reason instanceof AuthRequiredError,
    );
    if (signedOut?.status === "rejected") throw signedOut.reason;
    const updated: MessageFlags["updated"] = [];
    const failed: MessageFlags["failed"] = [];
    settled.forEach((outcome, index) => {
      const id = named[index];
      if (outcome.status === "fulfilled") {
        // Read back what Graph stored. Where it said nothing, the flag it was
        // given is the honest answer.
        updated.push({
          id,
          isRead:
            typeof outcome.value?.isRead === "boolean"
              ? outcome.value.isRead
              : read,
        });
        return;
      }
      failed.push({ id, reason: reasonOfGraphError(outcome.reason) });
    });
    return { updated, failed };
  }

  /**
   * Read the message an attachment is about to be written to, refuse it if it is
   * not a draft, and hand back what was read for the answer to echo.
   *
   * Both attachment writes work on a message that has already been sent, which
   * is not what either tool here is for: one would add a file to mail that has
   * gone out, the other would take one away from it - and a message the user
   * already sent is the one thing in this mailbox nothing here may rewrite.
   * Graph answers `isDraft` for any message, so the flag is read first, and only
   * an explicit `false` stops the call: where a mailbox says nothing, refusing
   * would block a real draft on silence.
   *
   * The one read carries the draft's own fields, because the write that follows
   * it cannot change a subject, a body or a recipient: what is true a moment
   * before is honest to report as true afterwards.
   */
  async function draftToWrite(messageId: string): Promise<GraphDraftRecord> {
    const draft = await client
      .api(messagePath(messageId))
      .select(draftWriteFields.split(","))
      .get<GraphDraftRecord>();
    if (draft?.isDraft === false)
      throw new Error(
        `Message ${messageId} is not a draft, so nothing was changed. Attachments here are only for a message that has not been sent; a sent message keeps the files it went out with.`,
      );
    return draft ?? {};
  }

  type UploadSessionResponse = {
    uploadUrl?: string;
    expirationDateTime?: string;
    nextExpectedRanges?: unknown;
  };

  /** The first byte a session still wants, when it says. */
  function nextOffsetOf(ranges: unknown): number | undefined {
    const first = Array.isArray(ranges) ? ranges[0] : undefined;
    if (typeof first !== "string") return undefined;
    const offset = Number(first.split("-")[0]);
    return Number.isSafeInteger(offset) && offset >= 0 ? offset : undefined;
  }

  /** The session URL Graph gave, parsed and checked, or nothing. */
  function sessionUrlOf(value: string): URL | undefined {
    try {
      const url = new URL(value);
      return url.protocol === "https:" &&
        url.port === "" &&
        url.username === "" &&
        url.password === ""
        ? url
        : undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * The session URL is pre-authenticated - a bearer token sits in its query
   * string - and it is the only destination the file's bytes are sent to, so it
   * is read as a URL rather than trusted as a string: https, no credentials of
   * its own, no port of its own. Deliberately no exact-host allowlist: sovereign
   * clouds answer under names this project does not deploy to, and refusing them
   * is an outage nobody asked for. The URL arrives from Graph over TLS on a path
   * this file chose, which is the assurance `listAttachments` already takes for a
   * nextLink - and unlike a nextLink it is fetched without the user's token,
   * which is what bounds the damage one that was tampered with could do.
   */
  function assertSessionUrl(uploadUrl: string | undefined): string {
    if (typeof uploadUrl !== "string" || sessionUrlOf(uploadUrl) === undefined)
      throw new Error(
        "Graph returned an upload session this server will not send bytes to, so nothing was uploaded. Expected an https upload URL with no port and no credentials.",
      );
    return uploadUrl;
  }

  /** The attachment id Graph puts in the `Location` header of the final response. */
  function attachmentIdOf(location: string | null): string | undefined {
    const found = /Attachments\('([^']+)'\)/.exec(location ?? "");
    if (!found) return undefined;
    try {
      return decodeURIComponent(found[1]);
    } catch {
      return found[1];
    }
  }

  /**
   * An upload session URL, credential and all, can come back inside an error
   * body, and an error body is the one thing this project hands straight to the
   * caller. Strike it there too: the redaction in `log.ts` protects a log line,
   * not a message, and a token in a message reaches a model.
   */
  function withoutCredential(text: string): string {
    return text.replace(/authtoken=[^\s&"']*/gi, "authtoken=[redacted]");
  }

  /** How a chunk failed: Graph's own words, short, credential-free, never empty. */
  function reasonOfChunk(body: string): string {
    const text = withoutCredential(body ?? "");
    return reasonOfGraphError(
      new Error(text === "" ? "no reason given" : text),
    );
  }

  /** The ranges a session answer names, when it answers with parseable JSON. */
  function rangesOf(body: string): unknown {
    try {
      return (JSON.parse(body) as UploadSessionResponse).nextExpectedRanges;
    } catch {
      return undefined;
    }
  }

  /**
   * Send one file up an upload session: ranges, in order, until the last byte
   * lands. Graph materialises the attachment only at that last byte, so an
   * interrupted session leaves the draft exactly as it was - which is what makes
   * every failure below safe to answer with a plain error.
   *
   * Where the session reports a next range, that number wins over the arithmetic
   * here: it is the server's state, and honouring it is what turns a retried
   * chunk into a resumed upload rather than a refused one. It may not go
   * backwards, because that is a loop that never ends.
   */
  async function uploadThroughSession(
    messageId: string,
    source: UploadSource,
  ): Promise<UploadedAttachment> {
    const session = await client
      .api(`${messagePath(messageId)}/attachments/createUploadSession`)
      .post<UploadSessionResponse>({
        AttachmentItem: {
          attachmentType: "file",
          name: source.name,
          contentType: source.contentType,
          size: source.size,
        },
      });
    const uploadUrl = assertSessionUrl(session?.uploadUrl);
    const expiresAt = Date.parse(session?.expirationDateTime ?? "");
    let offset = nextOffsetOf(session?.nextExpectedRanges) ?? 0;
    while (offset < source.size) {
      if (!Number.isNaN(expiresAt) && Date.now() >= expiresAt)
        throw new Error(
          `The upload session for ${source.name} expired with ${source.size - offset} of ${source.size} bytes still to send, so nothing was attached. Call add_attachment again.`,
        );
      const length = Math.min(uploadChunkBytes, source.size - offset);
      const bytes = await source.read(offset, length);
      if (bytes.byteLength !== length)
        throw new Error(
          `${source.name} gave ${bytes.byteLength} bytes where ${length} were expected at offset ${offset}, so nothing was attached. The file is not the size it said it was.`,
        );
      const stop = stopwatch();
      const response = await put(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Type": "application/octet-stream",
          "Content-Range": `bytes ${offset}-${offset + length - 1}/${source.size}`,
        },
        body: bytes,
      });
      const text = (await response.body().catch(() => "")) ?? "";
      log.debug("attachment chunk uploaded", {
        method: "PUT",
        // Never the URL: it carries the session's bearer token. The range says
        // everything a person debugging needs.
        path: "attachment session",
        range: `${offset}-${offset + length - 1}/${source.size}`,
        ms: stop(),
        status: response.status,
      });
      if (response.status === 201)
        return {
          ...(attachmentIdOf(response.header("Location")) === undefined
            ? {}
            : { id: attachmentIdOf(response.header("Location")) }),
          name: source.name,
          contentType: source.contentType,
          bytes: source.size,
          transport: "upload_session",
        };
      if (response.status !== 200)
        throw new Error(
          `Graph refused bytes ${offset}-${offset + length - 1} of ${source.name} with HTTP ${response.status}: ${reasonOfChunk(text)}. Nothing was attached to the draft.`,
        );
      const resumed = nextOffsetOf(rangesOf(text));
      const sent = offset + length;
      if (resumed === undefined || resumed === sent) {
        offset = sent;
        continue;
      }
      if (resumed <= offset)
        throw new Error(
          `Graph asked for byte ${resumed} of ${source.name} after accepting byte ${sent - 1}, so the upload stopped rather than repeat itself. Nothing was attached.`,
        );
      offset = resumed;
    }
    throw new Error(
      `Graph accepted all ${source.size} bytes of ${source.name} but never confirmed the upload, so it is not known whether the file is on the draft. Check with list_attachments before adding it again.`,
    );
  }

  /**
   * Put one file on a draft, by whichever of the two ways Graph wants: a single
   * `POST` while it is under 3 MB, an upload session above it. The size is
   * checked against the configured ceiling first, so a file that cannot be
   * accepted costs no request at all, and the draft is not touched.
   */
  async function uploadAttachment(
    messageId: string,
    source: UploadSource,
  ): Promise<{ draft: GraphDraftRecord; attachment: UploadedAttachment }> {
    if (!Number.isSafeInteger(source.size) || source.size <= 0)
      throw new Error(
        `${source.name} has no bytes to attach: Graph takes no attachment with nothing in it, and nothing was changed.`,
      );
    if (source.size > maxUploadBytes)
      throw new Error(
        `${source.name} is ${source.size} bytes and the upload limit here is ${maxUploadBytes} bytes (OUTLOOK_MCP_MAX_UPLOAD_BYTES, which Graph caps at ${maxGraphUploadBytes}). A mailbox refuses fewer bytes still at Send, so the draft was not changed.`,
      );
    try {
      const draft = await draftToWrite(messageId);
      if (source.size < singleRequestAttachmentBytes) {
        const bytes = await source.read(0, source.size);
        if (bytes.byteLength !== source.size)
          throw new Error(
            `${source.name} gave ${bytes.byteLength} bytes where ${source.size} were expected, so nothing was attached.`,
          );
        const attachment = await client
          .api(`${messagePath(messageId)}/attachments`)
          .post<{ id?: string }>({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: source.name,
            contentType: source.contentType,
            contentBase64: bytes.toString("base64"),
          });
        return {
          draft,
          attachment: {
            ...(attachment?.id === undefined ? {} : { id: attachment.id }),
            name: source.name,
            contentType: source.contentType,
            bytes: source.size,
            transport: "single_request",
          },
        };
      }
      return {
        draft,
        attachment: await uploadThroughSession(messageId, source),
      };
    } finally {
      source.close?.();
    }
  }

  /** Take one file off a draft, and hand back the draft for the answer to echo. */
  async function deleteAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<GraphDraftRecord> {
    const draft = await draftToWrite(messageId);
    await client
      .api(
        `${messagePath(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
      )
      .delete();
    return draft;
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
    setMessageRead,
    uploadAttachment,
    deleteAttachment,
  };
}

export type Graph = ReturnType<typeof createGraph>;
