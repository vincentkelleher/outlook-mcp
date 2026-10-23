import type { IncomingMessage, ServerResponse } from "node:http";

/**
 * The HTTP half of the MCP surface: who may talk to a server that authenticates
 * nobody, and how a request body reaches the transport.
 *
 * The transport itself comes from the MCP SDK. What lives here is the decision
 * made before it, because an unauthenticated port on a machine that holds a
 * mailbox is only as private as the requests it accepts, and the two ways a
 * request arrives that nobody meant to allow are the two checked here: a `Host`
 * header pointing somewhere else (DNS rebinding: a page you visited re-resolves
 * `localhost` to your own machine), and a browser `Origin` that is not the one
 * the server was addressed as.
 */

/** Nothing else can answer to these, so nothing else may be served as them. */
const loopbackHosts = new Set(["127.0.0.1", "localhost", "::1", "[::1]"]);

/** Where the MCP endpoint lives, named once: the server serves it, `config` prints it. */
export const mcpPath = "/mcp";
export const healthPath = "/health";

/** The bound the MCP transport itself puts on a body it reads; the same here. */
export const maxRequestBytes = 4 * 1024 * 1024;

/** `127.0.0.1:3939`, `[::1]:3939`, `localhost` — host and port, lower case. */
function authority(value: string): { host: string; port: string | undefined } {
  const raw = value.trim().toLowerCase();
  if (raw.startsWith("[")) {
    const end = raw.indexOf("]");
    if (end > 0) {
      const port = raw.slice(end + 1).replace(/^:/, "");
      return {
        host: raw.slice(0, end + 1),
        port: port === "" ? undefined : port,
      };
    }
  }
  const colon = raw.lastIndexOf(":");
  if (colon > 0 && !raw.slice(colon + 1).includes(":")) {
    const port = raw.slice(colon + 1);
    return { host: raw.slice(0, colon), port: port === "" ? undefined : port };
  }
  return { host: raw, port: undefined };
}

/** What a URL names, or nothing at all when it is not a URL. */
function originAuthority(
  value: string,
): { host: string; port?: string } | undefined {
  try {
    const url = new URL(value);
    // `[::1]` for an IPv6 literal, which is how `authority()` spells it too.
    return { host: url.hostname.toLowerCase(), port: url.port || undefined };
  } catch {
    return undefined;
  }
}

/**
 * One origin, for real: the same host, and the same port unless one side names
 * none. A bare `localhost` means the scheme's own port, which there is nothing
 * to compare against.
 */
function sameOrigin(origin: string, host: string): boolean {
  const from = originAuthority(origin);
  if (!from) return false;
  const named = authority(host);
  if (from.host !== named.host) return false;
  return !from.port || !named.port || from.port === named.port;
}

export function isLoopbackHost(value: string): boolean {
  return loopbackHosts.has(authority(value).host);
}

export type Admission =
  { allowed: true } | { allowed: false; status: number; reason: string };

/**
 * May this request be served?
 *
 * `Host` must name this machine as a loopback address, or as one of
 * `allowedHosts`: a request that arrived under some other name got here through
 * a name that was pointed at this machine, which is what DNS rebinding means.
 * An `Origin` — only a browser sends one — must be the origin the request was
 * addressed as, or one of `allowedOrigins`. A request with no `Origin` is a
 * native MCP client, which is the ordinary case here, and is answered on the
 * strength of its `Host` alone.
 */
export function admitRequest(input: {
  host?: string | string[];
  origin?: string | string[];
  allowedHosts: string[];
  allowedOrigins: string[];
}): Admission {
  const host = typeof input.host === "string" ? input.host : undefined;
  if (!host || host.trim() === "")
    return {
      allowed: false,
      status: 400,
      reason: "The request carries no Host header.",
    };
  const named = host.trim().toLowerCase();
  const loopback = isLoopbackHost(named);
  if (!loopback && !input.allowedHosts.includes(authority(named).host))
    return {
      allowed: false,
      status: 403,
      reason:
        `The Host header names "${named}", which this server does not recognise as itself. ` +
        `Set OUTLOOK_MCP_ALLOWED_HOSTS to include it if you reach this server by that name.`,
    };

  const origin = typeof input.origin === "string" ? input.origin : undefined;
  if (origin && origin.trim() !== "") {
    if (input.allowedOrigins.includes(origin.trim().toLowerCase()))
      return { allowed: true };
    // Same-origin against the address the request itself claims, so a page on
    // another port, or on another name, is not the thing that was addressed.
    if (sameOrigin(origin, named)) return { allowed: true };
    return {
      allowed: false,
      status: 403,
      reason:
        `The Origin header says "${origin.trim()}", which is not the origin this request was ` +
        `addressed as. A browser that is not the one this server serves must be named in ` +
        `OUTLOOK_MCP_ALLOWED_ORIGINS.`,
    };
  }
  return { allowed: true };
}

export type Body =
  | { allowed: true; value: unknown }
  | { allowed: false; status: number; reason: string };

/**
 * Read and parse a request body under a hard byte cap, so a caller cannot make
 * the process buffer without limit. The parsed value goes to the transport as
 * `parsedBody`, which is what stops the transport reading the stream a second
 * time.
 */
export async function readJsonBody(
  req: IncomingMessage,
  maxBytes = maxRequestBytes,
): Promise<Body> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > maxBytes)
      return {
        allowed: false,
        status: 413,
        reason: `The request body is over ${maxBytes} bytes.`,
      };
    chunks.push(chunk as Buffer);
  }
  const text = Buffer.concat(chunks).toString("utf8");
  if (text === "")
    return {
      allowed: false,
      status: 400,
      reason: "POST to /mcp needs a JSON-RPC body.",
    };
  try {
    return { allowed: true, value: JSON.parse(text) };
  } catch {
    return {
      allowed: false,
      status: 400,
      reason: "The request body is not JSON.",
    };
  }
}

/** One JSON answer, written once, with the headers a client reads. */
export function respondJson(
  res: ServerResponse,
  status: number,
  payload: Record<string, unknown>,
): void {
  if (res.headersSent || res.writableEnded) return;
  const text = `${JSON.stringify(payload)}\n`;
  res.writeHead(status, {
    "content-type": "application/json",
    "content-length": Buffer.byteLength(text),
    "cache-control": "no-store",
  });
  res.end(text);
}
