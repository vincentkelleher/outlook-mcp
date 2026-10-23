import { Writable } from "node:stream";
import pino from "pino";
import type { DestinationStream, Logger as Pino } from "pino";

/**
 * Logging for this project, in one place, because two of its surfaces have a
 * hard constraint and a soft one:
 *
 * - **stdout is a protocol.** Over stdio it carries JSON-RPC and the CLI's JSON
 *   commands carry data a person pipes into `jq`. A log line there breaks a
 *   client or a pipeline, so every line this module writes goes to stderr — the
 *   destination is chosen here and nowhere else, and `configureLogging` only
 *   takes a sink, never a flag meaning "stdout".
 * - **a log line is a copy of a mailbox.** Message ids, sender addresses, search
 *   queries, subjects and bodies are somebody's mail, and stderr is not
 *   owner-only: a container log, a CI artefact or a pasted terminal is a place
 *   mail can end up in without anyone deciding that. So the default level states
 *   *what happened* — the tool, the folder, how many items, how long — and the
 *   identifiers that name one message appear only at `debug`, where someone
 *   asked for them. `debug` is the level a person turns on to investigate, not
 *   the level a machine reads at 3 a.m.
 *
 * pino supplies the level filter, the child bindings and the redaction; the text
 * format is rendered from the same record pino serialises, which is why a value
 * redacted for the JSON format cannot reappear in the human one.
 */

export const logLevels = ["debug", "info", "warn", "error", "silent"] as const;
export type LogLevel = (typeof logLevels)[number];

export const logFormats = ["text", "json"] as const;
export type LogFormat = (typeof logFormats)[number];

/**
 * `info` is the default because the question asked most often of this server is
 * "what did it actually do", and one line per command or tool call is what
 * answers it without a second process. `silent` exists for a test suite that
 * asserts on stderr, and for a container whose operator wants only problems.
 */
export const defaultLogLevel: LogLevel = "info";
export const defaultLogFormat: LogFormat = "text";

/**
 * A level from the environment is read once and fails loudly, because
 * `OUTLOOK_LOG_LEVEL=verbose` silently meaning `info` is how someone debugging
 * at midnight concludes the logging is broken. Case-insensitive, and empty means
 * the default.
 */
export function logLevelFor(name: string, value: string | undefined): LogLevel {
  const raw = (value ?? "").trim().toLowerCase();
  if (raw === "") return defaultLogLevel;
  if (!(logLevels as readonly string[]).includes(raw))
    throw new Error(
      `${name} must be one of ${logLevels.join(", ")}, not "${value}".`,
    );
  return raw as LogLevel;
}

/** The same decision for the line format. */
export function logFormatFor(
  name: string,
  value: string | undefined,
): LogFormat {
  const raw = (value ?? "").trim().toLowerCase();
  if (raw === "") return defaultLogFormat;
  if (!(logFormats as readonly string[]).includes(raw))
    throw new Error(
      `${name} must be one of ${logFormats.join(", ")}, not "${value}".`,
    );
  return raw as LogFormat;
}

/**
 * Keys that never reach a log line at any level: token material, and mail or
 * file content. pino redacts them before serialising, so they are gone from both
 * formats, and a call site that passes one by accident loses the value rather
 * than the secret. `userCode` is deliberately absent — a device-code flow is
 * useless to anyone who cannot see the code, and the caller masks it until
 * someone asks for `debug`.
 */
export const neverLoggedKeys = [
  "accessToken",
  "access_token",
  "refreshToken",
  "refresh_token",
  "idToken",
  "id_token",
  "token",
  "authorization",
  "clientSecret",
  "client_secret",
  "codeVerifier",
  "code_verifier",
  "cacheContent",
  "body",
  "contentBase64",
  "contentBytes",
] as const;

/** Fields a line carries besides its message. */
export type LogFields = Record<string, unknown>;

export type Logger = {
  readonly scope: string;
  /**
   * Whether a `debug` line would be written. Call sites deepen their own fields
   * on this rather than logging twice: the identifiers are in the one line, when
   * somebody asked for them.
   */
  readonly detailed: boolean;
  debug(message: string, fields?: LogFields): void;
  info(message: string, fields?: LogFields): void;
  warn(message: string, fields?: LogFields): void;
  error(message: string, fields?: LogFields): void;
  /** A logger that says which session or call a line belongs to. */
  child(fields: LogFields): Logger;
};

/** pino's numeric levels, spelled the way this module names them. */
const labelOf = new Map<number, string>([
  [10, "trace"],
  [20, "debug"],
  [30, "info"],
  [40, "warn"],
  [50, "error"],
  [60, "fatal"],
]);

/**
 * A value that names one thing in the mailbox: short enough to recognise in a
 * line, long enough to match against the id a caller holds, and never the whole
 * thing where nobody asked.
 */
export function identifier(value: string, detailed: boolean): string {
  if (detailed) return value;
  const head = value.slice(0, 8);
  return value.length > 8 ? `${head}…${value.length}` : head;
}

/** A mailbox address: at default level the domain, and one letter of the person. */
export function address(
  value: string | undefined,
  detailed: boolean,
): string | undefined {
  if (value === undefined || detailed) return value;
  const at = value.indexOf("@");
  if (at < 0) return `${value.slice(0, 1)}…`;
  return `${value.slice(0, 1)}***${value.slice(at)}`;
}

/** A device-code flow's code, which the user is shown anyway, until `debug`. */
export function verificationCode(
  value: string | undefined,
  detailed: boolean,
): string | undefined {
  if (value === undefined || detailed) return value;
  return `${value.slice(0, 4)}***`;
}

const identifiers = new Set([
  "id",
  "messageId",
  "attachmentId",
  "cursor",
  "account",
  "email",
]);

/**
 * Mail content, and the file name that came with it: at default level their
 * size, never their text. A folder is not content — `inbox` and `drafts` are
 * this server's own vocabulary, and which folder failed is the first thing
 * anybody asks.
 */
const content = new Set([
  "query",
  "subject",
  "body",
  "bodyPreview",
  "comment",
  "name",
  "dir",
  "contentBase64",
  "contentBytes",
]);

/** Recipients, at default level, as a number: who they are is the user's business. */
const recipientLists = new Set(["to", "cc", "bcc", "recipients"]);

/**
 * Tool arguments or CLI flags as fields that are safe to write at the default
 * level: switches, counts and sizes in place of the strings a mailbox is made of.
 * At `debug` the arguments go through as they came, which is what reproducing a
 * call needs.
 */
export function fieldsForLogging(args: unknown, detailed: boolean): LogFields {
  if (typeof args !== "object" || args === null) return {};
  const out: LogFields = {};
  for (const [key, value] of Object.entries(args)) {
    if (value === undefined) continue;
    if (detailed) {
      out[key] = value;
      continue;
    }
    if (recipientLists.has(key)) {
      out[`${key}Count`] = Array.isArray(value) ? value.length : 1;
    } else if (identifiers.has(key) && typeof value === "string") {
      out[key] =
        key === "account" || key === "email"
          ? address(value, false)
          : identifier(value, false);
    } else if (content.has(key)) {
      if (typeof value === "string") out[`${key}Chars`] = value.length;
    } else if (typeof value === "object") {
      out[`${key}Present`] = true;
    } else {
      out[key] = value;
    }
  }
  return out;
}

/**
 * An error as fields, with a stack only where somebody is debugging. Microsoft's
 * own text is in there, and it is the reason a call failed — which is what a
 * warning line is for; the stack is for the one reading code afterwards.
 */
export function errorFields(error: unknown, detailed: boolean): LogFields {
  const fields: LogFields = {
    reason: error instanceof Error ? error.message : String(error),
  };
  if (!(error instanceof Error)) return fields;
  const named = error as { name?: string; cause?: unknown; stack?: string };
  if (named.name && named.name !== "Error") fields.errorName = named.name;
  if (!detailed) return fields;
  if (named.stack) fields.stack = named.stack;
  const cause = named.cause;
  if (cause !== undefined && cause !== null)
    fields.causedBy =
      cause instanceof Error
        ? `${cause.name}: ${cause.message}`
        : String(cause);
  return fields;
}

let level: LogLevel = defaultLogLevel;
let format: LogFormat = defaultLogFormat;
let pinned = false;
let sink: NodeJS.WritableStream = process.stderr;
let root: Pino | undefined;

/** Cached per scope and bindings, because a line should not build a logger. */
const instances = new Map<string, Pino>();

/**
 * How many bound loggers the cache holds. A session id is bounded by the session
 * cap, but a resource read binds a message id, and a server that runs for a week
 * would otherwise accumulate one entry per message anybody looked at. Dropping
 * the cache costs a pino `child()` per line until it fills again, never a lost
 * line and never a growing heap.
 */
const maxCachedLoggers = 128;

/**
 * pino hands us one serialised record per line. Rendering the human format from
 * that record, rather than from the arguments before serialisation, is what makes
 * redaction unconditional: there is no path around it.
 *
 * The callback runs synchronously so nothing sits in a buffer when a command
 * exits on an error, and a line that will not parse is written as it arrived —
 * a diagnostic is never allowed to eat another diagnostic.
 */
function destination(): DestinationStream {
  return new Writable({
    write(chunk: Buffer | string, _encoding, callback) {
      const line = typeof chunk === "string" ? chunk : chunk.toString("utf8");
      sink.write(format === "json" ? line : textOf(line));
      callback();
    },
    // Ending a logger never ends the real stderr, which outlives any one
    // composition of them.
    final(callback) {
      callback();
    },
  }) as unknown as DestinationStream;
}

/** One record, read back as the line a person reads. */
export function textOf(record: string): string {
  let parsed: Record<string, unknown>;
  try {
    parsed = JSON.parse(record) as Record<string, unknown>;
  } catch {
    return record;
  }
  const { level: numeric, time, pid, scope, msg, ...rest } = parsed;
  const parts = [
    typeof time === "string" ? time : new Date(Number(time) || 0).toISOString(),
    labelOf.get(Number(numeric)) ?? String(numeric ?? "log"),
    pid === undefined ? undefined : String(pid),
    `${String(scope ?? "outlook")}:`,
    typeof msg === "string" ? msg : "",
  ];
  const fields: string[] = [];
  for (const [key, value] of Object.entries(rest)) {
    const rendered =
      typeof value === "string"
        ? // Quote only what needs it. A masked address, a URL, a percent and a
          // comma-separated list stay bare because they are the values a person
          // scans for; anything with a space, a `=` or a brace is JSON-quoted so
          // the `key=value` shape of the line survives.
          /^[\w+./:@#*,%;?…-]+$/u.test(value)
          ? value
          : JSON.stringify(value)
        : JSON.stringify(value);
    fields.push(`${key}=${rendered}`);
  }
  return `${parts.filter((part) => part !== undefined).join(" ")}${
    fields.length > 0 ? ` ${fields.join(" ")}` : ""
  }\n`;
}

function build(): Pino {
  return pino(
    {
      // pino's own `fatal` and `trace` stay unused: this project reports at the
      // four levels its callers mean, and `silent` turns them all off.
      level,
      base: { pid: process.pid },
      timestamp: pino.stdTimeFunctions.isoTime,
      redact: { paths: [...neverLoggedKeys], censor: "[redacted]" },
    },
    destination(),
  );
}

/**
 * Set the level and format from configuration, and — for tests — redirect the
 * sink. Everything else is a no-op, so it is safe to call from every entry point
 * in any order: the logger exists before the configuration is read, and a person
 * starting `outlook emails list` wants their `OUTLOOK_LOG_LEVEL` honoured for the
 * very first line.
 */
export function configureLogging(options: {
  level?: LogLevel;
  format?: LogFormat;
  sink?: NodeJS.WritableStream;
  /**
   * Claim the level and format for the whole process, so the configuration a
   * command is about to apply cannot raise them again. A test suite that
   * asserts on stderr needs this — otherwise every `serve()` and every
   * `main()` in the suite re-applies its own default and the output swallows
   * the assertion that failed.
   */
  force?: boolean;
}): void {
  if (options.level !== undefined && (!pinned || options.force))
    level = options.level;
  if (options.format !== undefined && (!pinned || options.force))
    format = options.format;
  if (options.force) pinned = true;
  if (options.sink !== undefined && options.sink !== sink) {
    sink = options.sink;
    root = undefined;
    // Every cached instance belongs to the old root, so the cache goes with it:
    // a test that redirects the sink after some other module has logged still
    // sees its own lines.
    instances.clear();
  }
  // pino's children follow their parent's level, so this reaches every logger
  // this process has already handed out.
  if (root) root.level = level;
}

export function currentLogLevel(): LogLevel {
  return level;
}

export function currentLogFormat(): LogFormat {
  return format;
}

function instance(scope: string, bindings: LogFields): Pino {
  const key = `${scope}|${JSON.stringify(bindings)}`;
  const cached = instances.get(key);
  if (cached) return cached;
  if (!root) root = build();
  // The scope is a binding like any other, so a line says who wrote it in both
  // formats and `grep scope=mcp` works on a json log.
  const created = root.child({ scope, ...bindings });
  if (instances.size >= maxCachedLoggers) instances.clear();
  instances.set(key, created);
  return created;
}

function wrap(scope: string, bound: () => LogFields): Logger {
  const at = () => instance(scope, bound());
  return {
    scope,
    get detailed() {
      return at().isLevelEnabled("debug");
    },
    debug: (message, fields) => at().debug(fields ?? {}, message),
    info: (message, fields) => at().info(fields ?? {}, message),
    warn: (message, fields) => at().warn(fields ?? {}, message),
    error: (message, fields) => at().error(fields ?? {}, message),
    // Resolved at write time, so a child of a session's logger keeps following
    // that session instead of freezing the bindings it happened to see.
    child: (extra) => wrap(scope, () => ({ ...bound(), ...extra })),
  };
}

/**
 * A logger whose bindings are chosen when the line is written, not when the
 * logger is made. An HTTP session needs this: its id exists after the MCP server
 * that has to log with it, and a session's tool calls are the lines worth
 * finding afterwards.
 */
export function scopedLogger(scope: string, bindings: () => LogFields): Logger {
  return wrap(scope, bindings);
}

const scoped = new Map<string, Logger>();

/** The logger for a module: `logger("auth")` says who wrote the line. */
export function logger(scope: string): Logger {
  const existing = scoped.get(scope);
  if (existing) return existing;
  const created = wrap(scope, () => ({}));
  scoped.set(scope, created);
  return created;
}

/**
 * Time one call. `const stop = stopwatch()` at the top and `stop()` in the line
 * at the bottom, so a call that threw still has a duration when it is the reason
 * someone is reading a warning.
 */
export function stopwatch(): () => number {
  const started = process.hrtime.bigint();
  return () =>
    Math.max(0, Math.round(Number(process.hrtime.bigint() - started) / 1e6));
}
