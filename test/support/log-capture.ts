import { Writable } from "node:stream";
import { configureLogging } from "../../src/core/log.ts";
import type { LogFormat, LogLevel } from "../../src/core/log.ts";

/**
 * A sink a test can read, plus the level and format the lines were written at.
 *
 * The suite runs with the level pinned to `silent` (see `test/setup.ts`), so a
 * test that wants to see a line has to take the logger over — and give it back,
 * or every later test in the file reads its leftovers.
 */
export type LogCapture = {
  lines: string[];
  /** Every line written so far, as one string. */
  text(): string;
  /** The parsed records, for a test that cares about a field rather than a line. */
  records(): Record<string, unknown>[];
  /** One line per call, matched against a whole line of output. */
  matching(pattern: RegExp): string[];
  restore(): void;
};

export function captureLogs(
  options: { level: LogLevel; format?: LogFormat } = { level: "debug" },
): LogCapture {
  const lines: string[] = [];
  const sink = new Writable({
    write(chunk: Buffer | string, _encoding, callback) {
      lines.push(typeof chunk === "string" ? chunk : chunk.toString("utf8"));
      callback();
    },
  });
  configureLogging({
    level: options.level,
    format: options.format ?? "text",
    sink,
    force: true,
  });
  return {
    lines,
    text: () => lines.join(""),
    records: () =>
      lines
        .join("")
        .split("\n")
        .filter((line) => line !== "")
        .map((line) => JSON.parse(line) as Record<string, unknown>),
    matching: (pattern: RegExp) =>
      lines
        .join("")
        .split("\n")
        .filter((line) => pattern.test(line)),
    restore: () =>
      configureLogging({ level: "silent", format: "text", force: true }),
  };
}
