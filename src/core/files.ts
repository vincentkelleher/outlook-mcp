import {
  closeSync,
  constants,
  fstatSync,
  lstatSync,
  mkdirSync,
  openSync,
  readSync,
  realpathSync,
  writeSync,
} from "node:fs";
import { basename, join, resolve, sep } from "node:path";
import { safeFileName } from "./attachments.ts";
import type { Config } from "./config.ts";
import { logger } from "./log.ts";

/**
 * Every byte this server reads or writes on its own disk, and the checks that
 * stand between a caller's string and it. Two other modules used to hold the
 * write half; the read half added by `add_attachment serverPath` belongs beside
 * it because the containment is the same containment, and a guard duplicated is
 * a guard that drifts.
 */
const log = logger("storage");

const inside = (path: string, root: string) =>
  path === root || path.startsWith(root + sep);

/**
 * Every refusal here goes through this function. They are the checks between a
 * sender's file name — or a model's string — and this machine's disk, so when one
 * bites it is worth a warning even though the caller is answered with a reason:
 * a refusal is either a hostile path or a configuration nobody meant.
 */
function refuse(message: string, action: "read" | "write" = "write"): never {
  log.warn(`attachment ${action} refused`, { reason: message });
  throw new Error(message);
}

/**
 * Attachments may only land inside the configured attachment directory, so both
 * halves of the path are checked: a caller-supplied `dir` is resolved against it,
 * the sender's file name is reduced to one leaf, and the final path is verified
 * resolved and real - so neither `..` in a name nor a symlink waiting inside the
 * directory can write outside - before a single byte is written.
 */
/**
 * Open the file without following a symlink, so the gap between the realpath
 * check and the write cannot be used to swap the destination for a link.
 * Defence in depth rather than the guard: whoever could plant that symlink
 * already writes inside a 0700 directory. Where the flag does not exist the
 * realpath check stands alone, and an overwrite still works, which a "wx" flag
 * would break on the second download of the same attachment.
 */
const noFollowWriteFlag: number | string =
  typeof constants.O_NOFOLLOW === "number"
    ? constants.O_WRONLY |
      constants.O_CREAT |
      constants.O_TRUNC |
      constants.O_NOFOLLOW
    : "w";

/**
 * `writeFileSync` types its flag as a string, so a numeric one needs the lower
 * level open. Writing by descriptor also sets the mode at creation; a file that
 * already exists keeps the mode it has, which is what a re-download wants.
 */
export function writeNoFollow(path: string, bytes: Buffer): void {
  const descriptor = openSync(path, noFollowWriteFlag, 0o600);
  try {
    writeSync(descriptor, bytes);
  } finally {
    closeSync(descriptor);
  }
}

export function attachmentPath(
  config: Config,
  dir: string | undefined,
  rawName: string,
): { path: string; name: string } {
  const base = resolve(config.attachmentDir);
  const target = dir ? resolve(base, dir) : base;
  if (!inside(target, base))
    refuse(
      `dir must be inside the attachment directory ${base}, not ${target}.`,
    );
  mkdirSync(target, { recursive: true, mode: 0o700 });
  // Then compare what the paths really are, so a symlink cannot point outside.
  const realBase = realpathSync(base);
  const realTarget = realpathSync(target);
  if (!inside(realTarget, realBase))
    refuse(
      `dir must be inside the attachment directory ${realBase}, not ${realTarget}.`,
    );
  const name = safeFileName(rawName);
  const path = join(target, name);
  // The name is sender data, so its destination is verified rather than trusted:
  // no added segment, and nothing already waiting at that name may lead outside -
  // a symlink there would carry the write with it, dangling or not. Where nothing
  // exists yet, the directory's own real path answers for the file.
  if (!inside(path, target))
    refuse(
      `${name} would be written outside the attachment directory ${base}, not at ${path}.`,
    );
  // A link at the destination goes wherever its author chose, and `realpath`
  // cannot even resolve a dangling one, so no link is written through at all.
  const waiting = lstatSync(path, { throwIfNoEntry: false });
  if (waiting?.isSymbolicLink())
    refuse(
      `${name} is already a symlink, so writing it could land outside the attachment directory ${realBase}.`,
    );
  const real =
    waiting === undefined ? join(realTarget, name) : realpathSync(path);
  if (!inside(real, realBase))
    refuse(
      `${name} would be written outside the attachment directory ${realBase}, not at ${real}.`,
    );
  return { path, name };
}

/** A file inside the attachment directory, opened and sized, ready to be read in pieces. */
export type AttachmentFile = {
  path: string;
  name: string;
  size: number;
  /** `length` bytes from `offset`; a short read returns fewer, which the caller must notice. */
  read(offset: number, length: number): Buffer;
  close(): void;
};

/**
 * The read half of the same containment, for `add_attachment serverPath`.
 *
 * This is the one place a string from a model selects a file to read, so nothing
 * is taken on trust: an absolute path or a `..` resolves out and is refused, a
 * symlink is refused before it is opened and again by `O_NOFOLLOW` at the open,
 * and only a regular file is accepted — a FIFO would block the server, a device
 * would read the machine. The size comes from the descriptor that was opened,
 * not from a path that could have been swapped behind it, so a chunk loop cannot
 * be told a length the bytes do not support.
 */
export function openAttachmentFile(
  config: Config,
  rawPath: string,
): AttachmentFile {
  const base = resolve(config.attachmentDir);
  const path = resolve(base, rawPath);
  if (!inside(path, base))
    refuse(
      `serverPath must be inside the attachment directory ${base}, not ${rawPath}.`,
      "read",
    );
  const waiting = lstatSync(path, { throwIfNoEntry: false });
  if (waiting === undefined)
    refuse(
      `no file at ${rawPath} in the attachment directory ${base}. Save one there with download_attachment delivery=server_file, or use contentBase64.`,
      "read",
    );
  if (waiting.isSymbolicLink())
    refuse(
      `${rawPath} is a symlink, so reading it could read outside the attachment directory ${base}.`,
      "read",
    );
  if (!waiting.isFile())
    refuse(`serverPath must name a regular file, not ${rawPath}.`, "read");
  const realBase = realpathSync(base);
  const real = realpathSync(path);
  if (!inside(real, realBase))
    refuse(
      `${rawPath} resolves to ${real}, outside the attachment directory ${realBase}.`,
      "read",
    );
  const noFollowReadFlag: number | string =
    typeof constants.O_NOFOLLOW === "number"
      ? constants.O_RDONLY | constants.O_NOFOLLOW
      : "r";
  const descriptor = openSync(path, noFollowReadFlag);
  try {
    const opened = fstatSync(descriptor);
    // The descriptor is what will be read, so it is what gets checked a second
    // time: whatever the path said, this is the object behind it.
    if (!opened.isFile())
      refuse(`serverPath must name a regular file, not ${rawPath}.`, "read");
    const size = opened.size;
    return {
      path,
      name: basename(path),
      size,
      read(offset: number, length: number) {
        const buffer = Buffer.allocUnsafe(length);
        const got = readSync(descriptor, buffer, 0, length, offset);
        return buffer.subarray(0, got);
      },
      close: () => closeSync(descriptor),
    };
  } catch (error) {
    closeSync(descriptor);
    throw error;
  }
}
