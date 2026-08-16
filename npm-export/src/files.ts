import { randomUUID } from "node:crypto";
import type { BigIntStats } from "node:fs";
import {
  access,
  constants,
  link,
  open,
  realpath,
  stat,
  unlink,
} from "node:fs/promises";
import { basename, dirname, isAbsolute, join, normalize, resolve } from "node:path";
import type { RenderFileDestinations } from "./contracts.js";
import { DocxodusExportError, exportError } from "./contracts.js";

export interface StableInputFile {
  bytes: Uint8Array;
  absolutePath: string;
  realPath: string;
}

export interface PreparedDestination {
  kind: keyof RenderFileDestinations;
  requestedPath: string;
  absolutePath: string;
  resolvedPath: string;
  parentPath: string;
}

function sameIdentity(
  left: BigIntStats,
  right: BigIntStats,
): boolean {
  return left.dev === right.dev
    && left.ino === right.ino
    && left.size === right.size
    && left.mtimeNs === right.mtimeNs
    && left.ctimeNs === right.ctimeNs;
}

export async function readStableInputFile(
  inputPath: string,
  maximumBytes: number,
): Promise<StableInputFile> {
  if (typeof inputPath !== "string" || inputPath.trim() === "") {
    exportError("invalid_document", "input_validation", "An input path is required.",
      "Pass a path to one regular DOCX file.");
  }
  const absolutePath = resolve(inputPath);
  let handle: Awaited<ReturnType<typeof open>> | undefined;
  try {
    handle = await open(absolutePath, "r");
    const realPathBefore = await realpath(absolutePath);
    const before = await handle.stat({ bigint: true });
    if (!before.isFile()) {
      exportError("invalid_document", "input_validation", "The input path is not a regular file.",
        "Pass one stable DOCX file.");
    }
    if (before.size > BigInt(maximumBytes)) {
      exportError(
        "resource_limit",
        "package_preflight",
        `compressedDocxBytes limit exceeded (${before.size} > ${maximumBytes}).`,
        "Use a smaller document or a deployment with a reviewed limits contract.",
      );
    }
    const bytes = new Uint8Array(await handle.readFile());
    const after = await handle.stat({ bigint: true });
    const pathAfter = await stat(absolutePath, { bigint: true });
    const realPathAfter = await realpath(absolutePath);
    if (!sameIdentity(before, after)
      || !sameIdentity(after, pathAfter)
      || realPathBefore !== realPathAfter
      || BigInt(bytes.byteLength) !== after.size) {
      exportError(
        "invalid_document",
        "input_validation",
        "The input file changed while it was being read.",
        "Retry after the producer has finished writing the DOCX.",
      );
    }
    return { bytes, absolutePath, realPath: realPathAfter };
  } catch (cause) {
    if (cause instanceof DocxodusExportError) throw cause;
    return exportError(
      "invalid_document",
      "input_validation",
      `The input DOCX could not be read: ${absolutePath}`,
      "Verify the path and file permissions.",
      { cause },
    );
  } finally {
    await handle?.close().catch(() => undefined);
  }
}

export async function prepareDestinations(
  input: StableInputFile,
  destinations: RenderFileDestinations,
): Promise<PreparedDestination[]> {
  const entries = Object.entries(destinations)
    .filter((entry): entry is [keyof RenderFileDestinations, string] =>
      typeof entry[1] === "string" && entry[1].trim() !== "");
  if (entries.length === 0) {
    exportError(
      "invalid_document",
      "input_validation",
      "At least one output destination is required.",
      "Provide an HTML, PDF, PageMap, or render-report path.",
    );
  }

  const seen = new Set<string>();
  const prepared: PreparedDestination[] = [];
  for (const [kind, requestedPath] of entries) {
    const absolutePath = resolve(requestedPath);
    const parentPath = await realpath(dirname(absolutePath)).catch((cause) =>
      exportError(
        "filesystem_failure",
        "input_validation",
        `The destination parent does not exist or cannot be resolved: ${dirname(absolutePath)}`,
        "Create the parent directory before rendering.",
        { cause },
      ));
    const resolvedPath = normalize(join(parentPath, basename(absolutePath)));
    if (!isAbsolute(resolvedPath) || resolvedPath === input.realPath) {
      exportError(
        "invalid_document",
        "input_validation",
        "An output path aliases the input DOCX.",
        "Choose a distinct, new output path.",
        { detail: requestedPath },
      );
    }
    if (seen.has(resolvedPath)) {
      exportError(
        "invalid_document",
        "input_validation",
        "Two output destinations resolve to the same path.",
        "Choose a distinct path for every requested artifact.",
        { detail: requestedPath },
      );
    }
    seen.add(resolvedPath);
    try {
      await access(resolvedPath, constants.F_OK);
      exportError(
        "filesystem_failure",
        "input_validation",
        `The destination already exists: ${resolvedPath}`,
        "Choose a new path; Docxodus never overwrites an artifact.",
      );
    } catch (cause) {
      if (cause instanceof DocxodusExportError) throw cause;
      const code = (cause as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") {
        exportError(
          "filesystem_failure",
          "input_validation",
          `The destination could not be inspected safely: ${resolvedPath}`,
          "Verify filesystem permissions and choose a new path.",
          { cause },
        );
      }
    }
    prepared.push({ kind, requestedPath, absolutePath, resolvedPath, parentPath });
  }
  return prepared;
}

export async function writeNoReplace(
  destination: PreparedDestination,
  bytes: Uint8Array,
): Promise<void> {
  const temporary = join(
    destination.parentPath,
    `.docxodus-${randomUUID()}.tmp`,
  );
  let handle: Awaited<ReturnType<typeof open>> | undefined;
  let committed = false;
  let primaryError: unknown;
  try {
    try {
      handle = await open(temporary, "wx", 0o600);
      await handle.writeFile(bytes);
      await handle.sync();
      await handle.close();
      handle = undefined;
    } catch (cause) {
      exportError(
        "output_write_failure",
        "output_write",
        `The ${destination.kind} artifact could not be staged for writing.`,
        "Verify free space and destination-directory permissions.",
        { cause, detail: destination.resolvedPath },
      );
    }

    try {
      await link(temporary, destination.resolvedPath);
      committed = true;
    } catch (cause) {
      exportError(
        "filesystem_failure",
        "filesystem_commit",
        `The ${destination.kind} artifact could not be committed without replacement.`,
        "Use a filesystem that supports same-filesystem hard-link creation and a new destination.",
        { cause, detail: destination.resolvedPath },
      );
    }
  } catch (error) {
    primaryError = error;
    throw error;
  } finally {
    await handle?.close().catch(() => undefined);
    try {
      await unlink(temporary);
    } catch (cause) {
      if ((cause as NodeJS.ErrnoException).code !== "ENOENT" && primaryError === undefined) {
        throw new DocxodusExportError(
          "filesystem_failure",
          "cleanup",
          `The staged ${destination.kind} artifact could not be removed after commit.`,
          "Remove the named .docxodus temporary file after verifying the committed artifact.",
          {
            cause,
            detail: temporary,
            committedDestinations: committed ? [destination.resolvedPath] : [],
          },
        );
      }
    }
  }
}

export async function writeDiagnosticNoReplace(path: string, bytes: Uint8Array): Promise<void> {
  const absolutePath = resolve(path);
  const parentPath = await realpath(dirname(absolutePath));
  const destination: PreparedDestination = {
    kind: "reportPath",
    requestedPath: path,
    absolutePath,
    resolvedPath: join(parentPath, basename(absolutePath)),
    parentPath,
  };
  await writeNoReplace(destination, bytes);
}
