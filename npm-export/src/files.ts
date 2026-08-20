import { randomUUID } from "node:crypto";
import type { BigIntStats } from "node:fs";
import {
  access,
  constants,
  link,
  lstat,
  open,
  realpath,
  stat,
  unlink,
} from "node:fs/promises";
import { basename, dirname, isAbsolute, join, normalize, resolve, sep } from "node:path";
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

export interface ArtifactPublication {
  destination: PreparedDestination;
  bytes: Uint8Array;
}

interface StagedArtifact extends ArtifactPublication {
  temporaryPath: string;
  identity: BigIntStats;
  committed: boolean;
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

function portablePathKey(path: string): string {
  return normalize(path)
    .split(sep)
    .map((segment) => segment.normalize("NFC").replace(/[ .]+$/g, "").toLowerCase())
    .join("/");
}

export async function readStableInputFile(
  inputPath: string,
  maximumBytes: number,
): Promise<StableInputFile> {
  if (!Number.isSafeInteger(maximumBytes) || maximumBytes <= 0) {
    exportError(
      "invalid_argument",
      "input_validation",
      "maximumBytes must be a positive safe integer.",
      "Use the reviewed compressedDocxBytes resource limit.",
    );
  }
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
    const length = Number(before.size);
    const bytes = new Uint8Array(length);
    let offset = 0;
    while (offset < length) {
      const { bytesRead } = await handle.read(bytes, offset, length - offset, offset);
      if (bytesRead === 0) break;
      offset += bytesRead;
    }
    const probe = new Uint8Array(1);
    const extra = await handle.read(probe, 0, 1, offset);
    const after = await handle.stat({ bigint: true });
    const pathAfter = await stat(absolutePath, { bigint: true });
    const realPathAfter = await realpath(absolutePath);
    if (!sameIdentity(before, after)
      || !sameIdentity(after, pathAfter)
      || realPathBefore !== realPathAfter
      || offset !== length || extra.bytesRead !== 0
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
  for (const [kind, value] of Object.entries(destinations)) {
    if (value !== undefined && (typeof value !== "string" || value.trim() === "")) {
      exportError(
        "invalid_argument",
        "input_validation",
        `${kind} must be a non-empty path when provided.`,
        "Remove the destination or provide a valid new path.",
      );
    }
  }
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
  const inputKey = portablePathKey(input.realPath);
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
    const pathKey = portablePathKey(resolvedPath);
    if (!isAbsolute(resolvedPath) || pathKey === inputKey) {
      exportError(
        "invalid_document",
        "input_validation",
        "An output path aliases the input DOCX.",
        "Choose a distinct, new output path.",
        { detail: requestedPath },
      );
    }
    if (seen.has(pathKey)) {
      exportError(
        "invalid_document",
        "input_validation",
        "Two output destinations resolve to the same path.",
        "Choose a distinct path for every requested artifact.",
        { detail: requestedPath },
      );
    }
    seen.add(pathKey);
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

function throwIfPublicationAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  throw new DocxodusExportError(
    "operation_cancelled",
    "output_write",
    "Artifact publication was cancelled.",
    "Retry with a non-aborted signal and new destination paths.",
    { cause: signal.reason, pending: ["artifact publication"] },
  );
}

async function stageArtifact(
  publication: ArtifactPublication,
  signal: AbortSignal | undefined,
): Promise<StagedArtifact> {
  const { destination, bytes } = publication;
  const temporary = join(
    destination.parentPath,
    `.docxodus-${randomUUID()}.tmp`,
  );
  let handle: Awaited<ReturnType<typeof open>> | undefined;
  let initialIdentity: BigIntStats | undefined;
  try {
    throwIfPublicationAborted(signal);
    handle = await open(temporary, "wx", 0o600);
    throwIfPublicationAborted(signal);
    initialIdentity = await handle.stat({ bigint: true });
    let offset = 0;
    while (offset < bytes.byteLength) {
      throwIfPublicationAborted(signal);
      const { bytesWritten } = await handle.write(
        bytes,
        offset,
        bytes.byteLength - offset,
        offset,
      );
      if (bytesWritten === 0) throw new Error("artifact staging write made no progress");
      offset += bytesWritten;
      throwIfPublicationAborted(signal);
    }
    throwIfPublicationAborted(signal);
    await handle.sync();
    throwIfPublicationAborted(signal);
    const identity = await handle.stat({ bigint: true });
    if (!identity.isFile() || identity.dev !== initialIdentity.dev || identity.ino !== initialIdentity.ino
      || identity.size !== BigInt(bytes.byteLength)) {
      throw new Error("staged artifact identity or length changed during writing");
    }
    await handle.close();
    handle = undefined;
    throwIfPublicationAborted(signal);
    const pathIdentity = await lstat(temporary, { bigint: true });
    if (!sameIdentity(identity, pathIdentity) || pathIdentity.isSymbolicLink()) {
      throw new Error("staged artifact path was replaced during writing");
    }
    return { ...publication, temporaryPath: temporary, identity, committed: false };
  } catch (cause) {
    const cleanupFailures: unknown[] = [];
    await handle?.close().catch((error) => cleanupFailures.push(error));
    if (initialIdentity) {
      await unlinkOwnedPath(temporary, initialIdentity, false)
        .catch((cleanupCause) => cleanupFailures.push(cleanupCause));
    }
    const error = new DocxodusExportError(
      "output_write_failure",
      "output_write",
      `The ${destination.kind} artifact could not be staged for writing.`,
      "Verify free space and destination-directory permissions.",
      {
        cause: cleanupFailures.length === 0
          ? cause
          : new AggregateError([cause, ...cleanupFailures]),
        detail: destination.resolvedPath,
      },
    );
    throw cause instanceof DocxodusExportError
      ? republishError(cause, [], cleanupFailures)
      : error;
  }
}

async function syncDirectory(path: string): Promise<void> {
  let handle: Awaited<ReturnType<typeof open>> | undefined;
  try {
    handle = await open(path, "r");
    await handle.sync();
  } catch (cause) {
    const code = (cause as NodeJS.ErrnoException).code;
    if (process.platform === "win32" && new Set(["EINVAL", "ENOTSUP", "EPERM", "EISDIR"]).has(code ?? "")) {
      return;
    }
    throw cause;
  } finally {
    await handle?.close().catch(() => undefined);
  }
}

async function unlinkOwnedPath(
  path: string,
  identity: BigIntStats,
  requireUnmodifiedContent: boolean,
): Promise<void> {
  let current: BigIntStats;
  try {
    current = await lstat(path, { bigint: true });
  } catch (cause) {
    if ((cause as NodeJS.ErrnoException).code === "ENOENT") return;
    throw cause;
  }
  if (!current.isFile() || current.isSymbolicLink()
    || current.dev !== identity.dev || current.ino !== identity.ino
    || (requireUnmodifiedContent
      && (current.size !== identity.size || current.mtimeNs !== identity.mtimeNs))) {
    throw new Error(`Refusing to unlink a replaced or modified path: ${path}`);
  }
  await unlink(path);
}

function republishError(
  error: unknown,
  retainedDestinations: readonly string[],
  cleanupFailures: readonly unknown[],
): DocxodusExportError {
  if (error instanceof DocxodusExportError) {
    return new DocxodusExportError(error.code, error.phase, error.message, error.remediation, {
      detail: error.detail,
      pending: error.pending,
      partUri: error.partUri,
      anchorId: error.anchorId,
      resource: error.resource,
      cause: cleanupFailures.length === 0
        ? error.cause
        : new AggregateError([
            ...(error.cause === undefined ? [] : [error.cause]),
            ...cleanupFailures,
          ]),
      report: error.report,
      committedDestinations: retainedDestinations,
    });
  }
  return new DocxodusExportError(
    "filesystem_failure",
    "filesystem_commit",
    "Artifact publication failed unexpectedly.",
    "Inspect the retained cause and destination filesystem.",
    {
      cause: cleanupFailures.length === 0
        ? error
        : new AggregateError([error, ...cleanupFailures]),
      committedDestinations: retainedDestinations,
    },
  );
}

/** Stage every artifact first, then commit all or roll back every still-owned destination. */
export async function publishNoReplace(
  publications: readonly ArtifactPublication[],
  signal?: AbortSignal,
): Promise<readonly string[]> {
  const staged: StagedArtifact[] = [];
  try {
    throwIfPublicationAborted(signal);
    for (const publication of publications) {
      throwIfPublicationAborted(signal);
      staged.push(await stageArtifact(publication, signal));
    }
    throwIfPublicationAborted(signal);
  } catch (error) {
    const cleanupFailures: unknown[] = [];
    for (const artifact of staged) {
      await unlinkOwnedPath(artifact.temporaryPath, artifact.identity, false)
        .catch((cause) => cleanupFailures.push(cause));
    }
    throw republishError(error, [], cleanupFailures);
  }

  const parents = [...new Set(staged.map((artifact) => artifact.destination.parentPath))].sort();
  try {
    for (const artifact of staged) {
      try {
        throwIfPublicationAborted(signal);
        await link(artifact.temporaryPath, artifact.destination.resolvedPath);
        artifact.committed = true;
        throwIfPublicationAborted(signal);
        const committedIdentity = await lstat(artifact.destination.resolvedPath, { bigint: true });
        if (artifact.identity.dev !== committedIdentity.dev
          || artifact.identity.ino !== committedIdentity.ino
          || artifact.identity.size !== committedIdentity.size
          || artifact.identity.mtimeNs !== committedIdentity.mtimeNs
          || committedIdentity.isSymbolicLink()) {
          throw new Error("committed artifact identity differs from its verified stage");
        }
        throwIfPublicationAborted(signal);
      } catch (cause) {
        if (cause instanceof DocxodusExportError) throw cause;
        exportError(
          "filesystem_failure",
          "filesystem_commit",
          `The ${artifact.destination.kind} artifact could not be committed without replacement.`,
          "Use a filesystem that supports same-filesystem hard-link creation and a new destination.",
          { cause, detail: artifact.destination.resolvedPath },
        );
      }
    }
    for (const parent of parents) {
      throwIfPublicationAborted(signal);
      await syncDirectory(parent);
      throwIfPublicationAborted(signal);
    }
  } catch (error) {
    const retained: string[] = [];
    const cleanupFailures: unknown[] = [];
    for (const artifact of [...staged].reverse()) {
      if (!artifact.committed) continue;
      await unlinkOwnedPath(artifact.destination.resolvedPath, artifact.identity, true)
        .catch((cause) => {
          retained.push(artifact.destination.resolvedPath);
          cleanupFailures.push(cause);
        });
    }
    for (const artifact of staged) {
      await unlinkOwnedPath(artifact.temporaryPath, artifact.identity, false)
        .catch((cause) => cleanupFailures.push(cause));
    }
    for (const parent of parents) {
      await syncDirectory(parent).catch((cause) => cleanupFailures.push(cause));
    }
    throw republishError(error, retained, cleanupFailures);
  }

  const cleanupFailures: unknown[] = [];
  for (const artifact of staged) {
    await unlinkOwnedPath(artifact.temporaryPath, artifact.identity, false)
      .catch((cause) => cleanupFailures.push(cause));
  }
  for (const parent of parents) {
    await syncDirectory(parent).catch((cause) => cleanupFailures.push(cause));
  }
  const committed = staged.map((artifact) => artifact.destination.resolvedPath);
  if (cleanupFailures.length > 0) {
    throw new DocxodusExportError(
      "filesystem_failure",
      "cleanup",
      "Committed artifacts are durable, but transaction staging cleanup did not complete.",
      "Remove the named .docxodus temporary files without modifying committed artifacts.",
      {
        cause: new AggregateError(cleanupFailures),
        committedDestinations: committed,
      },
    );
  }
  return committed;
}

export async function writeNoReplace(
  destination: PreparedDestination,
  bytes: Uint8Array,
): Promise<void> {
  await publishNoReplace([{ destination, bytes }]);
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
