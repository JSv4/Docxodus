import { constants } from "node:fs";
import type { BigIntStats } from "node:fs";
import { lstat, open, opendir, realpath, stat } from "node:fs/promises";
import { extname, isAbsolute, join, relative, resolve, sep } from "node:path";
import * as fontkit from "fontkit";
import {
  fontFamilyKey,
  normalizeFontFamilyName,
  type ExportResourceLimits,
  type FontFaceStyle,
  type FontFileFormat,
  type FontLicenseEvidence,
  type FontMediaType,
} from "docxodus/export-browser";
import { canonicalJson, sha256 } from "../canonical.js";
import type { FontLicenseAttestation, RenderOutput } from "../contracts.js";
import { DocxodusExportError, exportError } from "../contracts.js";

const FONT_EXTENSIONS = new Set([".ttf", ".otf", ".woff", ".woff2"]);
const WIDTH_CLASS_PERCENT = Object.freeze([0, 50, 62.5, 75, 87.5, 100, 112.5, 125, 150, 200]);
const FONT_CATALOG_CODE_POINTS_MAX = 0x11_0000;

export interface ConfiguredFontFace {
  id: string;
  directoryIndex: number;
  discoveryIndex: number;
  family: string;
  familyKey: string;
  postscriptName?: string;
  version: string;
  style: FontFaceStyle;
  weight: number;
  stretch: number;
  format: FontFileFormat;
  mediaType: FontMediaType;
  byteLength: number;
  expandedByteLength: number;
  sha256: string;
  bytes: Uint8Array;
  codePoints: readonly number[];
  permittedOutputs: readonly RenderOutput[];
  license: { ok: true; evidence: FontLicenseEvidence } | { ok: false; failure: string };
}

export interface FontCatalog {
  faces: readonly ConfiguredFontFace[];
  directoryCount: number;
  entryCount: number;
  fileCount: number;
  totalBytes: number;
  totalExpandedBytes: number;
}

interface Snapshot {
  bytes: Uint8Array;
}

interface DecodedFormat {
  format: FontFileFormat;
  mediaType: FontMediaType;
  expandedByteLength: number;
}

interface DirectorySnapshot {
  path: string;
  identity: BigIntStats;
}

function sameIdentity(left: BigIntStats, right: BigIntStats): boolean {
  return left.dev === right.dev
    && left.ino === right.ino
    && left.size === right.size
    && left.mtimeNs === right.mtimeNs
    && left.ctimeNs === right.ctimeNs;
}

function underRoot(root: string, candidate: string): boolean {
  const fromRoot = relative(root, candidate);
  return fromRoot === "" || (fromRoot !== ".." && !fromRoot.startsWith(`..${sep}`)
    && !isAbsolute(fromRoot));
}

function fontError(message: string, remediation: string, detail?: string, cause?: unknown): never {
  exportError("resource_policy_failure", "font_loading", message, remediation, {
    ...(detail ? { detail } : {}),
    ...(cause === undefined ? {} : { cause }),
  });
}

function limitError(name: keyof ExportResourceLimits, actual: number, maximum: number): never {
  exportError(
    "resource_limit",
    "font_loading",
    `${name} limit exceeded (${actual} > ${maximum}).`,
    "Use fewer or smaller configured font files, or lower document font diversity.",
  );
}

function enforceLimit(
  name: keyof ExportResourceLimits,
  actual: number,
  maximum: number,
): void {
  if (actual > maximum) limitError(name, actual, maximum);
}

function wellFormed(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(++index);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return false;
  }
  return true;
}

function safeMetadata(value: unknown, label: string, required = true): string | undefined {
  if (value === undefined || value === null || value === "") {
    if (!required) return undefined;
    fontError(`A configured font has no usable ${label}.`, "Replace the malformed font file.");
  }
  const text = String(value);
  if (!wellFormed(text)) {
    fontError(`A configured font has invalid ${label} Unicode.`, "Replace the malformed font file.");
  }
  const normalized = normalizeFontFamilyName(text);
  if (normalized.length === 0 || normalized.length > 256 || /[\u0000-\u001f\u007f]/u.test(normalized)) {
    fontError(`A configured font has invalid ${label} metadata.`, "Replace the malformed font file.");
  }
  return normalized;
}

// Duplicated (not shared) with npm/src/font-runtime.ts's readU32: this realm has no access to
// that package's internals, and the logic is identical enough that inlining beats a public
// cross-package export just to save these four lines.
function readU32(bytes: Uint8Array, offset: number): number {
  if (bytes.byteLength < offset + 4) {
    fontError("A configured font has a truncated format header.", "Replace the malformed font file.");
  }
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, false);
}

// The signature bytes here (wOFF/wOF2/OTTO/0x00010000/"true") are also asserted independently
// by npm/src/font-runtime.ts's fontSignatureMatches, which verifies a resolver's declared
// format against the bytes it actually sent. A new font format needs both updated.
function decodedFormat(bytes: Uint8Array, extension: string): DecodedFormat {
  if (bytes.byteLength < 4) {
    fontError("A configured font is too short to contain a valid header.", "Replace the malformed font file.");
  }
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  let format: FontFileFormat;
  let mediaType: FontMediaType;
  let expandedByteLength = bytes.byteLength;
  if (signature === "wOFF") {
    format = "woff";
    mediaType = "font/woff";
    if (readU32(bytes, 8) !== bytes.byteLength) {
      fontError("A configured WOFF file has an inconsistent declared length.",
        "Replace the malformed webfont file.");
    }
    expandedByteLength = readU32(bytes, 16);
  } else if (signature === "wOF2") {
    format = "woff2";
    mediaType = "font/woff2";
    if (readU32(bytes, 8) !== bytes.byteLength) {
      fontError("A configured WOFF2 file has an inconsistent declared length.",
        "Replace the malformed webfont file.");
    }
    expandedByteLength = readU32(bytes, 16);
  } else if (signature === "OTTO") {
    format = "otf";
    mediaType = "font/otf";
  } else if (readU32(bytes, 0) === 0x0001_0000 || signature === "true") {
    format = "ttf";
    mediaType = "font/ttf";
  } else {
    fontError("A configured font has an unsupported decoded format.",
      "Provide a valid TTF, OTF, WOFF, or WOFF2 font file.");
  }
  if (`.${format}` !== extension) {
    fontError("A configured font extension does not match its decoded format.",
      "Rename or replace the font so its extension and format agree.");
  }
  if (expandedByteLength <= 0) {
    fontError("A configured webfont declares an invalid expanded size.", "Replace the malformed font file.");
  }
  return { format, mediaType, expandedByteLength };
}

function os2LicenseEvidence(
  face: fontkit.Font,
  fileSha256: string,
): { evidence?: FontLicenseEvidence; failure?: string; prohibited?: boolean } {
  const os2 = face["OS/2"];
  const flags = os2?.fsType;
  if (!flags) return { failure: "The font has no readable OS/2 embedding-rights table." };
  if (flags.noEmbedding) {
    return { failure: "The OS/2 table forbids font embedding.", prohibited: true };
  }
  if (flags.bitmapOnly) {
    return { failure: "The OS/2 table permits bitmap embedding only.", prohibited: true };
  }
  if (Number(flags.viewOnly) + Number(flags.editable) > 1) {
    return {
      failure: "The OS/2 table has contradictory embedding-rights bits.",
      prohibited: true,
    };
  }
  const kind = flags.editable ? "editable" : flags.viewOnly ? "previewPrint" : "installable";
  const evidence = {
    kind,
    noSubsetting: Boolean(flags.noSubsetting),
    fileSha256,
    source: "os2-fstype",
  } as const;
  return {
    evidence: {
      kind,
      identity: sha256(canonicalJson(evidence)),
      noSubsetting: Boolean(flags.noSubsetting),
    },
  };
}

function attestedLicenseEvidence(
  attestation: FontLicenseAttestation | undefined,
  os2Evidence: FontLicenseEvidence | undefined,
): FontLicenseEvidence | undefined {
  if (!attestation) return undefined;
  return {
    kind: "attested",
    identity: sha256(canonicalJson({
      schemaVersion: attestation.schemaVersion,
      usage: attestation.usage,
      fileSha256: attestation.fileSha256,
      embeddingPermitted: true,
      permittedOutputs: attestation.permittedOutputs,
      subsettingPermitted: attestation.subsettingPermitted,
      basis: attestation.basis,
      ...(attestation.attester ? { attester: attestation.attester } : {}),
    })),
    noSubsetting: os2Evidence?.noSubsetting === true || !attestation.subsettingPermitted,
  };
}

function fontStyle(face: fontkit.Font): FontFaceStyle {
  if (face["OS/2"]?.fsSelection?.oblique) return "oblique";
  if (face["OS/2"]?.fsSelection?.italic || Number(face.italicAngle) !== 0) return "italic";
  return "normal";
}

function fontWeight(face: fontkit.Font): number {
  const value = face["OS/2"]?.usWeightClass;
  return Number.isSafeInteger(value) && value >= 1 && value <= 1000 ? value : 400;
}

function fontStretch(face: fontkit.Font): number {
  const widthClass = face["OS/2"]?.usWidthClass;
  return Number.isSafeInteger(widthClass) && widthClass >= 1 && widthClass <= 9
    ? WIDTH_CLASS_PERCENT[widthClass]
    : 100;
}

function faceKey(face: Pick<ConfiguredFontFace, "familyKey" | "style" | "weight" | "stretch">): string {
  return `${face.familyKey}\u0000${face.style}\u0000${face.weight}\u0000${face.stretch}`;
}

async function readStableFontFile(
  path: string,
  root: string,
  label: string,
  maximumBytes: number,
  signal?: AbortSignal,
): Promise<Snapshot> {
  let handle: Awaited<ReturnType<typeof open>> | undefined;
  try {
    const beforeLink = await lstat(path, { bigint: true });
    if (beforeLink.isSymbolicLink() || !beforeLink.isFile()) {
      fontError(`Configured ${label} is not a regular, non-symlink file.`,
        "Remove symlinks and non-regular entries from font directories.");
    }
    const noFollow = "O_NOFOLLOW" in constants ? constants.O_NOFOLLOW : 0;
    handle = await open(path, constants.O_RDONLY | noFollow);
    const realPathBefore = await realpath(path);
    if (!underRoot(root, realPathBefore)) {
      fontError(`Configured ${label} resolves outside its font directory.`,
        "Remove path aliases from font directories.");
    }
    const before = await handle.stat({ bigint: true });
    if (!before.isFile()) {
      fontError(`Configured ${label} is not a regular file.`, "Use ordinary local font files.");
    }
    if (before.size > BigInt(maximumBytes)) {
      limitError("fontFileBytes", Number(before.size), maximumBytes);
    }
    const length = Number(before.size);
    const bytes = new Uint8Array(length);
    let offset = 0;
    while (offset < length) {
      if (signal?.aborted) throw signal.reason;
      const chunk = Math.min(1024 * 1024, length - offset);
      const { bytesRead } = await handle.read(bytes, offset, chunk, offset);
      if (bytesRead === 0) break;
      offset += bytesRead;
    }
    const extra = new Uint8Array(1);
    const probe = await handle.read(extra, 0, 1, offset);
    const after = await handle.stat({ bigint: true });
    const pathAfter = await lstat(path, { bigint: true });
    const realPathAfter = await realpath(path);
    if (pathAfter.isSymbolicLink() || !pathAfter.isFile()
      || !sameIdentity(before, after) || !sameIdentity(after, pathAfter)
      || realPathBefore !== realPathAfter || offset !== length || probe.bytesRead !== 0
      || BigInt(bytes.byteLength) !== after.size) {
      fontError(`Configured ${label} changed while it was being snapshotted.`,
        "Retry after the font deployment is stable.");
    }
    return { bytes };
  } catch (cause) {
    if (cause instanceof DocxodusExportError) throw cause;
    if (signal?.aborted) throw signal.reason;
    return fontError(`Configured ${label} could not be read safely.`,
      "Verify font-directory permissions and remove path aliases.",
      cause instanceof Error ? cause.message : undefined, cause);
  } finally {
    await handle?.close().catch(() => undefined);
  }
}

function compareNames(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

async function directoryEntries(
  path: string,
  expectedIdentity: BigIntStats,
  alreadyCounted: number,
  maximumEntries: number,
): Promise<Array<{ name: string }>> {
  const entries: Array<{ name: string }> = [];
  let directory: Awaited<ReturnType<typeof opendir>> | undefined;
  try {
    const before = await stat(path, { bigint: true });
    if (!before.isDirectory() || !sameIdentity(before, expectedIdentity)) {
      fontError("A configured font directory changed before enumeration.",
        "Retry after the font deployment is stable.");
    }
    directory = await opendir(path);
    for await (const entry of directory) {
      enforceLimit("fontDirectoryEntries", alreadyCounted + entries.length + 1, maximumEntries);
      entries.push({ name: entry.name });
    }
    const after = await stat(path, { bigint: true });
    if (!after.isDirectory() || !sameIdentity(before, after)) {
      fontError("A configured font directory changed during enumeration.",
        "Retry after the font deployment is stable.");
    }
  } catch (cause) {
    if (cause instanceof DocxodusExportError) throw cause;
    fontError("A configured font directory could not be enumerated safely.",
      "Verify directory permissions and remove unstable entries.",
      cause instanceof Error ? cause.message : undefined, cause);
  } finally {
    await directory?.close().catch(() => undefined);
  }
  return entries.sort((left, right) => compareNames(left.name, right.name));
}

export async function discoverFontCatalog(
  directories: readonly string[],
  attestations: readonly FontLicenseAttestation[],
  limits: Readonly<ExportResourceLimits>,
  signal?: AbortSignal,
): Promise<FontCatalog> {
  const directoryInputs = Object.freeze([...directories]);
  const limitSnapshot = Object.freeze({ ...limits });
  const attestationByDigest = new Map(attestations.map((entry) => [
    entry.fileSha256,
    Object.freeze({
      ...entry,
      permittedOutputs: Object.freeze([...entry.permittedOutputs]),
    }),
  ]));
  enforceLimit("fontDirectoryEntries", directoryInputs.length, limitSnapshot.fontDirectoryEntries);
  const roots: DirectorySnapshot[] = [];
  const rootIdentities = new Set<string>();
  for (let index = 0; index < directoryInputs.length; index++) {
    if (signal?.aborted) throw signal.reason;
    try {
      const root = await realpath(resolve(directoryInputs[index]));
      const identity = await stat(root, { bigint: true });
      if (!identity.isDirectory()) {
        fontError(`fontDirectory[${index}] is not a directory.`, "Provide an ordinary local directory.");
      }
      const key = `${identity.dev}:${identity.ino}`;
      if (rootIdentities.has(key)) {
        fontError(`fontDirectory[${index}] duplicates an earlier directory.`,
          "List each resolved font directory once.");
      }
      rootIdentities.add(key);
      roots.push({ path: root, identity });
    } catch (cause) {
      if (cause instanceof DocxodusExportError) throw cause;
      fontError(`fontDirectory[${index}] could not be resolved safely.`,
        "Verify the directory exists and is readable.",
        cause instanceof Error ? cause.message : undefined, cause);
    }
  }

  const faces: ConfiguredFontFace[] = [];
  const seenFacesByDigest = new Map<string, ConfiguredFontFace>();
  let entryCount = 0;
  let fileCount = 0;
  let totalBytes = 0;
  let totalExpandedBytes = 0;
  let totalCodePoints = 0;
  let discoveryIndex = 0;

  for (let directoryIndex = 0; directoryIndex < roots.length; directoryIndex++) {
    const rootSnapshot = roots[directoryIndex];
    const root = rootSnapshot.path;
    const pending: DirectorySnapshot[] = [rootSnapshot];
    const directoryFaces = new Map<string, ConfiguredFontFace>();
    while (pending.length > 0) {
      if (signal?.aborted) throw signal.reason;
      const current = pending.pop()!;
      const entries = await directoryEntries(
        current.path,
        current.identity,
        entryCount,
        limitSnapshot.fontDirectoryEntries,
      );
      const batchStart = entryCount;
      entryCount += entries.length;
      const childDirectories: DirectorySnapshot[] = [];
      for (let entryIndex = 0; entryIndex < entries.length; entryIndex++) {
        const entry = entries[entryIndex];
        if (signal?.aborted) throw signal.reason;
        const candidate = join(current.path, entry.name);
        let info: BigIntStats;
        try {
          info = await lstat(candidate, { bigint: true });
        } catch (cause) {
          // A stat failure here isn't uniformly "the deployment changed underneath us" — a
          // permission problem needs different advice than a genuine ENOENT race, and either
          // way the real errno is the one thing that tells an operator what to actually fix.
          const code = cause && typeof cause === "object" && "code" in cause ? String(cause.code) : undefined;
          const permission = code === "EACCES" || code === "EPERM";
          fontError(
            `fontDirectory[${directoryIndex}] entry ${batchStart + entryIndex} ` +
            (permission ? "could not be read." : "changed during traversal."),
            permission
              ? "Grant the render host read access to the configured font directory."
              : "Retry after the font deployment is stable.",
            code, cause,
          );
        }
        if (info.isSymbolicLink()) {
          fontError(`fontDirectory[${directoryIndex}] contains a symlink.`,
            "Remove symlinks from configured font directories.");
        }
        if (info.isDirectory()) {
          const childRealPath = await realpath(candidate).catch(() =>
            fontError(`fontDirectory[${directoryIndex}] contains an unreadable directory.`,
              "Remove unstable directory entries."));
          if (!underRoot(root, childRealPath)) {
            fontError(`fontDirectory[${directoryIndex}] contains an escaping directory alias.`,
              "Remove aliases from configured font directories.");
          }
          const childIdentity = await stat(childRealPath, { bigint: true }).catch(() =>
            fontError(`fontDirectory[${directoryIndex}] contains an unreadable directory.`,
              "Remove unstable directory entries."));
          if (!childIdentity.isDirectory() || !sameIdentity(info, childIdentity)) {
            fontError(`fontDirectory[${directoryIndex}] contains an unstable directory alias.`,
              "Retry after the font deployment is stable.");
          }
          childDirectories.push({ path: childRealPath, identity: childIdentity });
          continue;
        }
        if (!info.isFile()) {
          fontError(`fontDirectory[${directoryIndex}] contains a non-regular entry.`,
            "Keep only directories and ordinary files in configured font roots.");
        }
        const extension = extname(entry.name).toLowerCase();
        if (!FONT_EXTENSIONS.has(extension)) continue;
        fileCount++;
        enforceLimit("fontFiles", fileCount, limitSnapshot.fontFiles);
        const snapshot = await readStableFontFile(
          candidate,
          root,
          `fontDirectory[${directoryIndex}] font[${fileCount - 1}]`,
          limitSnapshot.fontFileBytes,
          signal,
        );
        totalBytes += snapshot.bytes.byteLength;
        enforceLimit("fontTotalBytes", totalBytes, limitSnapshot.fontTotalBytes);
        const decoded = decodedFormat(snapshot.bytes, extension);
        enforceLimit("fontFileBytes", decoded.expandedByteLength, limitSnapshot.fontFileBytes);
        totalExpandedBytes += decoded.expandedByteLength;
        enforceLimit("fontTotalBytes", totalExpandedBytes, limitSnapshot.fontTotalBytes);
        const fileSha256 = sha256(snapshot.bytes);
        const existing = seenFacesByDigest.get(fileSha256);
        if (existing) {
          const key = faceKey(existing);
          const conflict = directoryFaces.get(key);
          if (conflict && conflict.sha256 !== fileSha256) {
            fontError(`fontDirectory[${directoryIndex}] contains ambiguous files for one family and face.`,
              "Keep one byte identity for each family/style/weight/stretch combination.",
              `${conflict.sha256},${fileSha256}`);
          }
          directoryFaces.set(key, existing);
          continue;
        }

        let parsed: fontkit.Font | fontkit.FontCollection;
        try {
          parsed = fontkit.create(Buffer.from(
            snapshot.bytes.buffer,
            snapshot.bytes.byteOffset,
            snapshot.bytes.byteLength,
          ));
        } catch (cause) {
          fontError(`fontDirectory[${directoryIndex}] contains a malformed font.`,
            "Replace the font with a valid, bounded OpenType or webfont file.", fileSha256, cause);
        }
        let face: ConfiguredFontFace;
        try {
          if ("fonts" in parsed) {
            fontError(`fontDirectory[${directoryIndex}] contains an unsupported font collection.`,
              "Provide separate TTF, OTF, WOFF, or WOFF2 face files.", fileSha256);
          }
          if (Object.keys(parsed.variationAxes ?? {}).length > 0) {
            fontError(`fontDirectory[${directoryIndex}] contains an unsupported variable font.`,
              "Provide static face files with explicit style, weight, and stretch.", fileSha256);
          }
          const family = safeMetadata(parsed.familyName, "family name")!;
          const postscriptName = safeMetadata(parsed.postscriptName, "PostScript name", false);
          const version = safeMetadata(parsed.version, "version")!;
          const os2License = os2LicenseEvidence(parsed, fileSha256);
          const attestation = attestationByDigest.get(fileSha256);
          const attested = attestedLicenseEvidence(attestation, os2License.evidence);
          const requiresAttestation = decoded.format === "woff" || decoded.format === "woff2";
          const licenseEvidence = os2License.prohibited
            ? undefined
            : attested ?? (requiresAttestation ? undefined : os2License.evidence);
          const licenseFailure = os2License.prohibited
            ? os2License.failure
            : requiresAttestation && !attested
              ? "A WOFF/WOFF2 font requires an exact embedding-rights attestation."
              : !licenseEvidence ? os2License.failure : undefined;
          // licenseEvidence/licenseFailure above are always exactly one-or-the-other, never
          // both and never neither — every branch of os2LicenseEvidence() and the two
          // requiresAttestation checks maintain that. Express it once here rather than
          // leaving every consumer to re-derive "did licensing succeed" from two independently
          // optional fields.
          const license = licenseEvidence
            ? { ok: true as const, evidence: Object.freeze({ ...licenseEvidence }) }
            : { ok: false as const, failure: licenseFailure! };
          const codePoints: number[] = [];
          let priorCodePoint = -1;
          for (const codePoint of parsed.characterSet) {
            if (!Number.isSafeInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff
              || (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
              fontError(`fontDirectory[${directoryIndex}] contains an invalid character map.`,
                "Replace the malformed font file.", fileSha256);
            }
            if (codePoint === priorCodePoint) continue;
            if (codePoint < priorCodePoint) {
              fontError(`fontDirectory[${directoryIndex}] contains an unsorted character map.`,
                "Replace the malformed font file.", fileSha256);
            }
            codePoints.push(codePoint);
            priorCodePoint = codePoint;
          }
          totalCodePoints += codePoints.length;
          if (totalCodePoints > FONT_CATALOG_CODE_POINTS_MAX) {
            exportError(
              "resource_limit",
              "font_loading",
              `Configured font character maps exceed ${FONT_CATALOG_CODE_POINTS_MAX} entries.`,
              "Use a smaller set of static font faces.",
            );
          }
          face = Object.freeze({
            id: `font-${fileSha256}`,
            directoryIndex,
            discoveryIndex: discoveryIndex++,
            family,
            familyKey: fontFamilyKey(family),
            ...(postscriptName ? { postscriptName } : {}),
            version,
            style: fontStyle(parsed),
            weight: fontWeight(parsed),
            stretch: fontStretch(parsed),
            format: decoded.format,
            mediaType: decoded.mediaType,
            byteLength: snapshot.bytes.byteLength,
            expandedByteLength: decoded.expandedByteLength,
            sha256: fileSha256,
            bytes: snapshot.bytes,
            codePoints: Object.freeze(codePoints),
            permittedOutputs: Object.freeze([
              ...(attestation?.permittedOutputs.includes("html") ? ["html" as const] : []),
              ...(attestation?.permittedOutputs.includes("pdf") ? ["pdf" as const] : []),
              ...(!attestation && licenseEvidence ? ["html" as const, "pdf" as const] : []),
            ]),
            license,
          });
        } catch (cause) {
          if (cause instanceof DocxodusExportError) throw cause;
          fontError(`fontDirectory[${directoryIndex}] contains unreadable font metadata.`,
            "Replace the font with a valid, bounded OpenType or webfont file.", fileSha256, cause);
        }
        const key = faceKey(face);
        const conflict = directoryFaces.get(key);
        if (conflict && conflict.sha256 !== face.sha256) {
          fontError(`fontDirectory[${directoryIndex}] contains ambiguous files for one family and face.`,
            "Keep one byte identity for each family/style/weight/stretch combination.",
            `${conflict.sha256},${face.sha256}`);
        }
        directoryFaces.set(key, face);
        seenFacesByDigest.set(fileSha256, face);
        faces.push(face);
      }
      const afterProcessing = await stat(current.path, { bigint: true }).catch(() =>
        fontError(`fontDirectory[${directoryIndex}] changed during traversal.`,
          "Retry after the font deployment is stable."));
      if (!afterProcessing.isDirectory() || !sameIdentity(current.identity, afterProcessing)) {
        fontError(`fontDirectory[${directoryIndex}] changed during traversal.`,
          "Retry after the font deployment is stable.");
      }
      childDirectories.sort((left, right) => compareNames(left.path, right.path)).reverse();
      pending.push(...childDirectories);
    }
  }

  return Object.freeze({
    faces: Object.freeze(faces),
    directoryCount: roots.length,
    entryCount,
    fileCount,
    totalBytes,
    totalExpandedBytes,
  });
}
