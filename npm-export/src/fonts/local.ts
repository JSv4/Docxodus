import type { BigIntStats } from "node:fs";
import { lstat, opendir, realpath, stat } from "node:fs/promises";
import { homedir } from "node:os";
import { extname, join, posix, win32 } from "node:path";
import * as fontkit from "fontkit";
import { fontFamilyKey, type ExportResourceLimits, type RenderWarning } from "docxodus/export-browser";
import { DocxodusExportError, exportError, type RenderOutput } from "../contracts.js";
import { decodeFontFace, faceKey, fontMetadata, openFontFile, readStableFontFile, sameIdentity, underRoot,
  type ConfiguredFontFace } from "./discovery.js";
import { collectionOffsets, packSfnt, sfntTables } from "./sfnt.js";

const EXTENSIONS = new Set([".ttf", ".otf", ".ttc", ".otc"]);
const METADATA_TABLES = new Set(["name", "OS/2", "head", "post"]);
const METADATA_BYTES = 1024 * 1024;

export interface LocalFontOptions {
  directories?: readonly string[];
  strict?: boolean;
  warn?: (warning: RenderWarning) => void;
}

export function localFontDirectories(
  platform = process.platform,
  home = homedir(),
  environment: NodeJS.ProcessEnv = process.env,
): string[] {
  if (platform === "win32") return [
    ...(environment.WINDIR ? [win32.join(environment.WINDIR, "Fonts")] : []),
    ...(environment.LOCALAPPDATA ? [
      win32.join(environment.LOCALAPPDATA, "Microsoft", "Windows", "Fonts"),
      win32.join(environment.LOCALAPPDATA, "Microsoft", "FontCache", "4", "CloudFonts"),
    ] : []),
  ];
  if (platform === "darwin") return [
    posix.join(home, "Library/Fonts"), "/Library/Fonts", "/System/Library/Fonts",
    ...["/Applications", posix.join(home, "Applications")].flatMap(applications =>
      ["DFonts", "Fonts"].map(folder => posix.join(applications, "Microsoft Word.app/Contents/Resources", folder))),
  ];
  if (platform === "linux") return [posix.join(home, ".local/share/fonts"), posix.join(home, ".fonts"),
    "/usr/local/share/fonts", "/usr/share/fonts"];
  return [];
}

export interface LocalFontFace extends ReturnType<typeof fontMetadata> {
  path: string;
  root: string;
  identity: BigIntStats;
  offset?: number;
}

// Read only the SFNT directory and metadata tables, never outlines or cmap. The
// independent metadata bound applies even to files too large to embed later.
async function metadata(path: string, root: string, signal?: AbortSignal): Promise<LocalFontFace[]> {
  const handle = await openFontFile(path);
  try {
    const before = await handle.stat({ bigint: true });
    if (await realpath(path) !== path || !underRoot(root, path)) throw new Error("unsafe font path");
    if (!before.isFile() || before.size > BigInt(Number.MAX_SAFE_INTEGER)) throw new Error("not a regular font");
    const fileLength = Number(before.size);
    let totalRead = 0;
    const read = async (offset: number, length: number): Promise<Buffer> => {
      if (signal?.aborted) throw signal.reason;
      totalRead += length;
      if (offset < 0 || length < 0 || offset + length > fileLength || totalRead > METADATA_BYTES) {
        throw new Error("font metadata is truncated or too large");
      }
      const bytes = Buffer.alloc(length);
      let done = 0;
      while (done < length) {
        const { bytesRead } = await handle.read(bytes, done, length - done, offset + done);
        if (!bytesRead) throw new Error("truncated font metadata");
        done += bytesRead;
      }
      return bytes;
    };
    const header = await read(0, 12);
    const collection = header.toString("ascii", 0, 4) === "ttcf";
    const offsets = collection
      ? collectionOffsets(await read(0, 12 + header.readUInt32BE(8) * 4 + (header.readUInt32BE(4) === 0x20000 ? 12 : 0)), fileLength)
      : [0];
    const faces: LocalFontFace[] = [];
    for (const offset of offsets) {
      const faceHeader = await read(offset, 12);
      const directory = await read(offset, 12 + faceHeader.readUInt16BE(4) * 16);
      const tables = [];
      for (const table of sfntTables(directory, fileLength)) {
        if (METADATA_TABLES.has(table.tag)) tables.push({ tag: table.tag, bytes: await read(table.offset, table.length) });
      }
      const parsed = fontkit.create(packSfnt(directory.readUInt32BE(0), tables, METADATA_BYTES));
      if ("fonts" in parsed) throw new Error("unexpected nested collection");
      faces.push({ ...fontMetadata(parsed), path, root, identity: before, ...(collection ? { offset } : {}) });
    }
    if (!sameIdentity(before, await handle.stat({ bigint: true }))
      || !sameIdentity(before, await lstat(path, { bigint: true })) || await realpath(path) !== path) {
      throw new Error("font changed while reading metadata");
    }
    return faces;
  } finally { await handle.close(); }
}

export class LocalFontIndex {
  readonly faces: LocalFontFace[] = [];
  private readonly diagnostics = new Set<string>();
  private readonly warned = new Set<string>();

  constructor(private readonly options: LocalFontOptions) {}

  warning(family: string, reason: string): void {
    const message = `Local font ${family}: ${reason}.`;
    if (this.options.strict) exportError("resource_policy_failure", "font_loading", message,
      "Install an unambiguous font whose embedding rights allow the requested output, or disable strictFonts.");
    if (this.warned.has(message)) return;
    this.warned.add(message);
    this.options.warn?.({ code: "local_font_skipped", severity: "warning", phase: "font_loading",
      message, resource: family,
      remediation: "Install a readable, unambiguous font with appropriate embedding rights or configure an explicit font directory." });
  }

  async prepare(limits: Readonly<ExportResourceLimits>, signal?: AbortSignal): Promise<void> {
    const roots: string[] = [];
    for (const path of this.options.directories ?? localFontDirectories()) {
      try { const root = await realpath(path); if (!roots.includes(root)) roots.push(root); }
      catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "ENOENT") this.diagnostics.add("skipped an unreadable font location (family unavailable)");
      }
    }
    const visited = new Set<string>();
    const pending = [...roots].reverse();
    let entries = 0;
    while (pending.length) {
      if (signal?.aborted) throw signal.reason;
      const path = pending.pop()!;
      let target: string;
      let info: BigIntStats;
      try {
        target = await realpath(path);
        if (!roots.some(root => underRoot(root, target))) {
          this.diagnostics.add("skipped an unsafe symlink (family unavailable)"); continue;
        }
        if (visited.has(target)) continue;
        visited.add(target);
        info = await lstat(target, { bigint: true });
      } catch {
        this.diagnostics.add("skipped an unreadable file or unsafe symlink (family unavailable)"); continue;
      }
      if (info.isDirectory()) {
        const children: string[] = [];
        try {
          const directory = await opendir(target);
          for await (const entry of directory) {
            if (signal?.aborted) throw signal.reason;
            if (++entries > limits.fontDirectoryEntries) exportError("resource_limit", "font_loading",
              "Local font index exceeds fontDirectoryEntries limit.", "Reduce the number of installed font directory entries.");
            children.push(join(target, entry.name));
          }
        } catch (error) {
          if (signal?.aborted) throw signal.reason;
          if (error instanceof DocxodusExportError) throw error;
          this.diagnostics.add("skipped an unreadable directory (family unavailable)");
        }
        pending.push(...children.sort().reverse());
      } else if (info.isFile() && EXTENSIONS.has(extname(target).toLowerCase())) {
        try { this.faces.push(...await metadata(target, roots.find(root => underRoot(root, target))!, signal)); }
        catch {
          if (signal?.aborted) throw signal.reason;
          this.diagnostics.add("skipped an unreadable font or malformed name table (family unavailable)");
        }
      }
    }
  }

  candidates(family: string): LocalFontFace[] {
    const faces = this.faces.filter(face => face.familyKey === fontFamilyKey(family));
    if (!faces.length) for (const reason of this.diagnostics) this.warning(family, reason);
    return faces;
  }

  async load(
    entry: LocalFontFace,
    limits: Readonly<ExportResourceLimits>,
    outputs: readonly RenderOutput[],
    signal?: AbortSignal,
  ): Promise<ConfiguredFontFace | undefined> {
    let face: ConfiguredFontFace;
    try {
      if (!sameIdentity(entry.identity, await stat(entry.path, { bigint: true }))) throw new Error("changed font");
      const snapshot = await readStableFontFile(entry.path, entry.root, "local font", limits.fontFileBytes, signal);
      let bytes: Buffer = Buffer.from(snapshot.bytes);
      let extension = extname(entry.path).toLowerCase();
      if (entry.offset !== undefined) {
        const directory = bytes.subarray(entry.offset);
        bytes = packSfnt(directory.readUInt32BE(0), sfntTables(directory, bytes.length).map(table => ({
          tag: table.tag, bytes: bytes.subarray(table.offset, table.offset + table.length),
        })), limits.fontFileBytes);
        extension = bytes.toString("ascii", 0, 4) === "OTTO" ? ".otf" : ".ttf";
      }
      face = decodeFontFace(bytes, extension, Number.MAX_SAFE_INTEGER, 0);
      if (faceKey(face) !== faceKey(entry) || face.version !== entry.version || face.postscriptName !== entry.postscriptName) {
        throw new Error("changed metadata");
      }
    } catch (error) {
      if (signal?.aborted) throw signal.reason;
      this.warning(entry.family, error instanceof DocxodusExportError
        ? error.message.replace(/\.$/, "") : "file could not be read safely or changed after indexing");
      return undefined;
    }
    const reason = !face.license.ok ? face.license.failure
      : outputs.some(output => !face.permittedOutputs.includes(output)) ? "embedding rights do not allow the requested output"
      : outputs.includes("html") && face.license.evidence.kind === "previewPrint" ? "preview-and-print rights do not allow a reusable HTML font"
      : outputs.includes("pdf") && face.license.evidence.noSubsetting ? "embedding rights forbid PDF subsetting" : undefined;
    if (reason) { this.warning(entry.family, reason.replace(/\.$/, "")); return undefined; }
    return Object.freeze({ ...face, source: "local" });
  }
}
