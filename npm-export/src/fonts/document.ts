import { posix } from "node:path";
import { inflateRawSync } from "node:zlib";
import {
  fontFamilyKey,
  normalizeFontFamilyName,
  type ExportResourceLimits,
  type RenderWarning,
} from "docxodus/export-browser";
import { exportError, type RenderOutput } from "../contracts.js";
import { decodeFontFace, embeddingRefusal, faceKey, type ConfiguredFontFace } from "./discovery.js";

// Word binds each embedding to a style slot, not to the font's own OS/2 style bits.
const SLOTS = {
  embedRegular: { style: "normal", weight: 400 },
  embedBold: { style: "normal", weight: 700 },
  embedItalic: { style: "italic", weight: 400 },
  embedBoldItalic: { style: "italic", weight: 700 },
} as const;

interface ZipEntry { method: number; compressedSize: number; size: number; offset: number }

class OversizedPart extends Error {}

/**
 * The package's ZIP directory by lower-cased part name (OPC names are case-insensitive), or
 * undefined when the bytes are not a plain ZIP. A package this cannot read is the package
 * preflight's to reject; embedded fonts then simply contribute nothing.
 */
function zipDirectory(bytes: Buffer, maximumEntries: number): Map<string, ZipEntry> | undefined {
  const floor = Math.max(0, bytes.length - 22 - 0xffff);
  let end = bytes.length - 22;
  while (end >= floor && bytes.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < floor) return undefined;
  const count = bytes.readUInt16LE(end + 10);
  let cursor = bytes.readUInt32LE(end + 16);
  if (count > maximumEntries) return undefined;
  const entries = new Map<string, ZipEntry>();
  for (let index = 0; index < count; index++) {
    if (cursor + 46 > bytes.length || bytes.readUInt32LE(cursor) !== 0x02014b50) return undefined;
    const nameLength = bytes.readUInt16LE(cursor + 28);
    const name = bytes.toString("utf8", cursor + 46, cursor + 46 + nameLength).toLowerCase();
    if ((bytes.readUInt16LE(cursor + 8) & 1) === 0) {
      entries.set(name, { method: bytes.readUInt16LE(cursor + 10), compressedSize: bytes.readUInt32LE(cursor + 20),
        size: bytes.readUInt32LE(cursor + 24), offset: bytes.readUInt32LE(cursor + 42) });
    }
    cursor += 46 + nameLength + bytes.readUInt16LE(cursor + 30) + bytes.readUInt16LE(cursor + 32);
  }
  return entries;
}

function readEntry(bytes: Buffer, entry: ZipEntry, maximumBytes: number): Buffer {
  if (entry.size > maximumBytes || entry.compressedSize > maximumBytes) throw new OversizedPart();
  const header = entry.offset;
  if (header + 30 > bytes.length || bytes.readUInt32LE(header) !== 0x04034b50) throw new Error("bad local header");
  const start = header + 30 + bytes.readUInt16LE(header + 26) + bytes.readUInt16LE(header + 28);
  const data = bytes.subarray(start, start + entry.compressedSize);
  const content = entry.method === 0 ? data
    : entry.method === 8 ? inflateRawSync(data, { maxOutputLength: Math.max(1, entry.size) }) : undefined;
  if (!content || data.length !== entry.compressedSize || content.length !== entry.size) throw new Error("corrupt part");
  return content;
}

function decodeXml(value: string): string {
  return value.replace(/&(#x[0-9a-f]+|#[0-9]+|amp|lt|gt|quot|apos);/giu, (reference, entity: string) => {
    if (entity[0] !== "#") return ({ amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'" } as Record<string, string>)[entity.toLowerCase()];
    const hex = entity[1] === "x" || entity[1] === "X";
    const codePoint = Number.parseInt(entity.slice(hex ? 2 : 1), hex ? 16 : 10);
    // A reference past U+10FFFF names no character; keep it literally rather than throw.
    return codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : reference;
  });
}

/** Attributes by local name, so any namespace prefix the producer chose is accepted. */
function attributes(tag: string): Map<string, string> {
  const result = new Map<string, string>();
  for (const [, name, , value] of tag.matchAll(/([\w.:-]+)\s*=\s*(["'])(.*?)\2/gsu)) {
    result.set(name.slice(name.indexOf(":") + 1), decodeXml(value));
  }
  return result;
}

/** ECMA-376 font obfuscation: the GUID's 16 bytes, in reverse order, XOR the first 32 bytes. */
export function deobfuscateFont(bytes: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/gu, "");
  if (!/^[0-9a-f]{32}$/iu.test(hex)) throw new Error("the w:fontKey is not a GUID");
  if (bytes.length < 32) throw new Error("the part is shorter than its obfuscated header");
  const result = new Uint8Array(bytes);
  for (let index = 0; index < 32; index++) {
    result[index] ^= Number.parseInt(hex.slice(30 - (index % 16) * 2, 32 - (index % 16) * 2), 16);
  }
  return result;
}

function extensionOf(bytes: Uint8Array): string {
  const signature = Buffer.from(bytes.subarray(0, 4)).toString("latin1");
  return signature === "OTTO" ? ".otf" : signature === "wOFF" ? ".woff" : signature === "wOF2" ? ".woff2" : ".ttf";
}

/**
 * The faces a DOCX embeds through `w:fontTable` (`w:embedRegular`/`Bold`/`Italic`/`BoldItalic`),
 * de-obfuscated and decoded through the same face decoder and embedding-rights rules as every
 * other font source. A part that is missing, corrupt, keyed wrongly, oversized or not licensed for
 * the requested outputs is skipped with a `document_font_skipped` warning; it never fails export.
 */
export function documentFontFaces(
  docx: Uint8Array,
  limits: Readonly<ExportResourceLimits>,
  outputs: readonly RenderOutput[],
  warn: (warning: RenderWarning) => void,
): ConfiguredFontFace[] {
  const bytes = Buffer.from(docx.buffer, docx.byteOffset, docx.byteLength);
  let directory: Map<string, ZipEntry> | undefined;
  try { directory = zipDirectory(bytes, limits.opcEntries); } catch { return []; }
  if (!directory) return [];
  const part = (name: string, maximumBytes: number): Buffer | undefined => {
    const entry = directory.get(name.toLowerCase());
    return entry ? readEntry(bytes, entry, maximumBytes) : undefined;
  };
  const relationships = (source: string): Array<Map<string, string>> => {
    let xml: string | undefined;
    try { xml = part(posix.join(posix.dirname(source), "_rels", `${posix.basename(source)}.rels`), limits.xmlPartBytes)?.toString("utf8"); }
    catch { return []; }
    return [...(xml ?? "").matchAll(/<(?:[\w.-]+:)?Relationship\b([^<>]*)>/gu)].map(match => {
      const record = attributes(match[1]);
      const target = record.get("Target") ?? "";
      record.set("Target", posix.normalize(target.startsWith("/") ? target.slice(1) : posix.join(posix.dirname(source), target)));
      return record;
    }).filter(record => record.get("TargetMode") !== "External");
  };
  const related = (source: string, type: string) =>
    relationships(source).find(record => record.get("Type")?.endsWith(`/${type}`))?.get("Target");
  const main = related("", "officeDocument");
  const fontTable = main && related(main, "fontTable");
  let tableXml: string | undefined;
  try { tableXml = fontTable && part(fontTable, limits.xmlPartBytes)?.toString("utf8"); } catch { tableXml = undefined; }
  if (!fontTable || !tableXml) return [];
  const targets = new Map(relationships(fontTable).map(record => [record.get("Id"), record.get("Target")]));
  const faces: ConfiguredFontFace[] = [];
  const keys = new Set<string>();
  // Like directory discovery, every embedding read counts against fontFiles and every byte read
  // against fontTotalBytes, used or not, so a hostile font table cannot buy unbounded work.
  let attempts = 0;
  let bytesRead = 0;
  const limit = (name: "fontFiles" | "fontTotalBytes", actual: number): void => {
    if (actual > limits[name]) exportError("resource_limit", "font_loading", `${name} limit exceeded (${actual} > ${limits[name]}).`,
      "Use a document with fewer or smaller embedded fonts.");
  };
  // w:font elements do not nest, so each one's children lie between its tag and the next: one
  // linear pass with no closing-tag search a hostile part could make quadratic.
  const fonts = [...tableXml.matchAll(/<(?:[\w.-]+:)?font\b([^<>]*)>/gu)];
  for (const [index, font] of fonts.entries()) {
    const family = normalizeFontFamilyName(attributes(font[1]).get("name") ?? "");
    if (!family || family.length > 256 || /[\u0000-\u001f\u007f]/u.test(family)) continue;
    const body = tableXml.slice(font.index + font[0].length, fonts[index + 1]?.index ?? tableXml.length);
    for (const [, slot, slotAttributes] of body.matchAll(/<(?:[\w.-]+:)?(embed(?:Regular|Bold|Italic|BoldItalic))\b([^<>]*)>/gu)) {
      const skip = (reason: string): void => warn({ code: "document_font_skipped", severity: "warning",
        phase: "font_loading", message: `Embedded font ${family} (${slot}): ${reason}.`, resource: family,
        remediation: "Re-embed the font from Word, or configure a font directory that supplies the family." });
      const record = attributes(slotAttributes);
      const target = targets.get(record.get("id"));
      limit("fontFiles", ++attempts);
      let stored: Buffer | undefined;
      try { stored = target ? part(target, limits.fontFileBytes) : undefined; }
      catch (error) {
        skip(error instanceof OversizedPart ? "its font part exceeds fontFileBytes" : "its font part is corrupt");
        continue;
      }
      if (!stored) { skip("its font part is missing"); continue; }
      limit("fontTotalBytes", bytesRead += stored.length);
      try {
        const key = record.get("fontKey");
        const program = key ? deobfuscateFont(stored, key) : new Uint8Array(stored);
        const decoded = decodeFontFace(program, extensionOf(program), 0, faces.length);
        const face: ConfiguredFontFace = Object.freeze({ ...decoded, ...SLOTS[slot as keyof typeof SLOTS],
          id: `document-${faces.length}-${decoded.sha256}`, source: "document" as const, family,
          familyKey: fontFamilyKey(family) });
        const refusal = embeddingRefusal(face, outputs);
        if (refusal) { skip(refusal); continue; }
        if (keys.has(faceKey(face))) { skip("an earlier embedding already supplies this style"); continue; }
        keys.add(faceKey(face));
        faces.push(face);
      } catch {
        skip("it could not be decoded (a corrupt part or a wrong w:fontKey)");
      }
    }
  }
  return faces;
}
