// Collection table offsets are absolute within the collection. Rebuild a standalone
// SFNT for Chromium, including checksums; a collection's DSIG cannot sign new bytes.
// https://learn.microsoft.com/en-us/typography/opentype/spec/otff
export interface SfntTable {
  tag: string;
  offset: number;
  length: number;
}

function check(condition: boolean, message: string): asserts condition {
  if (!condition) throw new Error(`Invalid font collection: ${message}.`);
}

export function collectionOffsets(header: Buffer, fileLength: number): number[] {
  check(header.length >= 12 && header.toString("ascii", 0, 4) === "ttcf", "truncated header");
  const version = header.readUInt32BE(4);
  check(version === 0x10000 || version === 0x20000, "unsupported version");
  const count = header.readUInt32BE(8);
  const headerLength = 12 + count * 4 + (version === 0x20000 ? 12 : 0);
  check(count > 0 && headerLength <= header.length, "truncated face directory");
  return Array.from({ length: count }, (_, index) => {
    const offset = header.readUInt32BE(12 + index * 4);
    check(offset >= headerLength && offset + 12 <= fileLength, "face offset outside file");
    return offset;
  });
}

export function sfntTables(directory: Buffer, fileLength: number): SfntTable[] {
  check(directory.length >= 12, "truncated SFNT header");
  const signature = directory.readUInt32BE(0);
  check([0x10000, 0x4f54544f, 0x74727565].includes(signature), "unsupported SFNT format");
  const count = directory.readUInt16BE(4);
  check(count > 0 && count <= 4095 && 12 + count * 16 <= directory.length, "truncated table directory");
  const tags = new Set<string>();
  return Array.from({ length: count }, (_, index) => {
    const record = 12 + index * 16;
    const tag = directory.toString("latin1", record, record + 4);
    const offset = directory.readUInt32BE(record + 8);
    const length = directory.readUInt32BE(record + 12);
    check(!tags.has(tag) && offset + length <= fileLength, "duplicate or out-of-range table");
    tags.add(tag);
    return { tag, offset, length };
  });
}

function checksum(bytes: Buffer): number {
  let sum = 0;
  for (let offset = 0; offset < bytes.length; offset += 4) {
    let word = 0;
    for (let byte = 0; byte < 4; byte++) word = (word << 8) | (bytes[offset + byte] ?? 0);
    sum = (sum + (word >>> 0)) >>> 0;
  }
  return sum;
}

export function packSfnt(
  signature: number,
  tables: readonly { tag: string; bytes: Buffer }[],
  maximumBytes: number,
): Buffer {
  const kept = tables.filter(table => table.tag !== "DSIG").sort((a, b) => a.tag < b.tag ? -1 : 1);
  const length = 12 + kept.length * 16 + kept.reduce((sum, table) => sum + Math.ceil(table.bytes.length / 4) * 4, 0);
  check(length <= maximumBytes, "expanded face exceeds fontFileBytes limit");
  const result = Buffer.alloc(length);
  result.writeUInt32BE(signature, 0);
  result.writeUInt16BE(kept.length, 4);
  const selector = Math.floor(Math.log2(kept.length));
  const searchRange = 2 ** selector * 16;
  result.writeUInt16BE(searchRange, 6);
  result.writeUInt16BE(selector, 8);
  result.writeUInt16BE(kept.length * 16 - searchRange, 10);
  let offset = 12 + kept.length * 16;
  let headOffset: number | undefined;
  for (const [index, table] of kept.entries()) {
    const record = 12 + index * 16;
    result.write(table.tag, record, 4, "latin1");
    result.writeUInt32BE(offset, record + 8);
    result.writeUInt32BE(table.bytes.length, record + 12);
    table.bytes.copy(result, offset);
    if (table.tag === "head") {
      check(table.bytes.length >= 12, "truncated head table");
      headOffset = offset;
      result.writeUInt32BE(0, offset + 8);
    }
    result.writeUInt32BE(checksum(result.subarray(offset, offset + table.bytes.length)), record + 4);
    offset += Math.ceil(table.bytes.length / 4) * 4;
  }
  if (headOffset !== undefined) result.writeUInt32BE((0xb1b0afba - checksum(result)) >>> 0, headOffset + 8);
  return result;
}

export function* collectionFaces(bytes: Uint8Array, maximumBytes: number, maximumFaces: number): Generator<Buffer> {
  const source = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const offsets = collectionOffsets(source, source.length);
  check(offsets.length <= maximumFaces, "face count exceeds fontFiles limit");
  for (const offset of offsets) {
    const directory = source.subarray(offset);
    const tables = sfntTables(directory, source.length);
    yield packSfnt(directory.readUInt32BE(0), tables.map(table => ({
      tag: table.tag,
      bytes: source.subarray(table.offset, table.offset + table.length),
    })), maximumBytes);
  }
}
