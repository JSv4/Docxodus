import { deflateSync, inflateSync } from 'node:zlib';

export interface RgbaImage {
  width: number;
  height: number;
  data: Uint8Array;
}

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const MAXIMUM_PNG_INPUT_BYTES = 64 * 1024 * 1024;
/** Shared with the PDF raster contract, which enforces the same ceiling on pdftoppm output. */
export const MAXIMUM_PNG_PIXELS = 4_000_000;

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
}

/** Decode the non-interlaced, 8-bit PNG formats emitted by Chromium and pdftoppm. */
export function decodePng(input: Uint8Array): RgbaImage {
  const bytes = Buffer.from(input);
  if (bytes.length < PNG_SIGNATURE.length || bytes.length > MAXIMUM_PNG_INPUT_BYTES ||
      !bytes.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    throw new Error(`PNG input must be a valid file no larger than ${MAXIMUM_PNG_INPUT_BYTES} bytes`);
  }

  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = -1;
  let interlace = -1;
  let palette: Buffer | undefined;
  let transparency: Buffer | undefined;
  const idat: Buffer[] = [];
  let sawHeader = false;
  let sawEnd = false;
  let sawImageData = false;
  let endedImageData = false;
  let chunkCount = 0;

  for (let offset = PNG_SIGNATURE.length; offset + 12 <= bytes.length;) {
    chunkCount++;
    if (chunkCount > 100_000) throw new Error('PNG contains too many chunks');
    const length = bytes.readUInt32BE(offset);
    const type = bytes.toString('ascii', offset + 4, offset + 8);
    const end = offset + 12 + length;
    if (!Number.isSafeInteger(end) || end > bytes.length) {
      throw new Error(`Truncated PNG ${type} chunk`);
    }
    const data = bytes.subarray(offset + 8, offset + 8 + length);
    const expectedCrc = bytes.readUInt32BE(offset + 8 + length);
    const actualCrc = crc32(bytes.subarray(offset + 4, offset + 8 + length));
    if (actualCrc !== expectedCrc) throw new Error(`PNG ${type} chunk failed its CRC check`);
    if (type === 'IHDR') {
      if (sawHeader || offset !== PNG_SIGNATURE.length || data.length !== 13) {
        throw new Error('PNG must contain exactly one leading 13-byte IHDR chunk');
      }
      sawHeader = true;
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      bitDepth = data[8];
      colorType = data[9];
      if (data[10] !== 0 || data[11] !== 0) {
        throw new Error('PNG uses an unsupported compression or filter method');
      }
      interlace = data[12];
    } else if (type === 'PLTE') {
      if (palette || sawImageData) throw new Error('PNG PLTE chunk is duplicate or out of order');
      palette = data;
    } else if (type === 'tRNS') {
      if (transparency || sawImageData) throw new Error('PNG tRNS chunk is duplicate or out of order');
      transparency = data;
    } else if (type === 'IDAT') {
      if (!sawHeader || sawEnd || endedImageData) throw new Error('PNG IDAT chunk is out of order');
      sawImageData = true;
      idat.push(data);
    } else if (type === 'IEND') {
      if (data.length !== 0) throw new Error('PNG IEND chunk must be empty');
      sawEnd = true;
      if (end !== bytes.length) throw new Error('PNG has trailing bytes after IEND');
      break;
    } else {
      if (sawImageData) endedImageData = true;
      if (type[0] === type[0]?.toUpperCase()) {
        throw new Error(`PNG contains unsupported critical chunk ${type}`);
      }
    }
    offset = end;
  }

  if (!sawHeader || !sawEnd || !width || !height || !idat.length) {
    throw new Error('PNG is missing IHDR, IDAT, or IEND data');
  }
  const pixels = width * height;
  if (!Number.isSafeInteger(pixels) || pixels > MAXIMUM_PNG_PIXELS) {
    throw new Error(`PNG dimensions exceed the ${MAXIMUM_PNG_PIXELS}-pixel limit`);
  }
  if (bitDepth !== 8 || interlace !== 0) {
    throw new Error(`Unsupported PNG format: bitDepth=${bitDepth}, interlace=${interlace}`);
  }

  const channelsByType: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };
  const channels = channelsByType[colorType];
  if (!channels) throw new Error(`Unsupported PNG color type ${colorType}`);
  if (colorType === 3 && !palette) throw new Error('Indexed PNG is missing its palette');
  if (palette && (palette.length === 0 || palette.length > 768 || palette.length % 3 !== 0)) {
    throw new Error('PNG palette has an invalid length');
  }
  if (transparency && colorType === 3 && transparency.length > (palette?.length ?? 0) / 3) {
    throw new Error('PNG transparency table exceeds its palette');
  }

  const stride = width * channels;
  const expectedFilteredBytes = (stride + 1) * height;
  if (!Number.isSafeInteger(expectedFilteredBytes)) throw new Error('PNG payload size is unsafe');
  const filtered = inflateSync(Buffer.concat(idat), { maxOutputLength: expectedFilteredBytes });
  if (filtered.length !== expectedFilteredBytes) {
    throw new Error(`Unexpected PNG payload length ${filtered.length}`);
  }

  const raw = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const sourceOffset = y * (stride + 1);
    const targetOffset = y * stride;
    const filter = filtered[sourceOffset];
    for (let x = 0; x < stride; x++) {
      const value = filtered[sourceOffset + 1 + x];
      const left = x >= channels ? raw[targetOffset + x - channels] : 0;
      const up = y > 0 ? raw[targetOffset + x - stride] : 0;
      const upLeft = y > 0 && x >= channels
        ? raw[targetOffset + x - stride - channels]
        : 0;
      let decoded: number;
      switch (filter) {
        case 0: decoded = value; break;
        case 1: decoded = value + left; break;
        case 2: decoded = value + up; break;
        case 3: decoded = value + Math.floor((left + up) / 2); break;
        case 4: decoded = value + paeth(left, up, upLeft); break;
        default: throw new Error(`Unsupported PNG row filter ${filter}`);
      }
      raw[targetOffset + x] = decoded & 0xff;
    }
  }

  const rgba = new Uint8Array(width * height * 4);
  for (let pixel = 0; pixel < width * height; pixel++) {
    const source = pixel * channels;
    const target = pixel * 4;
    if (colorType === 0) {
      rgba[target] = rgba[target + 1] = rgba[target + 2] = raw[source];
      rgba[target + 3] = 255;
    } else if (colorType === 2) {
      rgba[target] = raw[source];
      rgba[target + 1] = raw[source + 1];
      rgba[target + 2] = raw[source + 2];
      rgba[target + 3] = 255;
    } else if (colorType === 3) {
      const index = raw[source];
      rgba[target] = palette![index * 3] ?? 0;
      rgba[target + 1] = palette![index * 3 + 1] ?? 0;
      rgba[target + 2] = palette![index * 3 + 2] ?? 0;
      rgba[target + 3] = transparency?.[index] ?? 255;
    } else if (colorType === 4) {
      rgba[target] = rgba[target + 1] = rgba[target + 2] = raw[source];
      rgba[target + 3] = raw[source + 1];
    } else {
      rgba[target] = raw[source];
      rgba[target + 1] = raw[source + 1];
      rgba[target + 2] = raw[source + 2];
      rgba[target + 3] = raw[source + 3];
    }
  }

  return { width, height, data: rgba };
}

let crcTable: Uint32Array | undefined;

function crc32(bytes: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Buffer {
  const typeBytes = Buffer.from(type, 'ascii');
  const chunk = Buffer.alloc(data.length + 12);
  chunk.writeUInt32BE(data.length, 0);
  typeBytes.copy(chunk, 4);
  Buffer.from(data).copy(chunk, 8);
  chunk.writeUInt32BE(crc32(chunk.subarray(4, 8 + data.length)), 8 + data.length);
  return chunk;
}

/** Encode an RGBA image as a deterministic, non-interlaced PNG. */
export function encodePng(image: RgbaImage): Buffer {
  const pixels = image.width * image.height;
  if (!Number.isSafeInteger(pixels) || pixels < 1 || pixels > MAXIMUM_PNG_PIXELS) {
    throw new Error(`RGBA dimensions exceed the ${MAXIMUM_PNG_PIXELS}-pixel limit`);
  }
  if (image.data.length !== pixels * 4) {
    throw new Error('RGBA buffer length does not match its dimensions');
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(image.width, 0);
  header.writeUInt32BE(image.height, 4);
  header[8] = 8;
  header[9] = 6;
  header[10] = 0;
  header[11] = 0;
  header[12] = 0;

  const rows = Buffer.alloc((image.width * 4 + 1) * image.height);
  for (let y = 0; y < image.height; y++) {
    const row = y * (image.width * 4 + 1);
    rows[row] = 0;
    Buffer.from(image.data).copy(rows, row + 1, y * image.width * 4, (y + 1) * image.width * 4);
  }

  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk('IHDR', header),
    pngChunk('IDAT', deflateSync(rows, { level: 9 })),
    pngChunk('IEND', new Uint8Array()),
  ]);
}
