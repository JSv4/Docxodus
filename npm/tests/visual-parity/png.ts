import { deflateSync, inflateSync } from 'node:zlib';

export interface RgbaImage {
  width: number;
  height: number;
  data: Uint8Array;
}

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

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
  if (bytes.length < PNG_SIGNATURE.length ||
      !bytes.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    throw new Error('Not a PNG file');
  }

  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = -1;
  let interlace = -1;
  let palette: Buffer | undefined;
  let transparency: Buffer | undefined;
  const idat: Buffer[] = [];

  for (let offset = PNG_SIGNATURE.length; offset + 12 <= bytes.length;) {
    const length = bytes.readUInt32BE(offset);
    const type = bytes.toString('ascii', offset + 4, offset + 8);
    const data = bytes.subarray(offset + 8, offset + 8 + length);
    if (offset + 12 + length > bytes.length) throw new Error(`Truncated PNG ${type} chunk`);
    if (type === 'IHDR') {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      bitDepth = data[8];
      colorType = data[9];
      interlace = data[12];
    } else if (type === 'PLTE') {
      palette = data;
    } else if (type === 'tRNS') {
      transparency = data;
    } else if (type === 'IDAT') {
      idat.push(data);
    } else if (type === 'IEND') {
      break;
    }
    offset += length + 12;
  }

  if (!width || !height || !idat.length) throw new Error('PNG is missing IHDR or IDAT data');
  if (bitDepth !== 8 || interlace !== 0) {
    throw new Error(`Unsupported PNG format: bitDepth=${bitDepth}, interlace=${interlace}`);
  }

  const channelsByType: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };
  const channels = channelsByType[colorType];
  if (!channels) throw new Error(`Unsupported PNG color type ${colorType}`);
  if (colorType === 3 && !palette) throw new Error('Indexed PNG is missing its palette');

  const stride = width * channels;
  const filtered = inflateSync(Buffer.concat(idat));
  if (filtered.length !== (stride + 1) * height) {
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
  if (image.data.length !== image.width * image.height * 4) {
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
