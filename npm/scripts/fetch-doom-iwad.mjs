#!/usr/bin/env node
// Put a Freedoom IWAD in the Playwright webroot.
//
// The Doom cartridge's IWAD is not in this repository — it is 10 MB of binary
// that would never diff usefully, so the shipped pages fetch it from a pinned
// jsDelivr URL (see docs/demo/vendor/NOTICE.md). The specs must not: a browser
// test suite that depends on a CDN being up, and on a sibling repository having
// been tagged, fails for reasons that have nothing to do with the change under
// test.
//
// So the specs pass `?wad=./vendor/freedoom1.wad.gz`, same-origin inside the
// test webroot, and this script is what puts it there. It pulls Freedoom's own
// release asset — GitHub serves release assets without CORS, which is why the
// BROWSER cannot use this URL, but a build step has no such problem — verifies
// the SHA-256 recorded in the notice, and gzips it exactly as the CDN copy is
// stored, so the specs exercise the same DecompressionStream path a visitor
// does.
//
// Cached by existence: it costs one 24 MB download the first time and nothing
// afterwards, so a local `npm test` pays it once.
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { gzipSync, inflateRawSync } from 'node:zlib';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(HERE, '..', 'dist', 'wasm', 'vendor');
const OUT = join(OUT_DIR, 'freedoom1.wad.gz');

const RELEASE =
  'https://github.com/freedoom/freedoom/releases/download/v0.13.0/freedoom-0.13.0.zip';
const MEMBER = 'freedoom1.wad';
// From docs/demo/vendor/NOTICE.md. If this ever fails to match, the release
// asset changed under a fixed tag and the notice is the thing to trust.
const WAD_SHA256 = '7323bcc168c5a45ff10749b339960e98314740a734c30d4b9f3337001f9e703d';

/** Extract one member from a ZIP, without a dependency.
 *
 *  Only the two storage methods a release zip actually uses are handled:
 *  0 (stored) and 8 (deflate). Reads the central directory rather than
 *  scanning for local headers, so a member's own header lying about its sizes
 *  (which the streaming variant is allowed to do) cannot mislead it. */
function extractFromZip(zip, wanted) {
  // End of central directory: scan back from the end for its signature.
  let eocd = -1;
  for (let i = zip.length - 22; i >= 0 && i > zip.length - 66000; i--) {
    if (zip.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('not a zip file (no end-of-central-directory record)');

  const count = zip.readUInt16LE(eocd + 10);
  let p = zip.readUInt32LE(eocd + 16);

  for (let i = 0; i < count; i++) {
    if (zip.readUInt32LE(p) !== 0x02014b50) throw new Error('corrupt central directory');
    const method = zip.readUInt16LE(p + 10);
    const compressedSize = zip.readUInt32LE(p + 20);
    const nameLen = zip.readUInt16LE(p + 28);
    const extraLen = zip.readUInt16LE(p + 30);
    const commentLen = zip.readUInt16LE(p + 32);
    const localOffset = zip.readUInt32LE(p + 42);
    const name = zip.subarray(p + 46, p + 46 + nameLen).toString('latin1');

    if (name.endsWith(wanted)) {
      // The local header's variable-length fields are the ones that say where
      // this member's bytes actually start.
      if (zip.readUInt32LE(localOffset) !== 0x04034b50) throw new Error('corrupt local header');
      const localNameLen = zip.readUInt16LE(localOffset + 26);
      const localExtraLen = zip.readUInt16LE(localOffset + 28);
      const start = localOffset + 30 + localNameLen + localExtraLen;
      const raw = zip.subarray(start, start + compressedSize);
      if (method === 0) return Buffer.from(raw);
      if (method === 8) return inflateRawSync(raw);
      throw new Error(`unsupported zip compression method ${method} for ${name}`);
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error(`${wanted} not found in the archive`);
}

if (existsSync(OUT)) {
  console.log(`doom iwad: ${OUT} already present, skipping download`);
  process.exit(0);
}

console.log(`doom iwad: fetching ${RELEASE}`);
const response = await fetch(RELEASE);
if (!response.ok) throw new Error(`IWAD release fetch failed: HTTP ${response.status}`);
const zip = Buffer.from(await response.arrayBuffer());

const wad = extractFromZip(zip, MEMBER);
const digest = createHash('sha256').update(wad).digest('hex');
if (digest !== WAD_SHA256) {
  throw new Error(`IWAD digest mismatch:\n  expected ${WAD_SHA256}\n  got      ${digest}`);
}
if (wad.subarray(0, 4).toString('latin1') !== 'IWAD') {
  throw new Error('extracted file is not an IWAD');
}

mkdirSync(OUT_DIR, { recursive: true });
writeFileSync(OUT, gzipSync(wad, { level: 9 }));
console.log(
  `doom iwad: wrote ${OUT} (${wad.length} bytes → ${readFileSync(OUT).length} gzipped)`,
);
