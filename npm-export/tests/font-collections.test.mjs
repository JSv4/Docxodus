import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { DEFAULT_EXPORT_RESOURCE_LIMITS } from "../dist/index.js";
import { discoverFontCatalog, resolveCatalogRequests } from "../dist/fonts/index.js";

const collection = await readFile(new URL("./fixtures/docxodus-test-collection.ttc", import.meta.url));
const hash = bytes => createHash("sha256").update(bytes).digest("hex");

async function discover(t, bytes = collection, extension = "ttc", limits = {}, attestations = []) {
  const folder = await mkdtemp(join(tmpdir(), "docxodus-collection-"));
  t.after(() => rm(folder, { recursive: true, force: true }));
  await writeFile(join(folder, `faces.${extension}`), bytes);
  return discoverFontCatalog([folder], attestations, { ...DEFAULT_EXPORT_RESOURCE_LIMITS, ...limits });
}

for (const extension of ["ttc", "otc"]) {
  test(`extracts independently usable faces from a .${extension} collection`, async t => {
    const catalog = await discover(t, collection, extension);
    assert.equal(catalog.fileCount, 1);
    assert.deepEqual(catalog.faces.map(face => face.family), ["Docxodus Policy Test", "Docxodus Metric Test"]);
    for (const face of catalog.faces) {
      assert.equal(hash(face.bytes), face.sha256);
      assert.equal(face.license.ok, true);
      assert.equal(face.format, "ttf");
      const bytes = Buffer.from(face.bytes);
      let checksum = 0;
      for (let offset = 0; offset < bytes.length; offset += 4) checksum = (checksum + bytes.readUInt32BE(offset)) >>> 0;
      assert.equal(checksum, 0xb1b0afba);
      const result = resolveCatalogRequests(catalog, [{
        id: "face", familyStack: [face.family], familyKinds: ["named"],
        style: "normal", weight: 400, stretch: 100, sampleCodePoints: [65, 90],
      }], ["html", "pdf"]);
      assert.equal(result.outcomes[0].status, "resolved");
      assert.equal(result.faces[0].sha256, face.sha256);
    }
  });
}

test("rejects truncated collection directories and escaping table ranges", async t => {
  const truncated = Buffer.from(collection.subarray(0, 18));
  await assert.rejects(discover(t, truncated), /collection|truncated/i);
  const corrupt = Buffer.from(collection);
  const firstFace = corrupt.readUInt32BE(12);
  corrupt.writeUInt32BE(0xfffffff0, firstFace + 12 + 8);
  await assert.rejects(discover(t, corrupt), /collection|table|range/i);
});

test("preserves both TrueType and CFF outlines in a mixed OpenType collection", async t => {
  const mixed = await readFile(new URL("./fixtures/docxodus-test-collection.otc", import.meta.url));
  const catalog = await discover(t, mixed, "otc");
  assert.deepEqual(catalog.faces.map(face => face.format), ["ttf", "otf"]);
  assert.equal(catalog.faces[1].family, "Docxodus CFF Test");
  assert.equal(catalog.faces[1].license.evidence.kind, "editable");
  assert.ok([32, 65, 90].every(point => catalog.faces[1].codePoints.includes(point)));
});

test("bounds collection expansion and retains per-face embedding restrictions", async t => {
  await assert.rejects(discover(t, collection, "ttc", { fontTotalBytes: collection.length }), /fontTotalBytes/);
  const restricted = Buffer.from(collection);
  const firstFace = restricted.readUInt32BE(12);
  for (let index = 0; index < restricted.readUInt16BE(firstFace + 4); index++) {
    const record = firstFace + 12 + index * 16;
    if (restricted.toString("ascii", record, record + 4) === "OS/2") {
      restricted.writeUInt16BE(2, restricted.readUInt32BE(record + 8) + 8);
    }
  }
  const catalog = await discover(t, restricted);
  assert.equal(catalog.faces[0].license.ok, false);
  assert.throws(() => resolveCatalogRequests(catalog, [{
    id: "restricted", familyStack: [catalog.faces[0].family], familyKinds: ["named"],
    style: "normal", weight: 400, stretch: 100, sampleCodePoints: [65],
  }]), /forbids font embedding/);
});

test("collection attestations bind source bytes while output evidence binds each extracted face", async t => {
  const attestation = fileSha256 => ({ schemaVersion: 1, usage: "standalone-document-font-embedding",
    fileSha256, embeddingPermitted: true, permittedOutputs: ["pdf"], subsettingPermitted: true,
    basis: "Original geometric test font; PDF-only test permission" });
  const catalog = await discover(t, collection, "ttc", {}, [attestation(hash(collection))]);
  for (const face of catalog.faces) {
    assert.equal(face.license.evidence.kind, "attested");
    assert.notEqual(face.sha256, hash(collection));
    const request = { id: "face", familyStack: [face.family], familyKinds: ["named"],
      style: "normal", weight: 400, stretch: 100, sampleCodePoints: [65] };
    assert.equal(resolveCatalogRequests(catalog, [request], ["pdf"]).faces[0].sha256, face.sha256);
    assert.throws(() => resolveCatalogRequests(catalog, [request], ["html"]), /not attested for every requested output/);
  }
  const extracted = await discover(t, collection, "ttc", {}, [attestation(catalog.faces[0].sha256)]);
  assert.equal(extracted.faces[0].license.evidence.kind, "installable");
});
