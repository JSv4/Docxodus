import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { DEFAULT_EXPORT_RESOURCE_LIMITS, renderDocxArtifacts } from "../dist/index.js";
import { createNodeFontRuntime, documentFontFaces } from "../dist/fonts/index.js";
import { storedZip } from "./mixed-section-fixture.mjs";

// The issue's fixture: "Packaged Probe" in word/fonts/font1.odttf, linked by w:embedRegular.
const fixture = await readFile(new URL("./fixtures/embedded-font.docx", import.meta.url));
const KEY = "{6F1D2A10-3B4C-4D5E-8F60-718293A4B5C6}";
// SHA-256 of the de-obfuscated program, computed independently with fontTools.
const PROBE_SHA256 = "c1465fe5c320df8a9cd79c246448b5d6d89b6f2ceea12842d457469a0001d88c";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const limits = DEFAULT_EXPORT_RESOURCE_LIMITS;
const outputs = ["html", "pdf"];
const sha256 = bytes => createHash("sha256").update(bytes).digest("hex");

function extract(docx, options = {}) {
  const warnings = [];
  const faces = documentFontFaces(docx, { ...limits, ...options.limits }, options.outputs ?? outputs,
    warning => warnings.push(warning));
  return { faces, warnings };
}

const probe = Buffer.from(extract(fixture).faces[0]?.bytes ?? []);

/** Obfuscates the way Word does: the key GUID's bytes, reversed, XOR the first 32 bytes. */
function obfuscate(bytes, key = KEY) {
  const guid = Buffer.from(key.replace(/[{}-]/g, ""), "hex").reverse();
  const out = Buffer.from(bytes);
  for (let i = 0; i < 32; i++) out[i] ^= guid[i % 16];
  return out;
}

/** A distinct but equally valid program: only head.checkSumAdjustment differs. */
function variant(bytes, n) {
  const out = Buffer.from(bytes);
  for (let i = 0; i < out.readUInt16BE(4); i++) {
    const record = 12 + i * 16;
    if (out.toString("latin1", record, record + 4) === "head") out.writeUInt32BE(n, out.readUInt32BE(record + 8) + 8);
  }
  return out;
}

function withFsType(bytes, value) {
  const out = Buffer.from(bytes);
  for (let i = 0; i < out.readUInt16BE(4); i++) {
    const record = 12 + i * 16;
    if (out.toString("latin1", record, record + 4) === "OS/2") out.writeUInt16BE(value, out.readUInt32BE(record + 8) + 8);
  }
  return out;
}

/** A package whose "Packaged Probe" entry embeds `slots`: [element, part bytes, fontKey?]. */
function packageWith(slots, family = "Packaged Probe") {
  const xml = text => Buffer.from(text, "utf8");
  const embeds = slots.map(([element, , key], i) =>
    `<w:${element} r:id="f${i}"${key ? ` w:fontKey="${key}"` : ""}/>`).join("");
  return storedZip([
    { name: "[Content_Types].xml", data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
      + "</Types>") },
    { name: "_rels/.rels", data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
      + `<Relationship Id="d" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`) },
    { name: "word/_rels/document.xml.rels", data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
      + `<Relationship Id="t" Type="${R}/fontTable" Target="fontTable.xml"/></Relationships>`) },
    { name: "word/document.xml", data: xml(`<w:document xmlns:w="${W}"><w:body><w:p><w:r>`
      + '<w:rPr><w:rFonts w:ascii="Packaged Probe" w:hAnsi="Packaged Probe"/></w:rPr><w:t>TIDE TABLES</w:t></w:r></w:p>'
      + "</w:body></w:document>") },
    { name: "word/fontTable.xml", data: xml(`<w:fonts xmlns:w="${W}" xmlns:r="${R}">`
      + `<w:font w:name="${family}">${embeds}</w:font></w:fonts>`) },
    { name: "word/_rels/fontTable.xml.rels", data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
      + slots.map((_, i) => `<Relationship Id="f${i}" Type="${R}/font" Target="fonts/font${i}.odttf"/>`).join("")
      + "</Relationships>") },
    ...slots.map(([, data], i) => ({ name: `word/fonts/font${i}.odttf`, data })),
  ]);
}

function request(id, family, style = "normal", weight = 400) {
  return { id, familyStack: [family, "serif"], familyKinds: ["named", "generic"],
    style, weight, stretch: 100, sampleCodePoints: [65, 84] };
}

function resolve(documentFaces, requests, options = {}) {
  return createNodeFontRuntime(options.directories ?? [], [], { ...limits, ...options.limits }, outputs,
    options.local, documentFaces).resolver({ schemaVersion: 1, requests }, new AbortController().signal);
}

test("de-obfuscates the issue fixture's embedded regular face with its w:fontKey", () => {
  const { faces, warnings } = extract(fixture);
  assert.deepEqual(warnings, []);
  assert.equal(faces.length, 1);
  const [face] = faces;
  assert.equal(face.sha256, PROBE_SHA256);
  assert.deepEqual([face.source, face.family, face.style, face.weight, face.postscriptName],
    ["document", "Packaged Probe", "normal", 400, "PackagedProbe-Regular"]);
});

test("an out-of-range character reference in the font table is kept literally, never an error", () => {
  const embed = [["embedRegular", obfuscate(probe), KEY]];
  // Control: a valid reference decodes.
  assert.equal(extract(packageWith(embed, "Packaged&#x20;Probe")).faces[0]?.family, "Packaged Probe");
  // Past U+10FFFF there is no character; String.fromCodePoint would throw a RangeError.
  const { faces } = extract(packageWith(embed, "Probe &#99999999; &#x110000;"));
  assert.equal(faces[0]?.family, "Probe &#99999999; &#x110000;");
});

test("a wrong key or a truncated part is skipped with a warning, never an error", () => {
  for (const [label, data] of [["wrong key", obfuscate(probe, "{00000000-0000-0000-0000-000000000001}")],
    ["truncated", obfuscate(probe).subarray(0, 200)]]) {
    const { faces, warnings } = extract(packageWith([["embedRegular", data, KEY]]));
    assert.deepEqual(faces, [], label);
    assert.equal(warnings.length, 1, label);
    assert.equal(warnings[0].code, "document_font_skipped");
    assert.match(warnings[0].message, /Packaged Probe \(embedRegular\): .*wrong w:fontKey/);
  }
  // Control: the same package with the right key and the whole part yields the face.
  assert.equal(extract(packageWith([["embedRegular", obfuscate(probe), KEY]])).faces[0].sha256, PROBE_SHA256);
});

test("an embedding without w:fontKey is read as a plain font program, as Word does", () => {
  assert.equal(extract(packageWith([["embedRegular", probe]])).faces[0].sha256, PROBE_SHA256);
});

test("each embed slot serves its own style and weight", async () => {
  const slots = ["embedRegular", "embedBold", "embedItalic", "embedBoldItalic"];
  const programs = slots.map((_, i) => variant(probe, i + 1));
  const { faces } = extract(packageWith(slots.map((slot, i) => [slot, obfuscate(programs[i]), KEY])));
  const response = await resolve(faces, [request("r", "Packaged Probe"), request("b", "Packaged Probe", "normal", 700),
    request("i", "Packaged Probe", "italic"), request("bi", "Packaged Probe", "italic", 700)]);
  const shaById = new Map(response.faces.map(face => [face.id, face.sha256]));
  assert.deepEqual(response.outcomes.map(outcome => [outcome.status, outcome.faceMatch, shaById.get(outcome.faceId)]),
    programs.map(program => ["resolved", "exact", sha256(program)]));
  assert.ok(response.faces.every(face => face.source === "document"));
});

test("the font's own embedding bits decide, by the rules every other source uses", () => {
  const embed = (fsType, options) => extract(packageWith([["embedRegular", obfuscate(withFsType(probe, fsType)), KEY]]), options);
  const restricted = embed(0x0002);
  assert.deepEqual(restricted.faces, []);
  assert.match(restricted.warnings[0].message, /forbids font embedding/);
  const noSubsetting = embed(0x0100);
  assert.deepEqual(noSubsetting.faces, []);
  assert.match(noSubsetting.warnings[0].message, /forbid PDF subsetting/);
  // Control: the same no-subsetting face is usable when no PDF is requested.
  assert.equal(embed(0x0100, { outputs: ["html"] }).faces.length, 1);
});

test("embedded faces count against the font file, count and total-byte limits", async t => {
  const tooLarge = extract(fixture, { limits: { fontFileBytes: probe.length - 1 } });
  assert.deepEqual(tooLarge.faces, []);
  assert.match(tooLarge.warnings[0].message, /exceeds fontFileBytes/);
  assert.equal(extract(fixture, { limits: { fontFileBytes: probe.length } }).faces.length, 1);
  // Every embedding read counts, as every directory file does, so the font table cannot buy unbounded work.
  const two = packageWith([["embedRegular", obfuscate(probe), KEY], ["embedBold", obfuscate(variant(probe, 3)), KEY]]);
  assert.throws(() => extract(two, { limits: { fontFiles: 1 } }), error => error.code === "resource_limit" && /fontFiles/.test(error.message));
  assert.throws(() => extract(two, { limits: { fontTotalBytes: 2 * probe.length - 1 } }),
    error => error.code === "resource_limit" && /fontTotalBytes/.test(error.message));
  assert.equal(extract(two, { limits: { fontFiles: 2, fontTotalBytes: 2 * probe.length } }).faces.length, 2);

  const root = await mkdtemp(join(tmpdir(), "docxodus-document-fonts-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await writeFile(join(root, "other.ttf"), variant(probe, 9));
  const { faces } = extract(fixture);
  const requests = [request("r", "Packaged Probe")];
  await assert.rejects(resolve(faces, requests, { directories: [root], limits: { fontFiles: 1 } }),
    error => error.code === "resource_limit" && /fontFiles/.test(error.message));
  await assert.rejects(resolve(faces, requests, { limits: { fontTotalBytes: probe.length - 1 } }),
    error => error.code === "resource_limit" && /fontTotalBytes/.test(error.message));
  // Controls: one more file, or one more byte, of headroom resolves.
  assert.equal((await resolve(faces, requests, { directories: [root], limits: { fontFiles: 2 } })).outcomes[0].status, "resolved");
  assert.equal((await resolve(faces, requests, { limits: { fontTotalBytes: probe.length } })).outcomes[0].status, "resolved");
});

test("a configured or installed family takes precedence over the document's embedded faces", async t => {
  const root = await mkdtemp(join(tmpdir(), "docxodus-document-fonts-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await writeFile(join(root, "installed.ttf"), variant(probe, 7));
  const { faces } = extract(fixture);
  const requests = [request("r", "Packaged Probe")];
  const source = async options => {
    const response = await resolve(faces, requests, options);
    return [response.faces[0].source, response.faces[0].sha256];
  };
  assert.deepEqual(await source({ directories: [root] }), [undefined, sha256(variant(probe, 7))]);
  assert.deepEqual(await source({ local: { directories: [root] } }), ["local", sha256(variant(probe, 7))]);
  // Control: without either, the embedded face serves the family.
  assert.deepEqual(await source({}), ["document", PROBE_SHA256]);
});

test("an embedded family stops the search before any local substitute is read", async t => {
  const root = await mkdtemp(join(tmpdir(), "docxodus-document-fonts-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  // A restricted Carlito, Calibri's substitute: reading it under strict local fonts would throw.
  const carlito = await readFile(new URL("./fixtures/synthetic-carlito.ttf", import.meta.url));
  await writeFile(join(root, "carlito.ttf"), withFsType(carlito, 0x0002));
  const faces = extract(fixture).faces.map(face => ({ ...face, family: "Calibri", familyKey: "calibri" }));
  const response = await resolve(faces, [request("r", "Calibri")], { local: { directories: [root], strict: true } });
  assert.equal(response.faces[0].source, "document");
});

const render = (source, outputs) => renderDocxArtifacts(source, {
  outputs, reviewProfile: "final", commentProfile: "hidden", timeoutMs: 120_000,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { browserExecutablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
});

test("export prints the issue fixture in its embedded face and reports it as a document source", { timeout: 180_000 }, async () => {
  const result = await render(fixture, ["pdf"]);
  const font = result.renderReport.fonts.find(entry => entry.requestedFamily === "Packaged Probe");
  assert.deepEqual([font.status, font.source, font.fileSha256, font.verified],
    ["resolved", "document", PROBE_SHA256, true]);
  assert.ok(Buffer.from(result.pdf).includes("PackagedProbe-Regular"), "the PDF embeds the document's face");
  assert.ok(!result.renderReport.warnings.some(warning => warning.code === "document_font_skipped"));
});

test("export of a wrongly keyed embedding completes with a warning and the browser fallback", { timeout: 180_000 }, async () => {
  const result = await render(packageWith([["embedRegular", obfuscate(probe, "{00000000-0000-0000-0000-000000000001}"), KEY]]), ["pdf"]);
  assert.ok(result.renderReport.warnings.some(warning => warning.code === "document_font_skipped"));
  const font = result.renderReport.fonts.find(entry => entry.requestedFamily === "Packaged Probe");
  assert.equal(font.source, "browser");
  assert.ok(!Buffer.from(result.pdf).includes("PackagedProbe-Regular"));
});
