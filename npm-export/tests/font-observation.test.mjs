import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, before, test } from "node:test";
import { chromium } from "playwright-core";
import { renderDocxArtifacts } from "../dist/index.js";
import { canonicalJson } from "../dist/canonical.js";
import { observeFonts } from "../dist/font-observation.js";
import { fontObservationDocx, MISSING_FAMILY, CONFIGURED_FAMILY } from "./font-observation-fixture.mjs";

const digest = bytes => createHash('sha256').update(bytes).digest('hex');
let browser, root, options;
before(async () => {
  browser = await chromium.launch({ chromiumSandbox: true,
    ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) });
  root = await mkdtemp(join(tmpdir(), 'docxodus-font-observation-'));
  const directory = join(root, 'fonts'); await mkdir(directory);
  const bytes = await readFile(new URL('../../docs/demo/fonts/docxodus-canvas-mono.woff2', import.meta.url));
  await writeFile(join(directory, 'fixture.woff2'), bytes);
  options = { browser, reviewProfile: 'final', commentProfile: 'hidden', fontDirectories: [directory],
    fontLicenseAttestations: [{ schemaVersion: 1, usage: 'standalone-document-font-embedding',
      fileSha256: digest(bytes), embeddingPermitted: true, permittedOutputs: ['html', 'pdf'],
      subsettingPermitted: true, basis: 'Docxodus test fixture license' }] };
});
after(async () => { await browser?.close(); if (root) await rm(root, { recursive: true, force: true }); });

test('original missing, configured and unused declarations retain their resolver diagnostics', async () => {
  for (const setting of [{ family: MISSING_FAMILY }, {}, { unused: 'style' }, { unused: 'table' }]) {
    const result = await renderDocxArtifacts(fontObservationDocx(setting), { ...options, outputs: ['html'] });
    const missing = setting.family === MISSING_FAMILY;
    assert.equal(result.renderReport.fonts.length, 1);
    assert.equal(result.renderReport.fonts[0].status, missing ? 'missing' : 'resolved');
    assert.equal(result.renderReport.fonts[0].verified, !missing);
    assert.equal(result.warnings.some(w => w.code === 'font_unavailable'), missing);
  }
});

for (const output of ['html', 'pdf']) test(`${output} reports actual fallback faces without attesting a file`, async () => {
  const source = fontObservationDocx({ family: MISSING_FAMILY });
  const result = await renderDocxArtifacts(source, { ...options, outputs: [output] });
  const report = result.renderReport, observation = report.fontObservation;
  assert.equal(observation?.status, 'observed');
  assert.equal(observation.media, output === 'pdf' ? 'print' : 'screen');
  assert.equal(observation.fileIdentity, 'unavailable');
  assert.match(report.fontObservationDigest, /^[a-f0-9]{64}$/);
  const uses = observation.uses.filter(use => use.requestId === report.fonts[0].requestId);
  assert.ok(uses.length > 0);
  assert.ok(uses.every(use => use.anchorId && /^[a-f0-9]{64}$/.test(use.textDigest)));
  assert.ok(uses.every(use => use.faces.length > 0 && use.faces.every(face =>
    face.family !== MISSING_FAMILY && face.glyphCount > 0 && !('fileSha256' in face))));
  assert.equal(report.fonts[0].status, 'missing');
  assert.equal(report.fonts[0].verified, false);
  assert.equal(report.source.rawPackageBytesDigest, digest(source));
});

test('configured ASCII keeps source attribution after Word SVG painting', async () => {
  const result = await renderDocxArtifacts(fontObservationDocx(), { ...options, outputs: ['html', 'pdf'] });
  assert.match(result.html, /<svg\b/);
  const observation = result.renderReport.fontObservation;
  assert.equal(observation?.status, 'observed');
  assert.equal(observation.media, 'print');
  const uses = observation.uses.filter(use => use.requestId === result.renderReport.fonts[0].requestId);
  assert.ok(uses.length > 0);
  assert.ok(uses.every(use => use.faces.some(face => face.isCustomFont && face.family === CONFIGURED_FAMILY)));
});

test('mixed-script and per-glyph fallback retain every actual face for one text request', async () => {
  const result = await renderDocxArtifacts(fontObservationDocx({ text: 'A漢😀' }), { ...options, outputs: ['html', 'pdf'] });
  const font = result.renderReport.fonts[0];
  assert.equal(font.glyphCoverage, 'partial');
  assert.equal(font.verified, false);
  const observation = result.renderReport.fontObservation;
  assert.equal(observation?.status, 'observed');
  const uses = observation.uses.filter(use => use.requestId === font.requestId);
  assert.ok(uses.some(use => use.faces.length >= 2));
  assert.ok(uses.some(use => use.faces.some(face => face.isCustomFont)));
  assert.ok(uses.some(use => use.faces.some(face => !face.isCustomFont)));
});

test('repeated exports bind stable observations to the report and PageMap', async () => {
  const source = fontObservationDocx();
  const first = await renderDocxArtifacts(source, { ...options, outputs: ['html'] });
  const second = await renderDocxArtifacts(source, { ...options, outputs: ['html'] });
  const a = first.renderReport, b = second.renderReport;
  assert.deepEqual(a.fontObservation, b.fontObservation);
  assert.equal(a.fontObservationDigest, digest(`docxodus:font-observation:v1\0${canonicalJson(a.fontObservation)}`));
  assert.equal(a.fontObservationDigest, b.fontObservationDigest);
  assert.equal(a.environment.rendererFingerprint, b.environment.rendererFingerprint);
  assert.equal(first.pageMap.rendererFingerprint, a.environment.rendererFingerprint);
  assert.equal(a.bindings.pageMapDigest, digest(canonicalJson(first.pageMap)));
  assert.equal(a.fontObservation.uses.some(use => 'text' in use), false);
});

for (const outputs of [['html'], ['pdf'], ['html', 'pdf']]) {
  test(`${outputs.join('+')} keeps every paragraph when font observations exceed the budget`, async () => {
    const paragraphs = Array.from({ length: 120 }, (_, i) => `Orchard entry ${String(i).padStart(3, '0')}.`);
    const source = fontObservationDocx({ text: paragraphs });
    const result = await renderDocxArtifacts(source, { ...options, outputs, strictFonts: true,
      limits: { renderDiagnostics: 128 } });
    const report = result.renderReport;
    assert.deepEqual(report.fontObservation, { status: 'unavailable', reason: 'diagnostic_limit' });
    assert.equal(report.fontObservationDigest,
      digest(`docxodus:font-observation:v1\0${canonicalJson(report.fontObservation)}`));
    assert.equal(report.environment.rendererFingerprint, result.pageMap.rendererFingerprint);
    assert.equal(report.fonts[0].verified, true);
    assert.equal(report.fonts[0].status, 'resolved');
    assert.equal(result.warnings.filter(w => w.code === 'font_observation_limit').length, 1);
    if (outputs.length === 2) {
      const complete = await renderDocxArtifacts(source, { ...options, outputs, strictFonts: true });
      assert.equal(complete.renderReport.fontObservation.status, 'observed');
      assert.equal(complete.pageCount, result.pageCount);
      assert.deepEqual(complete.renderReport.fonts, report.fonts);
      assert.deepEqual(complete.renderReport.fontIdentity, report.fontIdentity);
      assert.equal(complete.warnings.some(w => w.code === 'font_observation_limit'), false);
    }
    const assertParagraphs = text => assert.deepEqual(
      text.replace(/\s/g, '').match(/Orchardentry\d{3}\./g), paragraphs.map(p => p.replace(/\s/g, '')));
    if (result.html) {
      const context = await browser.newContext({ javaScriptEnabled: false });
      try {
        const page = await context.newPage();
        await page.setContent(result.html);
        assertParagraphs((await page.locator('.page-box').allTextContents()).join(''));
      } finally { await context.close(); }
    }
    if (result.pdf) {
      const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs');
      const task = getDocument({ data: new Uint8Array(result.pdf), verbosity: 0 });
      try {
        const pdf = await task.promise, text = [];
        assert.equal(pdf.numPages, result.pageCount);
        for (let n = 1; n <= pdf.numPages; n++) {
          text.push((await (await pdf.getPage(n)).getTextContent()).items.map(item => item.str ?? '').join(' '));
        }
        assertParagraphs(text.join(' '));
      } finally { await task.destroy(); }
    }
  });
}

test('an observation budget does not bypass strict font verification', async () => {
  const source = fontObservationDocx({ family: MISSING_FAMILY, text: Array(120).fill('Orchard entry.') });
  await assert.rejects(renderDocxArtifacts(source, { ...options, outputs: ['pdf'], strictFonts: true,
    limits: { renderDiagnostics: 128 } }), { code: 'resource_policy_failure', phase: 'font_loading' });
});

test('text-node queries separate nested faces, SVG and unattributed generated text', async () => {
  const context = await browser.newContext({ javaScriptEnabled: false });
  try {
    const page = await context.newPage();
    const bytes = await readFile(new URL('../../docs/demo/fonts/docxodus-canvas-mono.woff2', import.meta.url));
    await page.setContent(`<style>@font-face{font-family:Probe;src:url(data:font/woff2;base64,${bytes.toString('base64')})}</style>
      <div data-page-number="1" data-source-anchor-id="p:body:probe" data-page-fragment-id="probe-0">
      <p data-docx-font-request="outer" style="font-family:serif">Outer<span data-docx-font-request="inner" style="font-family:Probe">Inner</span>Tail</p>
      <svg data-docx-font-request="inner"><text y="20" style="font-family:Probe">Svg</text></svg>
      <p>Generated</p><p style="display:none" data-docx-font-request="outer">Hidden</p></div>`);
    await page.evaluate(() => document.fonts.ready);
    const observation = await observeFonts(context, page, { fonts: [{ requestId: 'outer' }, { requestId: 'inner' }],
      resources: [], warnings: [], unsupportedContent: [], options: { policy: { limits: { renderDiagnostics: 100 } } } }, 'screen');
    const uses = observation.uses;
    assert.equal(uses.length, 5);
    for (const text of ['Outer', 'Tail']) {
      const use = uses.find(u => u.textDigest === digest(text));
      assert.equal(use.requestId, 'outer');
      assert.ok(use.faces.every(f => !f.isCustomFont));
    }
    for (const text of ['Inner', 'Svg']) {
      const use = uses.find(u => u.textDigest === digest(text));
      assert.equal(use.requestId, 'inner');
      assert.deepEqual(use.faces.map(f => f.family), [CONFIGURED_FAMILY]);
    }
    assert.equal(uses.find(u => u.textDigest === digest('Generated')).requestId, undefined);
    assert.ok(uses.every(u => u.anchorId === 'p:body:probe' && u.fragmentId === 'probe-0' && u.pageNumber === 1));
  } finally { await context.close(); }
});
