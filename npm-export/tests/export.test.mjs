import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import {
  mkdir,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { after, before, describe, test } from "node:test";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";
import { getDocument, OPS } from "pdfjs-dist/legacy/build/pdf.mjs";
import { PDFDocument } from "pdf-lib";
import {
  convertDocxToPdf,
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DocxodusExportError,
  renderDocxArtifacts,
  renderDocxFile,
} from "../dist/index.js";
import {
  exposePrintReadinessBindings,
  validatePrintDocument,
} from "../dist/browser-session.js";
import { canonicalJson } from "../dist/canonical.js";
import { discoverFontCatalog, pathFreeCatalogManifest } from "../dist/fonts/index.js";
import {
  generateFontProbeDocx,
  generateLongFootnoteDocx,
  generateMixedSectionDocx,
} from "./mixed-section-fixture.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const packageRoot = dirname(here);
const repositoryRoot = dirname(packageRoot);
const fixtures = join(repositoryRoot, "TestFiles");
const artifacts = join(packageRoot, "test-artifacts");
const successArtifacts = join(artifacts, "success");
const failureArtifacts = join(artifacts, "failure");
const configuredFontDirectory = join(artifacts, "configured-fonts");
const configuredFontFixture = join(repositoryRoot, "docs", "demo", "fonts", "docxodus-canvas-mono.woff2");
const fontFixtureRoot = join(packageRoot, "tests", "fixtures");
const carlitoFontDirectory = join(artifacts, "carlito-fonts");
const narrowCarlitoFontDirectory = join(artifacts, "carlito-narrow-fonts");
const wideCarlitoFontDirectory = join(artifacts, "carlito-wide-fonts");
const loadFailureFontDirectory = join(artifacts, "load-failure-fonts");
const emptyFontDirectory = join(artifacts, "empty-fonts");
const ambiguousFontDirectory = join(artifacts, "ambiguous-fonts");
const fontScenarioArtifacts = join(successArtifacts, "font-scenarios");
const cliFontAttestationPath = join(artifacts, "font-license-attestations.json");
const hostFontPolicyPath = join(artifacts, "host-font-policy.json");
const ambiguousHostFontPolicyPath = join(artifacts, "ambiguous-host-font-policy.json");
const cliEntry = join(packageRoot, "dist", "cli.js");
const hostEntry = join(packageRoot, "dist", "host.js");
const baseOptions = Object.freeze({
  reviewProfile: "final",
  commentProfile: "hidden",
  timeoutMs: 120_000,
});
const tinyPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

let browser;
let ambiguousFontAttestations;

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function domainDigest(domain, value) {
  return digest(Buffer.from(`${domain}\0${value}`));
}

async function writeJson(path, value) {
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`);
}

async function captureStandaloneScreenshot(html, path) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 960 } });
  try {
    const page = await context.newPage();
    await page.setContent(html, { waitUntil: "load" });
    await page.evaluate(async () => { await document.fonts.ready; });
    await page.screenshot({ path, fullPage: true });
  } finally {
    await context.close();
  }
}

async function writeFontScenario(id, source, result) {
  const directory = join(fontScenarioArtifacts, id);
  await mkdir(directory, { recursive: true });
  await writeFile(join(directory, "source.docx"), source);
  await writeFile(join(directory, "standalone.html"), result.html);
  await writeFile(join(directory, "output.pdf"), result.pdf);
  await writeJson(join(directory, "page-map.json"), result.pageMap);
  await writeJson(join(directory, "render-report.json"), result.renderReport);
  await writeJson(join(directory, "error.json"), { ok: true, error: null });
  await captureStandaloneScreenshot(result.html, join(directory, "screenshot.png"));
  return {
    id,
    pageCount: result.pageCount,
    rendererFingerprint: result.rendererFingerprint,
    sourceDigest: digest(source),
    htmlDigest: result.renderReport.bindings.htmlDigest,
    pdfDigest: result.renderReport.bindings.pdfDigest,
    pageMapDigest: result.renderReport.bindings.pageMapDigest,
    fontIdentity: result.renderReport.fontIdentity,
    fonts: result.renderReport.fonts,
    pageGeometry: {
      pages: result.pageMap.pages,
      fragments: result.pageMap.fragments.map((fragment) => ({
        fragmentId: fragment.fragmentId,
        anchorId: fragment.anchorId,
        pageNumber: fragment.pageNumber,
        geometry: fragment.geometry,
      })),
    },
    warningCodes: result.warnings.map(({ code }) => code),
    artifacts: {
      source: `${id}/source.docx`,
      html: `${id}/standalone.html`,
      pdf: `${id}/output.pdf`,
      pageMap: `${id}/page-map.json`,
      report: `${id}/render-report.json`,
      screenshot: `${id}/screenshot.png`,
    },
  };
}

async function currentRuntimeDirectories() {
  return new Set((await readdir(tmpdir())).filter((name) =>
    name.startsWith("docxodus-export-")
    || name.startsWith("playwright-artifacts-")
    || name.startsWith("playwright_chromiumdev_profile-")));
}

async function inspectPdf(bytes) {
  const geometryDocument = await PDFDocument.load(new Uint8Array(bytes), {
    ignoreEncryption: false,
    updateMetadata: false,
    throwOnInvalidObject: true,
  });
  const geometryPages = geometryDocument.getPages();
  const loading = getDocument({
    data: new Uint8Array(bytes),
    isEvalSupported: false,
    useSystemFonts: true,
  });
  const pdf = await loading.promise;
  const pages = [];
  let text = "";
  let links = 0;
  let vectorPaths = 0;
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const pageText = content.items.map((item) => "str" in item ? item.str : "").join(" ");
    const annotations = await page.getAnnotations();
    const operations = await page.getOperatorList();
    const pageLinks = annotations.filter((annotation) => annotation.subtype === "Link");
    const pagePaths = operations.fnArray.filter((operation) => operation === OPS.constructPath).length;
    text += `${pageText}\n`;
    links += pageLinks.length;
    vectorPaths += pagePaths;
    pages.push({
      pageNumber,
      text: pageText,
      textRuns: content.items
        .filter((item) => "str" in item && item.str.trim().length > 0)
        .map((item) => ({
          text: item.str,
          transform: [...item.transform],
          width: item.width,
          height: item.height,
        })),
      mediaBox: geometryPages[pageNumber - 1].getMediaBox(),
      cropBox: geometryPages[pageNumber - 1].getCropBox(),
      links: pageLinks.map(({ url, unsafeUrl, dest }) => ({ url, unsafeUrl, dest })),
      constructPathOperations: pagePaths,
    });
  }
  const markInfo = await pdf.getMarkInfo();
  await loading.destroy();
  return {
    pageCount: pages.length,
    searchableText: text.trim(),
    linkAnnotations: links,
    vectorPathOperations: vectorPaths,
    marked: markInfo?.Marked === true,
    pages,
  };
}

function assertPdfBoxes(inspection, expectedPages, tolerance = 0.5) {
  assert.equal(inspection.pageCount, expectedPages.length);
  for (let index = 0; index < expectedPages.length; index++) {
    const expected = expectedPages[index];
    const actual = inspection.pages[index];
    for (const [name, box] of [["MediaBox", actual.mediaBox], ["CropBox", actual.cropBox]]) {
      assert.ok(Math.abs(box.x) <= tolerance, `page ${index + 1} ${name} x=${box.x}`);
      assert.ok(Math.abs(box.y) <= tolerance, `page ${index + 1} ${name} y=${box.y}`);
      assert.ok(
        Math.abs(box.width - expected.width) <= tolerance,
        `page ${index + 1} ${name} width ${box.width} != ${expected.width}`,
      );
      assert.ok(
        Math.abs(box.height - expected.height) <= tolerance,
        `page ${index + 1} ${name} height ${box.height} != ${expected.height}`,
      );
    }
  }
}

function assertEquivalentPdfBoxes(actual, expected, tolerance = 0.5) {
  assert.equal(actual.pageCount, expected.pageCount);
  for (let index = 0; index < actual.pages.length; index++) {
    for (const boxName of ["mediaBox", "cropBox"]) {
      for (const coordinate of ["x", "y", "width", "height"]) {
        const actualValue = actual.pages[index][boxName][coordinate];
        const expectedValue = expected.pages[index][boxName][coordinate];
        assert.ok(
          Math.abs(actualValue - expectedValue) <= tolerance,
          `page ${index + 1} ${boxName}.${coordinate} ${actualValue} != ${expectedValue}`,
        );
      }
    }
  }
}

function assertEquivalentPdfTextGeometry(actual, expected, tolerance = 0.25) {
  assert.equal(actual.pageCount, expected.pageCount);
  for (let pageIndex = 0; pageIndex < actual.pages.length; pageIndex++) {
    const actualRuns = actual.pages[pageIndex].textRuns;
    const expectedRuns = expected.pages[pageIndex].textRuns;
    assert.equal(
      actualRuns.length,
      expectedRuns.length,
      `page ${pageIndex + 1} text-run count changed`,
    );
    for (let runIndex = 0; runIndex < actualRuns.length; runIndex++) {
      const actualRun = actualRuns[runIndex];
      const expectedRun = expectedRuns[runIndex];
      assert.equal(
        actualRun.text,
        expectedRun.text,
        `page ${pageIndex + 1} text run ${runIndex + 1} changed`,
      );
      for (let component = 0; component < actualRun.transform.length; component++) {
        assert.ok(
          Math.abs(actualRun.transform[component] - expectedRun.transform[component]) <= tolerance,
          `page ${pageIndex + 1} text run ${runIndex + 1} transform[${component}] changed`,
        );
      }
      for (const dimension of ["width", "height"]) {
        assert.ok(
          Math.abs(actualRun[dimension] - expectedRun[dimension]) <= tolerance,
          `page ${pageIndex + 1} text run ${runIndex + 1} ${dimension} changed`,
        );
      }
    }
  }
}

async function writeViewer() {
  const fontScenarioLinks = ["exact", "substitution", "missing", "load-failure", "metric-difference"]
    .map((id) => `<li><strong>${id}</strong>: `
      + `<a href="success/font-scenarios/${id}/standalone.html">HTML</a> · `
      + `<a href="success/font-scenarios/${id}/output.pdf">PDF</a> · `
      + `<a href="success/font-scenarios/${id}/page-map.json">PageMap</a> · `
      + `<a href="success/font-scenarios/${id}/render-report.json">report</a> · `
      + `<a href="success/font-scenarios/${id}/screenshot.png">screenshot</a> · `
      + `<a href="success/font-scenarios/${id}/error.json">failure (if any)</a></li>`)
    .join("\n");
  const viewer = `<!doctype html><meta charset="utf-8"><title>Docxodus #439/#440/#441/#442 test artifacts</title>
<style>body{font:15px system-ui;margin:2rem;max-width:70rem}li{margin:.55rem 0}iframe,object,img{width:100%;min-height:48rem;border:1px solid #bbb}</style>
<h1>Docxodus #439/#440/#441/#442 Node/PDF export evidence</h1>
<p>Generated by the end-to-end Node, CLI, Chromium, and PDF-parser test suite.</p>
<ul>
<li><a href="success/generated.pdf">Generated searchable PDF</a></li>
<li><a href="success/standalone.html">Standalone offline HTML</a></li>
<li><a href="success/page-map.json">PageMap</a></li>
<li><a href="success/render-report.json">Complete render report</a></li>
<li><a href="success/print-readiness.json">Print-readiness phase and paginator evidence</a></li>
<li><a href="success/configured-font.docx">Configured-font source DOCX</a></li>
<li><a href="success/configured-font.pdf">Configured-font PDF</a></li>
<li><a href="success/configured-font.html">Configured-font standalone HTML</a></li>
<li><a href="success/configured-font-report.json">Path-free configured-font report</a></li>
<li><a href="success/font-scenarios/comparison-manifest.json">#442 scenario manifest and fingerprint comparison</a></li>
<li><a href="success/font-scenarios/catalog-manifest.json">#442 canonical path-free font catalog manifest</a></li>
<li><a href="success/pdf-inspection.json">PDF text/tag inspection</a></li>
<li><a href="success/request-log.json">Offline reopen request log</a></li>
<li><a href="success/offline-reopen.png">Offline reopen screenshot</a></li>
<li><a href="success/hyperlinks.pdf">Hyperlink preservation PDF</a></li>
<li><a href="success/chart-vector.pdf">Vector-chart preservation PDF</a></li>
<li><a href="success/cli.pdf">CLI PDF</a></li>
<li><a href="success/cli-configured-font.pdf">CLI repeatable font-directory PDF</a></li>
<li><a href="success/cli-configured-font-report.json">CLI configured-font report</a></li>
<li><a href="success/mixed-sections.docx">Mixed-section source DOCX</a></li>
<li><a href="success/mixed-sections.html">Mixed-section standalone HTML</a></li>
<li><a href="success/mixed-sections.pdf">Mixed-section PDF</a></li>
<li><a href="success/mixed-sections-scaled-viewer.pdf">PDF reprinted from an 80% screen view</a></li>
<li><a href="success/mixed-sections-page-map.json">Mixed-section PageMap</a></li>
<li><a href="success/mixed-sections-render-report.json">Mixed-section render report</a></li>
<li><a href="success/mixed-sections-inspection.json">Mixed-section PDF geometry/text inspection</a></li>
<li><a href="success/mixed-sections-scaled-viewer.png">80% viewer screenshot</a></li>
<li><a href="failure/pdf-limit-render-report.json">Structured PDF-limit failure report</a></li>
</ul>
<h2>#442 font scenario matrix</h2><ul>${fontScenarioLinks}</ul>
<p><strong>Metric baseline:</strong> <a href="success/font-scenarios/metric-baseline/standalone.html">HTML</a> · <a href="success/font-scenarios/metric-baseline/output.pdf">PDF</a> · <a href="success/font-scenarios/metric-baseline/page-map.json">PageMap</a> · <a href="success/font-scenarios/metric-baseline/render-report.json">report</a> · <a href="success/font-scenarios/metric-baseline/screenshot.png">screenshot</a></p>
<h2>PDF preview</h2><object data="success/generated.pdf" type="application/pdf"></object>
<h2>Mixed-section PDF preview</h2><object data="success/mixed-sections.pdf" type="application/pdf"></object>
<h2>Offline HTML preview</h2><iframe sandbox src="success/standalone.html"></iframe>`;
  await writeFile(join(artifacts, "view-artifacts.html"), viewer);
}

function framedRequest(value) {
  const payload = Buffer.from(JSON.stringify(value));
  const header = Buffer.alloc(4);
  header.writeUInt32BE(payload.byteLength);
  return Buffer.concat([header, payload]);
}

function framedProtocolRequest(value, blobs = []) {
  return Buffer.concat([
    framedRequest(value),
    ...blobs.map((blob) => framedRequest(blob)),
  ]);
}

function parseFrames(buffer) {
  assert.ok(buffer.byteLength >= 4);
  const length = buffer.readUInt32BE(0);
  const control = JSON.parse(buffer.subarray(4, length + 4).toString("utf8"));
  const blobs = [];
  let offset = length + 4;
  const descriptors = control.artifacts ?? control.diagnosticArtifacts ?? [];
  for (const descriptor of descriptors) {
    assert.ok(buffer.byteLength >= offset + 4);
    const blobLength = buffer.readUInt32BE(offset);
    assert.equal(blobLength, descriptor.byteLength);
    offset += 4;
    const bytes = buffer.subarray(offset, offset + blobLength);
    assert.equal(bytes.byteLength, descriptor.byteLength);
    assert.equal(digest(bytes), descriptor.sha256);
    blobs.push({ descriptor, bytes });
    offset += blobLength;
  }
  assert.equal(offset, buffer.byteLength);
  return { control, blobs };
}

before(async () => {
  await rm(artifacts, { recursive: true, force: true });
  await mkdir(successArtifacts, { recursive: true });
  await mkdir(failureArtifacts, { recursive: true });
  await mkdir(configuredFontDirectory, { recursive: true });
  await mkdir(carlitoFontDirectory, { recursive: true });
  await mkdir(narrowCarlitoFontDirectory, { recursive: true });
  await mkdir(wideCarlitoFontDirectory, { recursive: true });
  await mkdir(loadFailureFontDirectory, { recursive: true });
  await mkdir(emptyFontDirectory, { recursive: true });
  await mkdir(ambiguousFontDirectory, { recursive: true });
  await mkdir(fontScenarioArtifacts, { recursive: true });
  await writeFile(join(configuredFontDirectory, "face.woff2"), await readFile(configuredFontFixture));
  await writeFile(join(carlitoFontDirectory, "face.ttf"),
    await readFile(join(fontFixtureRoot, "synthetic-carlito.ttf")));
  await writeFile(join(narrowCarlitoFontDirectory, "face.ttf"),
    await readFile(join(fontFixtureRoot, "synthetic-carlito-narrow.ttf")));
  await writeFile(join(wideCarlitoFontDirectory, "face.ttf"),
    await readFile(join(fontFixtureRoot, "synthetic-carlito-wide.ttf")));
  await writeFile(join(loadFailureFontDirectory, "face.ttf"),
    await readFile(join(fontFixtureRoot, "docxodus-load-failure.ttf")));
  const configuredFontAttestations = [{
    schemaVersion: 1,
    usage: "standalone-document-font-embedding",
    fileSha256: digest(await readFile(configuredFontFixture)),
    embeddingPermitted: true,
    permittedOutputs: ["html", "pdf"],
    subsettingPermitted: true,
    basis: "Committed DejaVu-derived fixture license",
    attester: "Docxodus test suite",
  }];
  await writeJson(cliFontAttestationPath, configuredFontAttestations);
  await writeJson(hostFontPolicyPath, {
    schemaVersion: 1,
    fontDirectories: ["configured-fonts"],
    fontLicenseAttestations: configuredFontAttestations,
  });
  const ambiguousFirst = Buffer.from(await readFile(configuredFontFixture));
  const ambiguousSecond = Buffer.from(ambiguousFirst);
  ambiguousSecond.writeUInt16BE((ambiguousSecond.readUInt16BE(24) + 1) & 0xffff, 24);
  await writeFile(join(ambiguousFontDirectory, "first.woff2"), ambiguousFirst);
  await writeFile(join(ambiguousFontDirectory, "second.woff2"), ambiguousSecond);
  ambiguousFontAttestations = [ambiguousFirst, ambiguousSecond].map((bytes) => ({
    schemaVersion: 1,
    usage: "standalone-document-font-embedding",
    fileSha256: digest(bytes),
    embeddingPermitted: true,
    permittedOutputs: ["html", "pdf"],
    subsettingPermitted: true,
    basis: "Ambiguity preflight fixture",
    attester: "Docxodus test suite",
  }));
  await writeJson(ambiguousHostFontPolicyPath, {
    schemaVersion: 1,
    fontDirectories: ["ambiguous-fonts"],
    fontLicenseAttestations: ambiguousFontAttestations,
  });
  await writeViewer();
  browser = await chromium.launch({ headless: true });
});

after(async () => {
  await browser?.close();
  await writeViewer();
});

describe("@docxodus/export", { concurrency: false }, () => {
  test("fails pre-transfer with stable typed option and resource errors", async () => {
    const source = new Uint8Array(await readFile(join(fixtures, "CA", "CA001-Plain.docx")));
    await assert.rejects(
      convertDocxToPdf(source, { ...baseOptions, expectedSourceDigest: "0".repeat(64) }),
      (error) => error instanceof DocxodusExportError
        && error.code === "source_digest_mismatch"
        && error.phase === "package_preflight",
    );
    await assert.rejects(
      renderDocxArtifacts(generateFontProbeDocx(), {
        ...baseOptions,
        outputs: ["html"],
        browserExecutablePath: "/definitely-not-a-browser",
        fontDirectories: [ambiguousFontDirectory],
        fontLicenseAttestations: ambiguousFontAttestations,
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_policy_failure"
        && error.phase === "font_loading"
        && /ambiguous files/.test(error.message),
    );
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        limits: { compressedDocxBytes: source.byteLength - 1 },
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_limit"
        && error.phase === "package_preflight",
    );
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        fontDirectories: ["/deployment/fonts"],
        fontLicenseAttestations: [{
          schemaVersion: 1,
          usage: "standalone-document-font-embedding",
          fileSha256: "not-a-digest",
          embeddingPermitted: true,
          permittedOutputs: ["pdf"],
          subsettingPermitted: false,
          basis: "test",
        }],
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "invalid_argument"
        && error.phase === "input_validation",
    );
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        fontDirectories: ["/deployment/fonts"],
        fontLicenseAttestations: ["a", "b"].map(() => ({
          schemaVersion: 1,
          usage: "standalone-document-font-embedding",
          fileSha256: "a".repeat(64),
          embeddingPermitted: true,
          permittedOutputs: ["pdf"],
          subsettingPermitted: false,
          basis: "test-only attestation",
        })),
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "invalid_argument"
        && error.phase === "input_validation"
        && /duplicate file digest/.test(error.message),
    );
    const duplicateFace = {
      family: "Attested Family",
      postscriptName: "Attested-Family-Regular",
      style: "normal",
      weight: 400,
      stretch: 100,
      fileSha256: "b".repeat(64),
      version: "1.0",
    };
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        environmentAttestation: {
          schemaVersion: 1,
          usage: "docxodus-render-environment",
          chromiumProduct: "chromium",
          chromiumBuild: "test",
          launchFlags: [],
          hostFonts: [duplicateFace, { ...duplicateFace, fileSha256: "c".repeat(64) }],
          basis: "test-only attestation",
        },
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "invalid_argument"
        && error.phase === "input_validation"
        && /duplicate face identity/.test(error.message),
    );
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        environmentAttestation: {
          schemaVersion: 1,
          usage: "docxodus-render-environment",
          chromiumProduct: "chromium",
          chromiumBuild: "test",
          launchFlags: ["--unsafe\u0000flag"],
          hostFonts: [],
          basis: "test-only attestation",
        },
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "invalid_argument"
        && error.phase === "input_validation"
        && /bounded plain string/.test(error.message),
    );
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        limits: { fontFiles: 1 },
        environmentAttestation: {
          schemaVersion: 1,
          usage: "docxodus-render-environment",
          chromiumProduct: "chromium",
          chromiumBuild: "test",
          launchFlags: [],
          hostFonts: [duplicateFace, {
            ...duplicateFace,
            family: "Second Attested Family",
            postscriptName: "Second-Attested-Family-Regular",
            fileSha256: "e".repeat(64),
          }],
          basis: "test-only attestation",
        },
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_limit"
        && error.phase === "input_validation"
        && /hostFonts/.test(error.message),
    );

    const fileOutput = join(failureArtifacts, "pre-read-limit-must-not-exist.pdf");
    await assert.rejects(
      renderDocxFile(
        join(fixtures, "CA", "CA001-Plain.docx"),
        { pdfPath: fileOutput },
        { ...baseOptions, limits: { compressedDocxBytes: source.byteLength - 1 } },
      ),
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_limit"
        && error.phase === "package_preflight",
    );
    await assert.rejects(stat(fileOutput), { code: "ENOENT" });

    const runtimeBefore = await currentRuntimeDirectories();
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        browserExecutablePath: join(tmpdir(), "docxodus-browser-does-not-exist"),
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "browser_launch_failure"
        && error.phase === "browser_launch",
    );
    assert.deepEqual(await currentRuntimeDirectories(), runtimeBefore);

    const contextsBefore = browser.contexts().length;
    await assert.rejects(
      convertDocxToPdf(source, { ...baseOptions, browser, timeoutMs: 1 }),
      (error) => error instanceof DocxodusExportError
        && error.code === "readiness_timeout",
    );
    assert.equal(browser.isConnected(), true);
    assert.equal(browser.contexts().length, contextsBefore);
  });

  test("runs canonical print resource bindings with page JavaScript disabled", {
    timeout: 30_000,
  }, async () => {
    const context = await browser.newContext({
      javaScriptEnabled: false,
      serviceWorkers: "block",
      reducedMotion: "reduce",
      colorScheme: "light",
    });
    const page = await context.newPage();
    const requests = [];
    let printReads = 0;
    const origin = "https://print-readiness-test.docxodus.invalid";
    const path = "/final.html";
    const html = `<!doctype html><html><head><meta charset="utf-8"><style id="print-style">
      @page print-test { size: 612pt 792pt; margin: 0 }
      .page-box { box-sizing: border-box; width: 612pt; height: 792pt; page: print-test }
      .print-css, .print-svg { display: none }
      @media print {
        .print-css { display: block; width: 16px; height: 16px;
          background-image: image-set('${tinyPng}' 1x, '${tinyPng}' 2x) }
        .print-svg { display: block }
      }
    </style></head><body><main class="page-box" data-page-number="1"
      data-page-in-section="1" data-section-index="0" data-page-width-pt="612"
      data-page-height-pt="792">
      <div class="print-css" data-source-anchor-id="p:body:print-css"></div>
      <svg class="print-svg" data-source-anchor-id="p:body:print-svg"
        viewBox="0 0 32 32" width="32" height="32">
        <defs><rect id="print-target" width="8" height="8"></rect></defs>
        <image href="${tinyPng}" x="8" width="8" height="8"></image>
        <use href="#print-target" x="16"></use>
      </svg>
    </main></body></html>`;
    await context.route("**/*", async (route) => {
      const request = route.request();
      requests.push(request.url());
      if (request.url() === `${origin}${path}` && request.method() === "GET" && printReads++ === 0) {
        await route.fulfill({
          status: 200,
          body: html,
          contentType: "text/html; charset=utf-8",
          headers: {
            "Content-Security-Policy": "default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'",
          },
        });
      } else {
        await route.abort("blockedbyclient");
      }
    });
    try {
      // These are the exact host-owned production bindings. Their successful
      // use by validatePrintDocument proves they remain callable when page
      // JavaScript itself is disabled.
      await exposePrintReadinessBindings(page);
      await page.goto(`${origin}${path}`, { waitUntil: "load" });
      await page.emulateMedia({ media: "print", colorScheme: "light", reducedMotion: "reduce" });
      const observed = await page.evaluate(() => {
        const pageElement = document.querySelector(".page-box");
        const svg = document.querySelector("svg");
        const image = document.querySelector("svg image");
        const target = document.getElementById("print-target");
        if (!pageElement || !svg || !image || !target) throw new Error("print fixture is incomplete");
        const rect = pageElement.getBoundingClientRect();
        return {
          width: rect.width * 72 / 96,
          height: rect.height * 72 / 96,
          pageName: getComputedStyle(pageElement).page,
          svgMarkup: svg.outerHTML,
          imageMarkup: image.outerHTML,
          targetMarkup: target.outerHTML,
        };
      });
      const targetDigest = domainDigest("docxodus:svg-use-target:v1", observed.targetMarkup);
      const expectedResources = [
        {
          kind: "image",
          status: "embedded",
          resource: "svg-image:p:body:print-svg",
          readiness: "complete",
          contentKey: domainDigest(
            "docxodus:visual-resource:v1",
            JSON.stringify({ source: "svg-image", material: JSON.stringify({
              url: tinyPng,
              markup: observed.imageMarkup,
            }) }),
          ),
        },
        {
          kind: "svg",
          status: "inline",
          resource: "graphic:svg:p:body:print-svg",
          readiness: "complete",
          contentKey: domainDigest(
            "docxodus:visual-resource:v1",
            JSON.stringify({ source: "graphic", markup: observed.svgMarkup }),
          ),
        },
        {
          kind: "svg",
          status: "inline",
          resource: "svg-use:p:body:print-svg",
          readiness: "complete",
          contentKey: domainDigest(
            "docxodus:visual-resource:v1",
            JSON.stringify({ source: "svg-use", href: "#print-target", targetDigest }),
          ),
        },
      ];
      await validatePrintDocument(
        page,
        {
          schemaVersion: 1,
          mode: "paginated",
          availability: "available",
          documentVersion: 1,
          rendererFingerprint: "print-readiness-test",
          pages: [{
            pageNumber: 1,
            pageInSection: 1,
            width: observed.width,
            height: observed.height,
            sectionIndex: 0,
            pageName: observed.pageName,
          }],
          fragments: [],
        },
        5_000,
        {
          fontRequests: 16,
          fontSampleCodePoints: 128,
          visualResources: 16,
          domNodes: 32,
          automaticResourceBytes: 32_768,
        },
        [],
        expectedResources,
      );
      assert.equal(printReads, 1);
      assert.deepEqual(requests, [`${origin}${path}`]);

      await page.evaluate(() => {
        const style = document.getElementById("print-style");
        if (!style) throw new Error("print fixture style is missing");
        style.textContent += "@media print{.print-css{background-image:url(data:image/png;base64,AAAA)}}";
      });
      await assert.rejects(
        validatePrintDocument(
          page,
          {
            schemaVersion: 1,
            mode: "paginated",
            availability: "available",
            documentVersion: 1,
            rendererFingerprint: "print-readiness-test",
            pages: [{
              pageNumber: 1,
              pageInSection: 1,
              width: observed.width,
              height: observed.height,
              sectionIndex: 0,
              pageName: observed.pageName,
            }],
            fragments: [],
          },
          5_000,
          {
            fontRequests: 16,
            fontSampleCodePoints: 128,
            visualResources: 16,
            domNodes: 32,
            automaticResourceBytes: 32_768,
          },
          [],
          expectedResources,
        ),
        (error) => error instanceof DocxodusExportError
          && error.phase === "image_decoding"
          && error.code === "output_verification_failure",
      );
    } finally {
      await context.close();
    }
  });

  test("materializes one immutable batch into searchable tagged PDF and offline HTML", { timeout: 180_000 }, async () => {
    const source = new Uint8Array(await readFile(join(fixtures, "CA", "CA001-Plain.docx")));
    const sourceDigest = digest(source);
    const options = {
      ...baseOptions,
      browser,
      reviewProfileAlreadyApplied: true,
      expectedSourceDigest: sourceDigest,
      outputs: ["html", "pdf"],
      limits: { pdfOutputBytes: 1_000_000 },
    };
    const pending = renderDocxArtifacts(source, options);
    source.fill(0);
    options.outputs.splice(0, options.outputs.length, "invalid-after-entry");
    options.limits.pdfOutputBytes = 1;
    const result = await pending;

    assert.ok(result.html?.startsWith("<!doctype html>"));
    assert.ok(result.pdf?.byteLength > 1_000);
    assert.equal(result.pageCount, result.pageMap.pages.length);
    assert.equal(result.renderReport.source.rawPackageBytesDigest, sourceDigest);
    assert.equal(result.renderReport.bindings.pdfDigest, digest(result.pdf));
    assert.equal(result.renderReport.bindings.pageMapDigest,
      digest(Buffer.from(canonicalJson(result.pageMap))));
    assert.deepEqual(result.renderReport.options.outputs, ["html", "pdf"]);
    assert.equal(result.renderReport.options.policy.timeoutMs, baseOptions.timeoutMs);
    assert.equal(result.renderReport.options.reviewProfileAlreadyApplied, true);
    assert.equal(result.renderReport.environment.verification, "browserObserved");
    assert.deepEqual(result.renderReport.bindings.artifactRequestIds, []);
    assert.ok(result.renderReport.readiness.some((entry) =>
      entry.phase === "pdf_print" && entry.status === "complete"));

    const inspection = await inspectPdf(result.pdf);
    assert.equal(inspection.pageCount, result.pageCount);
    assert.ok(inspection.searchableText.length > 20);
    assert.equal(inspection.marked, true);

    const offlineContext = await browser.newContext({ serviceWorkers: "block" });
    const offline = await offlineContext.newPage();
    const requests = [];
    offline.on("request", (request) => requests.push(request.url()));
    await offline.setContent(result.html, { waitUntil: "load" });
    assert.equal(await offline.locator(".page-box").count(), result.pageCount);
    assert.equal((await offline.locator(".page-box").innerText()).trim().length > 20, true);
    assert.deepEqual(requests, []);
    const screenshot = await offline.screenshot({ fullPage: true });
    await offlineContext.close();

    await writeFile(join(successArtifacts, "generated.pdf"), result.pdf);
    await writeFile(join(successArtifacts, "standalone.html"), result.html);
    await writeJson(join(successArtifacts, "page-map.json"), result.pageMap);
    await writeJson(join(successArtifacts, "render-report.json"), result.renderReport);
    await writeJson(join(successArtifacts, "print-readiness.json"), {
      pageCount: result.pageCount,
      phases: result.renderReport.readiness,
      pagination: result.renderReport.readiness.find(({ phase }) => phase === "pagination"),
      pdfPrint: result.renderReport.readiness.find(({ phase }) => phase === "pdf_print"),
    });
    await writeJson(join(successArtifacts, "pdf-inspection.json"), inspection);
    await writeJson(join(successArtifacts, "request-log.json"), requests);
    await writeFile(join(successArtifacts, "offline-reopen.png"), screenshot);
  });

  test("publishes the five-scenario verified font artifact matrix", { timeout: 300_000 }, async () => {
    const exactSource = generateFontProbeDocx();
    const fontBytes = await readFile(configuredFontFixture);
    const fileSha256 = digest(fontBytes);
    const exactRuntime = {
      strictFonts: true,
      fontDirectories: [configuredFontDirectory],
      fontLicenseAttestations: [{
        schemaVersion: 1,
        usage: "standalone-document-font-embedding",
        fileSha256,
        embeddingPermitted: true,
        permittedOutputs: ["html", "pdf"],
        subsettingPermitted: true,
        basis: "Committed DejaVu-derived fixture license",
        attester: "Docxodus test suite",
      }],
    };
    const exactCatalogManifest = pathFreeCatalogManifest(await discoverFontCatalog(
      [configuredFontDirectory],
      exactRuntime.fontLicenseAttestations,
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    ));
    await writeJson(join(fontScenarioArtifacts, "catalog-manifest.json"), exactCatalogManifest);
    assert.equal(JSON.stringify(exactCatalogManifest).includes(artifacts), false);
    assert.equal(JSON.stringify(exactCatalogManifest).includes("bytesBase64"), false);

    const metricSource = generateFontProbeDocx(
      "Calibri Light",
      "DOCUMENTED METRIC DIFFERENCE WRAPPING DOCUMENTED METRIC DIFFERENCE WRAPPING",
      { pageWidth: 4320, pageHeight: 7920, margin: 360, paragraphCount: 24 },
    );
    let metricBaselineResult;
    let metricBaselineEvidence;
    const scenarios = [{
      id: "exact",
      source: exactSource,
      options: exactRuntime,
      verify(result) {
        assert.equal(result.renderReport.environment.verification, "nodeVerified");
        assert.ok(result.renderReport.fonts.every((font) =>
          font.status === "resolved"
          && font.resolvedFamily === "Docxodus Canvas Mono"
          && font.faceMatch === "exact"
          && font.glyphCoverage === "complete"
          && font.fileSha256 === fileSha256
          && font.licenseEvidence?.kind === "attested"));
      },
    }, {
      id: "substitution",
      source: generateFontProbeDocx("Calibri", "METRIC COMPATIBLE SUBSTITUTION"),
      options: { browser, fontDirectories: [carlitoFontDirectory] },
      verify(result) {
        assert.ok(result.renderReport.fonts.every((font) =>
          font.status === "substituted"
          && font.resolvedFamily === "Carlito"
          && font.metricCompatible === true));
      },
    }, {
      id: "missing",
      source: generateFontProbeDocx("Definitely Missing Font", "MISSING FONT FALLBACK"),
      options: { browser, fontDirectories: [carlitoFontDirectory] },
      verify(result) {
        assert.ok(result.renderReport.fonts.every((font) =>
          font.source === "browser"
          && font.status === "missing"
          && font.browserFallbackAvailable === true));
        assert.ok(result.warnings.some(({ code }) => code === "font_unavailable"));
      },
    }, {
      id: "load-failure",
      source: generateFontProbeDocx("Docxodus Load Failure", "DECODE LOAD FAILURE"),
      options: { browser, fontDirectories: [loadFailureFontDirectory] },
      verify(result) {
        assert.ok(result.renderReport.fonts.every((font) => font.status === "load_failed"));
        assert.ok(result.warnings.some(({ code }) => code === "font_load_failed"));
      },
    }, {
      id: "metric-difference",
      source: metricSource,
      options: { browser, fontDirectories: [wideCarlitoFontDirectory] },
      verify(result) {
        assert.ok(result.renderReport.fonts.every((font) =>
          font.status === "substituted"
          && font.resolvedFamily === "Carlito"
          && font.metricCompatible === false));
        assert.ok(result.warnings.some(({ code }) => code === "font_metric_mismatch"));
        assert.ok(metricBaselineResult);
        assert.ok(metricBaselineResult.renderReport.fonts.every((font) =>
          font.status === "substituted"
          && font.requestedFamily === "Calibri Light"
          && font.resolvedFamily === "Carlito"
          && font.metricCompatible === false));
        const resolutionSemantics = (font) => ({
          requestId: font.requestId,
          requestedFamily: font.requestedFamily,
          requestedFamilies: font.requestedFamilies,
          requestedFamilyKinds: font.requestedFamilyKinds,
          requestedStyle: font.requestedStyle,
          requestedWeight: font.requestedWeight,
          requestedStretch: font.requestedStretch,
          status: font.status,
          source: font.source,
          resolvedFamily: font.resolvedFamily,
          faceMatch: font.faceMatch,
          metricCompatible: font.metricCompatible,
          glyphCoverage: font.glyphCoverage,
        });
        assert.deepEqual(
          metricBaselineResult.renderReport.fonts.map(resolutionSemantics),
          result.renderReport.fonts.map(resolutionSemantics),
          "narrow and wide faces must use identical requested descriptors and resolution semantics",
        );
        const baselineFontDigests = new Set(metricBaselineResult.renderReport.fonts
          .map(({ fileSha256: sha256 }) => sha256));
        const comparisonFontDigests = new Set(result.renderReport.fonts
          .map(({ fileSha256: sha256 }) => sha256));
        assert.equal(baselineFontDigests.size, 1);
        assert.equal(comparisonFontDigests.size, 1);
        assert.notEqual(
          [...baselineFontDigests][0],
          [...comparisonFontDigests][0],
          "narrow and wide faces must use different immutable font bytes",
        );
        const baselineGeometry = JSON.stringify(metricBaselineResult.pageMap.fragments.map((fragment) => ({
          pageNumber: fragment.pageNumber,
          geometry: fragment.geometry,
        })));
        const comparisonGeometry = JSON.stringify(result.pageMap.fragments.map((fragment) => ({
          pageNumber: fragment.pageNumber,
          geometry: fragment.geometry,
        })));
        assert.ok(metricBaselineResult.pageCount !== result.pageCount
          || baselineGeometry !== comparisonGeometry,
        "narrow and wide faces must produce a concrete PageMap geometry difference");
        assert.notEqual(metricBaselineResult.rendererFingerprint, result.rendererFingerprint);
      },
    }];
    const evidence = [];
    const failures = [];
    const writeComparison = async () => writeJson(
      join(fontScenarioArtifacts, "comparison-manifest.json"),
      {
        schemaVersion: 1,
        catalogManifest: "catalog-manifest.json",
        scenarios: evidence,
        metricComparison: {
          sourceDigest: digest(metricSource),
          baseline: metricBaselineEvidence,
          comparison: evidence.find(({ id }) => id === "metric-difference"),
          sameSource: metricBaselineEvidence?.sourceDigest === digest(metricSource),
          pageCountChanged: metricBaselineEvidence?.pageCount
            !== evidence.find(({ id }) => id === "metric-difference")?.pageCount,
          pageMapDigestChanged: metricBaselineEvidence?.pageMapDigest
            !== evidence.find(({ id }) => id === "metric-difference")?.pageMapDigest,
          fragmentGeometryChanged: JSON.stringify(metricBaselineEvidence?.pageGeometry?.fragments)
            !== JSON.stringify(evidence.find(({ id }) => id === "metric-difference")
              ?.pageGeometry?.fragments),
          rendererFingerprintChanged: metricBaselineEvidence?.rendererFingerprint
            !== evidence.find(({ id }) => id === "metric-difference")?.rendererFingerprint,
        },
        fingerprintComparison: {
          distinctCount: new Set(evidence.map(({ rendererFingerprint }) => rendererFingerprint)
            .filter(Boolean)).size,
          exactVsSubstitutionDiffer: evidence.find(({ id }) => id === "exact")?.rendererFingerprint
            !== evidence.find(({ id }) => id === "substitution")?.rendererFingerprint,
          substitutionVsMetricDifferenceDiffer:
            evidence.find(({ id }) => id === "substitution")?.rendererFingerprint
            !== evidence.find(({ id }) => id === "metric-difference")?.rendererFingerprint,
        },
      },
    );

    try {
      metricBaselineResult = await renderDocxArtifacts(metricSource, {
        ...baseOptions,
        browser,
        fontDirectories: [narrowCarlitoFontDirectory],
        outputs: ["html", "pdf"],
      });
      metricBaselineEvidence = await writeFontScenario(
        "metric-baseline",
        metricSource,
        metricBaselineResult,
      );
    } catch (error) {
      const directory = join(fontScenarioArtifacts, "metric-baseline");
      await mkdir(directory, { recursive: true });
      const failure = error instanceof DocxodusExportError
        ? error.toJSON()
        : { name: error?.name ?? "Error", message: error?.message ?? String(error) };
      await writeJson(join(directory, "error.json"), failure);
      metricBaselineEvidence = { id: "metric-baseline", error: failure };
      failures.push(error);
    }
    await writeComparison();

    for (const scenario of scenarios) {
      let scenarioEvidence;
      try {
        const result = await renderDocxArtifacts(scenario.source, {
          ...baseOptions,
          ...scenario.options,
          outputs: ["html", "pdf"],
        });
        scenarioEvidence = await writeFontScenario(scenario.id, scenario.source, result);
        evidence.push(scenarioEvidence);
        const reportJson = JSON.stringify(result.renderReport);
        if (scenario.id === "exact") {
          await writeFile(join(successArtifacts, "configured-font.docx"), scenario.source);
          await writeFile(join(successArtifacts, "configured-font.pdf"), result.pdf);
          await writeFile(join(successArtifacts, "configured-font.html"), result.html);
          await writeJson(join(successArtifacts, "configured-font-report.json"), result.renderReport);
        }
        scenario.verify(result);
        assert.equal(reportJson.includes(artifacts), false);
        assert.equal(reportJson.includes("bytesBase64"), false);
        if (scenario.id === "exact") {
          assert.equal(reportJson.includes(Buffer.from(fontBytes).toString("base64")), false);
        }
      } catch (error) {
        const directory = join(fontScenarioArtifacts, scenario.id);
        await mkdir(directory, { recursive: true });
        const failure = error instanceof DocxodusExportError
          ? error.toJSON()
          : { name: error?.name ?? "Error", message: error?.message ?? String(error) };
        await writeJson(join(directory, "error.json"), failure);
        if (scenarioEvidence) scenarioEvidence.validationError = failure;
        else evidence.push({ id: scenario.id, error: failure, artifacts: { error: `${scenario.id}/error.json` } });
        failures.push(error);
      }
      await writeComparison();
    }

    const unobservableExecutable = await renderDocxArtifacts(exactSource, {
      ...baseOptions,
      ...exactRuntime,
      browser,
      outputs: ["html"],
      environmentAttestation: {
        schemaVersion: 1,
        usage: "docxodus-render-environment",
        chromiumProduct: "chromium",
        chromiumBuild: browser.version(),
        executableSha256: "d".repeat(64),
        launchFlags: [],
        hostFonts: [],
        basis: "Negative test: an injected browser cannot prove its executable bytes",
      },
    });
    assert.equal(unobservableExecutable.renderReport.environment.verification, "browserObserved");
    assert.equal(evidence.length, 5);
    assert.equal(failures.length, 0, failures.map((error) => error?.message ?? String(error)).join("\n"));
  });

  test("applies host font attestation before strict policy and final fingerprinting", { timeout: 180_000 }, async () => {
    const source = generateFontProbeDocx("Arial", "HOST ATTESTED FONT IDENTITY");
    const usedFace = {
      family: "Arial",
      postscriptName: "Arial-Regular",
      style: "normal",
      weight: 400,
      stretch: 100,
      fileSha256: "1".repeat(64),
      version: "test-1",
    };
    const attestation = (overrides = {}) => ({
      schemaVersion: 1,
      usage: "docxodus-render-environment",
      chromiumProduct: "chromium",
      chromiumBuild: browser.version(),
      launchFlags: [],
      hostFonts: [usedFace],
      basis: "Test deployment inventory",
      ...overrides,
    });
    const render = (environmentAttestation) => renderDocxArtifacts(source, {
      ...baseOptions,
      browser,
      strictFonts: true,
      environmentAttestation,
      outputs: ["html"],
    });

    const first = await render(attestation());
    assert.equal(first.renderReport.environment.verification, "callerAttested");
    assert.ok(first.renderReport.fonts.every((font) =>
      font.status === "resolved"
      && font.source === "attested"
      && font.resolvedFace === usedFace.postscriptName
      && font.fileSha256 === usedFace.fileSha256
      && font.licenseEvidence === undefined));
    assert.equal(first.warnings.some(({ code }) => code === "font_environment_unverified"), false);
    assert.equal(first.rendererFingerprint, first.renderReport.environment.rendererFingerprint);
    assert.equal(first.rendererFingerprint, first.pageMap.rendererFingerprint);

    const semanticallySame = await render(attestation({
      basis: "Different explanatory prose must not change layout identity",
      hostFonts: [usedFace, {
        ...usedFace,
        family: "Unused Host Face",
        postscriptName: "Unused-Host-Face-Regular",
        fileSha256: "2".repeat(64),
      }],
    }));
    assert.equal(semanticallySame.rendererFingerprint, first.rendererFingerprint);
    assert.equal(
      semanticallySame.renderReport.fontIdentity.resolutionDigest,
      first.renderReport.fontIdentity.resolutionDigest,
    );

    const changedFace = { ...usedFace, fileSha256: "3".repeat(64) };
    const changed = await render(attestation({ hostFonts: [changedFace] }));
    assert.notEqual(changed.rendererFingerprint, first.rendererFingerprint);
    assert.notEqual(
      changed.renderReport.fontIdentity.resolutionDigest,
      first.renderReport.fontIdentity.resolutionDigest,
    );

    await assert.rejects(
      render(attestation({ chromiumBuild: "mismatched-browser-build" })),
      (error) => error instanceof DocxodusExportError
        && error.code === "output_verification_failure"
        && error.phase === "output_verification"
        && error.report?.warnings.some((warning) =>
          warning.code === "font_environment_unverified" && warning.severity === "warning"),
    );

    await assert.rejects(
      render(attestation({ hostFonts: [{ ...usedFace, weight: 700 }] })),
      (error) => error instanceof DocxodusExportError
        && error.code === "output_verification_failure"
        && error.phase === "output_verification"
        && error.detail?.includes("unattestedFont=Arial"),
    );
  });

  test("binds effective injected-runtime attestation into the final fingerprint", { timeout: 180_000 }, async () => {
    const source = generateFontProbeDocx();
    const fontBytes = await readFile(configuredFontFixture);
    const runtime = {
      ...baseOptions,
      browser,
      strictFonts: true,
      fontDirectories: [configuredFontDirectory],
      fontLicenseAttestations: [{
        schemaVersion: 1,
        usage: "standalone-document-font-embedding",
        fileSha256: digest(fontBytes),
        embeddingPermitted: true,
        permittedOutputs: ["html", "pdf"],
        subsettingPermitted: true,
        basis: "Committed DejaVu-derived fixture license",
        attester: "Docxodus test suite",
      }],
      outputs: ["html"],
    };
    const environment = (launchFlags) => ({
      schemaVersion: 1,
      usage: "docxodus-render-environment",
      chromiumProduct: "chromium",
      chromiumBuild: browser.version(),
      launchFlags,
      hostFonts: [],
      basis: "Injected browser launch configuration",
    });

    const observed = await renderDocxArtifacts(source, runtime);
    const attestedA = await renderDocxArtifacts(source, {
      ...runtime,
      environmentAttestation: environment(["--deployment-profile=a"]),
    });
    const attestedB = await renderDocxArtifacts(source, {
      ...runtime,
      environmentAttestation: environment(["--deployment-profile=b"]),
    });

    assert.equal(observed.renderReport.environment.verification, "browserObserved");
    assert.equal(attestedA.renderReport.environment.verification, "callerAttested");
    assert.equal(attestedB.renderReport.environment.verification, "callerAttested");
    assert.deepEqual(attestedA.renderReport.fonts, observed.renderReport.fonts);
    assert.deepEqual(attestedB.renderReport.fonts, observed.renderReport.fonts);
    assert.notEqual(attestedA.rendererFingerprint, observed.rendererFingerprint);
    assert.notEqual(attestedB.rendererFingerprint, attestedA.rendererFingerprint);
  });

  test("preserves hyperlink annotations and chart vectors", { timeout: 180_000 }, async () => {
    const linkSource = new Uint8Array(await readFile(join(fixtures, "HC023-Hyperlink.docx")));
    const linkResult = await convertDocxToPdf(linkSource, { ...baseOptions, browser });
    const linkInspection = await inspectPdf(linkResult.pdf);
    assert.ok(linkInspection.linkAnnotations > 0);

    const chartSource = new Uint8Array(await readFile(join(fixtures, "HC043-Chart.docx")));
    const chartResult = await renderDocxArtifacts(chartSource, {
      ...baseOptions,
      browser,
      outputs: ["html", "pdf"],
    });
    assert.match(chartResult.html, /<svg\b/i);
    const chartInspection = await inspectPdf(chartResult.pdf);
    assert.ok(chartInspection.vectorPathOperations > 0);
    assert.equal(browser.isConnected(), true);

    await writeFile(join(successArtifacts, "hyperlinks.pdf"), linkResult.pdf);
    await writeJson(join(successArtifacts, "hyperlink-inspection.json"), linkInspection);
    await writeFile(join(successArtifacts, "chart-vector.pdf"), chartResult.pdf);
    await writeJson(join(successArtifacts, "chart-vector-inspection.json"), chartInspection);
  });

  test("preserves mixed section geometry, stories, notes, numbering, and print scale", { timeout: 180_000 }, async () => {
    const source = generateMixedSectionDocx();
    const result = await renderDocxArtifacts(source, {
      ...baseOptions,
      browser,
      outputs: ["html", "pdf"],
    });
    const repeated = await renderDocxArtifacts(source, {
      ...baseOptions,
      browser,
      outputs: ["html"],
    });
    const expectedPages = [
      { pageNumber: 1, pageInSection: 1, width: 612, height: 792, sectionIndex: 0 },
      { pageNumber: 2, pageInSection: 2, width: 612, height: 792, sectionIndex: 0 },
      { pageNumber: 3, pageInSection: 1, width: 792, height: 612, sectionIndex: 1 },
      { pageNumber: 4, pageInSection: 1, width: 595.3, height: 841.9, sectionIndex: 2 },
      { pageNumber: 5, pageInSection: 2, width: 595.3, height: 841.9, sectionIndex: 3 },
      { pageNumber: 6, pageInSection: 1, width: 612, height: 792, sectionIndex: 4 },
    ];
    assert.equal(result.pageCount, expectedPages.length);
    assert.deepEqual(
      result.pageMap.pages.map(({ pageNumber, pageInSection, width, height, sectionIndex }) => ({
        pageNumber,
        pageInSection,
        width,
        height,
        sectionIndex,
      })),
      expectedPages,
    );
    // The report inventory is the complete PageMap page contract, including pageInSection and
    // pageName; comparing projections here previously made this success test fail by construction.
    assert.deepEqual(result.renderReport.pages, result.pageMap.pages);
    assert.equal(repeated.pageCount, result.pageCount);
    assert.deepEqual(repeated.pageMap, result.pageMap);
    assert.deepEqual(repeated.renderReport.pages, result.renderReport.pages);
    assert.equal(repeated.rendererFingerprint, result.rendererFingerprint);
    assert.equal(repeated.renderReport.options.policy.timeoutMs, baseOptions.timeoutMs);
    assert.match(result.html, /column-count:\s*2/i);

    const normalInspection = await inspectPdf(result.pdf);
    assertPdfBoxes(normalInspection, expectedPages);
    const pageText = normalInspection.pages.map((page) => page.text.replace(/\s+/g, " ").trim());
    const expectedText = [
      ["HEADER-S0", "BODY-S0-P1 LETTER PORTRAIT", "FOOTER-S0 PAGE 1"],
      ["HEADER-S0", "BODY-S0-P2 EXPLICIT PAGE BREAK", "FOOTER-S0 PAGE 2"],
      ["HEADER-S1", "BODY-S1 LANDSCAPE", "LANDSCAPE FOOTNOTE TOKEN", "FOOTER-S1 PAGE 10"],
      ["HEADER-S2", "BODY-S2 A4 PORTRAIT BEFORE CONTINUOUS", "BODY-S3 TWO COLUMN SHARED PAGE", "FOOTER-S2 PAGE 11"],
      ["HEADER-S3", "BODY-S3 TWO COLUMN SPILL PAGE", "FOOTER-S3 PAGE 12"],
      ["HEADER-S4", "BODY-S4 LETTER PORTRAIT FINAL", "FOOTER-S4 PAGE 13"],
    ];
    const expectedTokenPages = new Map();
    expectedText.forEach((tokens, pageIndex) => {
      tokens.forEach((token) => {
        const owners = expectedTokenPages.get(token) ?? new Set();
        owners.add(pageIndex);
        expectedTokenPages.set(token, owners);
      });
    });
    expectedText.forEach((tokens, index) => {
      let precedingTokenIndex = -1;
      tokens.forEach((token) => {
        const tokenIndex = pageText[index].indexOf(token);
        assert.ok(
          tokenIndex > precedingTokenIndex,
          `page ${index + 1} token ${JSON.stringify(token)} is missing or out of order: ${pageText[index]}`,
        );
        assert.equal(
          pageText[index].indexOf(token, tokenIndex + token.length),
          -1,
          `page ${index + 1} repeats ${JSON.stringify(token)}: ${pageText[index]}`,
        );
        pageText.forEach((otherPageText, otherPageIndex) => {
          if (!expectedTokenPages.get(token).has(otherPageIndex)) {
            assert.equal(
              otherPageText.includes(token),
              false,
              `${JSON.stringify(token)} leaked from page ${index + 1} to page ${otherPageIndex + 1}`,
            );
          }
        });
        precedingTokenIndex = tokenIndex;
      });
    });
    assert.equal(pageText[3].includes("HEADER-S3"), false);
    assert.equal(pageText[3].includes("FOOTER-S3"), false);

    const scaledContext = await browser.newContext({ serviceWorkers: "block" });
    const scaledPage = await scaledContext.newPage();
    const requests = [];
    scaledPage.on("request", (request) => requests.push(request.url()));
    await scaledPage.setContent(result.html, { waitUntil: "load" });
    await scaledPage.evaluate(async () => { await document.fonts.ready; });
    const screenScales = await scaledPage.locator(".page-box").evaluateAll((nodes) =>
      nodes.map((node, index) => {
        // Alternate the two real viewer paths so every sheet remains at an effective 80% scale:
        // Chromium's inline zoom and the transform fallback with compensating margins. Print CSS
        // must neutralize both paths.
        if (index % 2 === 0) {
          node.style.zoom = "0.8";
        } else {
          node.style.transform = "scale(0.8)";
          node.style.transformOrigin = "top left";
          node.style.marginRight = "-100px";
          node.style.marginBottom = "-100px";
        }
        const style = getComputedStyle(node);
        return {
          zoom: style.zoom,
          transform: style.transform,
          transformOrigin: style.transformOrigin,
          marginRight: style.marginRight,
          marginBottom: style.marginBottom,
        };
      }));
    for (const [index, scale] of screenScales.entries()) {
      if (index % 2 === 0) {
        assert.equal(scale.zoom, "0.8");
        assert.equal(scale.transform, "none");
        assert.equal(scale.marginRight, "0px");
        assert.equal(scale.marginBottom, "0px");
      } else {
        assert.equal(scale.zoom, "1");
        assert.match(scale.transform, /^matrix\(0\.8, 0, 0, 0\.8, 0, 0\)$/);
        assert.equal(scale.transformOrigin, "0px 0px");
        assert.equal(scale.marginRight, "-100px");
        assert.equal(scale.marginBottom, "-100px");
      }
    }
    const columnContainers = await scaledPage.locator(".page-content > div").evaluateAll((nodes) =>
      nodes.filter((node) => getComputedStyle(node).columnCount === "2").length);
    assert.ok(columnContainers >= 2);
    const screenshot = await scaledPage.screenshot({ fullPage: true });
    await scaledPage.emulateMedia({ media: "print" });
    const printAudit = await scaledPage.evaluate((pageMap) => {
      const pages = Array.from(document.querySelectorAll(".page-box"));
      return {
        pages: pages.map((page) => {
          const style = getComputedStyle(page);
          const rect = page.getBoundingClientRect();
          return {
            zoom: style.zoom,
            transform: style.transform,
            marginTop: style.marginTop,
            marginRight: style.marginRight,
            marginBottom: style.marginBottom,
            marginLeft: style.marginLeft,
            widthPt: rect.width * 72 / 96,
            heightPt: rect.height * 72 / 96,
          };
        }),
        fragments: pageMap.fragments.map((expected) => {
          const element = document.querySelector(
            `[data-page-fragment-id="${CSS.escape(expected.fragmentId)}"]`,
          );
          const page = element?.closest(".page-box");
          if (!element || !page) return { fragmentId: expected.fragmentId, missing: true };
          const pageRect = page.getBoundingClientRect();
          const rect = element.getBoundingClientRect();
          const pageContract = pageMap.pages[expected.pageNumber - 1];
          const xRatio = pageContract.width / pageRect.width;
          const yRatio = pageContract.height / pageRect.height;
          return {
            fragmentId: expected.fragmentId,
            missing: false,
            geometry: {
              x: (rect.left - pageRect.left) * xRatio,
              y: (rect.top - pageRect.top) * yRatio,
              width: rect.width * xRatio,
              height: rect.height * yRatio,
            },
          };
        }),
      };
    }, result.pageMap);
    printAudit.pages.forEach((page, index) => {
      assert.equal(page.zoom, "1", `page ${index + 1} retained viewer zoom in print`);
      assert.equal(page.transform, "none", `page ${index + 1} retained viewer transform in print`);
      for (const margin of ["marginTop", "marginRight", "marginBottom", "marginLeft"]) {
        assert.equal(page[margin], "0px", `page ${index + 1} retained ${margin} in print`);
      }
      assert.ok(Math.abs(page.widthPt - expectedPages[index].width) <= 0.1);
      assert.ok(Math.abs(page.heightPt - expectedPages[index].height) <= 0.1);
    });
    printAudit.fragments.forEach((actual, index) => {
      const expected = result.pageMap.fragments[index];
      assert.equal(actual.missing, false, `print tree lost ${expected.fragmentId}`);
      for (const coordinate of ["x", "y", "width", "height"]) {
        assert.ok(
          Math.abs(actual.geometry[coordinate] - expected.geometry[coordinate]) <= 0.1,
          `${expected.fragmentId} print ${coordinate} changed`,
        );
      }
    });
    const scaledPdf = await scaledPage.pdf({
      printBackground: true,
      preferCSSPageSize: true,
      tagged: true,
      outline: false,
      displayHeaderFooter: false,
      scale: 1,
      margin: { top: "0", right: "0", bottom: "0", left: "0" },
    });
    await scaledContext.close();
    assert.deepEqual(requests, []);

    const scaledInspection = await inspectPdf(scaledPdf);
    assertPdfBoxes(scaledInspection, expectedPages);
    assertEquivalentPdfBoxes(scaledInspection, normalInspection);
    assertEquivalentPdfTextGeometry(scaledInspection, normalInspection);

    await writeFile(join(successArtifacts, "mixed-sections.docx"), source);
    await writeFile(join(successArtifacts, "mixed-sections.html"), result.html);
    await writeFile(join(successArtifacts, "mixed-sections.pdf"), result.pdf);
    await writeFile(join(successArtifacts, "mixed-sections-scaled-viewer.pdf"), scaledPdf);
    await writeJson(join(successArtifacts, "mixed-sections-page-map.json"), result.pageMap);
    await writeJson(join(successArtifacts, "mixed-sections-render-report.json"), result.renderReport);
    await writeJson(join(successArtifacts, "mixed-sections-inspection.json"), {
      expectedPages,
      normal: normalInspection,
      scaled: scaledInspection,
      screenScales,
      printAudit,
      offlineRequests: requests,
      columnContainers,
      repeated: {
        pageCount: repeated.pageCount,
        pages: repeated.renderReport.pages,
        pageMapDigest: repeated.renderReport.bindings.pageMapDigest,
        htmlDigest: repeated.renderReport.bindings.htmlDigest,
        rendererFingerprint: repeated.rendererFingerprint,
      },
    });
    await writeFile(join(successArtifacts, "mixed-sections-scaled-viewer.png"), screenshot);
  });

  test("preserves one converter-produced footnote paragraph across HTML and PDF pages", {
    timeout: 180_000,
  }, async () => {
    const source = generateLongFootnoteDocx(600);
    const result = await renderDocxArtifacts(source, {
      ...baseOptions,
      browser,
      outputs: ["html", "pdf"],
    });
    const inspection = await inspectPdf(result.pdf);

    assert.ok(result.pageCount > 2);
    assert.equal(inspection.pageCount, result.pageCount);
    assert.deepEqual(result.renderReport.pages, result.pageMap.pages);
    assertPdfBoxes(inspection, result.pageMap.pages);
    const paragraphFragments = result.pageMap.fragments.filter((fragment) =>
      fragment.story === "footnote" && fragment.anchorId.startsWith("p:fn:"));
    assert.ok(paragraphFragments.length > 2);
    assert.ok(new Set(paragraphFragments.map((fragment) => fragment.pageNumber)).size > 2);
    assert.deepEqual(
      paragraphFragments.map((fragment) => fragment.fragmentIndex),
      paragraphFragments.map((_, index) => index),
    );

    let priorPdfOffset = -1;
    for (const index of [1, 2, 150, 300, 450, 599, 600]) {
      const token = `footnote-1-1-${index}`;
      assert.equal(result.html.match(new RegExp(`\\b${token}\\b`, "g"))?.length, 1);
      const pdfOffset = inspection.searchableText.indexOf(token);
      assert.ok(pdfOffset > priorPdfOffset, `${token} missing or out of order in PDF text`);
      assert.equal(inspection.searchableText.indexOf(token, pdfOffset + token.length), -1);
      priorPdfOffset = pdfOffset;
    }

    await writeFile(join(successArtifacts, "long-footnote.pdf"), result.pdf);
    await writeFile(join(successArtifacts, "long-footnote.html"), result.html);
    await writeJson(join(successArtifacts, "long-footnote-page-map.json"), result.pageMap);
    await writeJson(join(successArtifacts, "long-footnote-pdf-inspection.json"), inspection);
  });

  test("retains a structured report when post-layout PDF verification fails", { timeout: 180_000 }, async () => {
    const source = new Uint8Array(await readFile(join(fixtures, "CA", "CA001-Plain.docx")));
    let failure;
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        browser,
        limits: { pdfOutputBytes: 1 },
      }),
      (error) => {
        failure = error;
        return error instanceof DocxodusExportError
          && error.code === "resource_limit"
          && error.phase === "pdf_print"
          && error.report?.status === "failed";
      },
    );
    assert.equal(failure.report.failure.code, "resource_limit");
    assert.ok(failure.report.partial.pages.length > 0);
    await writeJson(join(failureArtifacts, "pdf-limit-error.json"), failure.toJSON());
    await writeJson(join(failureArtifacts, "pdf-limit-render-report.json"), failure.report);
  });

  test("CLI writes verified artifacts, never stdout bytes, and never overwrites", { timeout: 180_000 }, async () => {
    const input = join(fixtures, "CA", "CA001-Plain.docx");
    const pdf = join(successArtifacts, "cli.pdf");
    const report = join(successArtifacts, "cli-render-report.json");
    const pageMap = join(successArtifacts, "cli-page-map.json");
    const runtimeBefore = await currentRuntimeDirectories();
    const args = [
      cliEntry,
      "convert",
      input,
      "--to", "pdf",
      "--output", pdf,
      "--review-profile", "final",
      "--comments", "hidden",
      "--report", report,
      "--page-map", pageMap,
      "--timeout", "120000",
    ];
    const environment = {
      ...process.env,
      DOCXODUS_CHROMIUM_PATH: chromium.executablePath(),
    };
    const first = spawnSync(process.execPath, args, {
      cwd: packageRoot,
      env: environment,
      encoding: null,
      maxBuffer: 10 * 1024 * 1024,
    });
    assert.equal(first.status, 0, first.stderr.toString());
    assert.equal(first.stdout.byteLength, 0);
    assert.match(first.stderr.toString(), /Rendered \d+ page/);
    const originalDigest = digest(await readFile(pdf));
    assert.ok((await stat(pdf)).size > 1_000);
    const pageMapBytes = await readFile(pageMap);
    const reportBytes = await readFile(report);
    assert.equal(pageMapBytes.at(-1), "}".charCodeAt(0));
    assert.equal(reportBytes.at(-1), "}".charCodeAt(0));
    const cliReport = JSON.parse(reportBytes.toString("utf8"));
    assert.equal(cliReport.bindings.pageMapDigest, digest(pageMapBytes));
    assert.equal(cliReport.bindings.pdfDigest, originalDigest);

    const second = spawnSync(process.execPath, args, {
      cwd: packageRoot,
      env: environment,
      encoding: null,
      maxBuffer: 10 * 1024 * 1024,
    });
    assert.equal(second.status, 1);
    assert.equal(second.stdout.byteLength, 0);
    assert.match(second.stderr.toString(), /destination already exists/i);
    assert.equal(digest(await readFile(pdf)), originalDigest);

    const fontInput = join(successArtifacts, "cli-configured-font.docx");
    const fontPdf = join(successArtifacts, "cli-configured-font.pdf");
    const fontReport = join(successArtifacts, "cli-configured-font-report.json");
    await writeFile(fontInput, generateFontProbeDocx());
    const fontCli = spawnSync(process.execPath, [
      cliEntry,
      "convert",
      fontInput,
      "--to", "pdf",
      "--output", fontPdf,
      "--review-profile", "final",
      "--comments", "hidden",
      "--strict-fonts",
      "--font-directory", emptyFontDirectory,
      "--font-directory", configuredFontDirectory,
      "--font-license-attestations", cliFontAttestationPath,
      "--report", fontReport,
      "--timeout", "120000",
    ], {
      cwd: packageRoot,
      env: environment,
      encoding: null,
      maxBuffer: 10 * 1024 * 1024,
    });
    assert.equal(fontCli.status, 0, fontCli.stderr.toString());
    assert.equal(fontCli.stdout.byteLength, 0);
    const fontCliReport = JSON.parse(await readFile(fontReport, "utf8"));
    assert.ok(fontCliReport.fonts.every((font) =>
      font.status === "resolved" && font.resolvedFamily === "Docxodus Canvas Mono"));
    assert.deepEqual(await currentRuntimeDirectories(), runtimeBefore);
  });

  test("framed host owns runtime/font policy and returns keyed artifact results", { timeout: 180_000 }, async () => {
    const source = await readFile(join(fixtures, "CA", "CA001-Plain.docx"));
    const sourceDescriptor = {
      id: "source-1",
      byteLength: source.byteLength,
      sha256: digest(source),
      mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    };
    const ambiguous = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      env: {
        ...process.env,
        DOCXODUS_CHROMIUM_PATH: "/definitely-not-a-browser",
        DOCXODUS_FONT_POLICY_PATH: ambiguousHostFontPolicyPath,
      },
      input: framedProtocolRequest({
        schemaVersion: 1,
        sources: [sourceDescriptor],
        batches: [{
          id: "ambiguous-font-policy",
          sourceId: "source-1",
          artifactRequestIds: ["ambiguous-report"],
          options: { ...baseOptions, outputs: ["html"] },
        }],
      }, [source]),
      maxBuffer: 10 * 1024 * 1024,
    });
    assert.equal(ambiguous.status, 0, ambiguous.stderr.toString());
    const ambiguousFailure = parseFrames(ambiguous.stdout).control.fatal;
    assert.equal(ambiguousFailure.code, "resource_policy_failure");
    assert.equal(ambiguousFailure.phase, "font_loading");
    assert.match(ambiguousFailure.message, /ambiguous files/);

    const forbidden = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      input: framedProtocolRequest({
        schemaVersion: 1,
        sources: [sourceDescriptor],
        batches: [{
          id: "forbidden-runtime",
          sourceId: "source-1",
          artifactRequestIds: ["delivery-1"],
          options: {
            ...baseOptions,
            outputs: ["pdf"],
            browserExecutablePath: "/tmp/untrusted-executable",
            fontDirectories: ["/tmp/untrusted-fonts"],
          },
        }],
      }, [source]),
      maxBuffer: 10 * 1024 * 1024,
    });
    assert.equal(forbidden.status, 0);
    assert.match(parseFrames(forbidden.stdout).control.fatal.message,
      /unknown fields: browserExecutablePath, fontDirectories/);

    const fontSource = generateFontProbeDocx();
    const fontSourceDescriptor = {
      id: "font-source",
      byteLength: fontSource.byteLength,
      sha256: digest(fontSource),
      mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    };
    const request = {
      schemaVersion: 1,
      sources: [sourceDescriptor, fontSourceDescriptor],
      batches: [{
        id: "artifact-1",
        sourceId: "source-1",
        artifactRequestIds: ["delivery-pdf"],
        options: { ...baseOptions, outputs: ["pdf"] },
      }, {
        id: "artifact-2",
        sourceId: "source-1",
        artifactRequestIds: ["delivery-html"],
        options: { ...baseOptions, outputs: ["html"] },
      }, {
        id: "strict-font-artifact",
        sourceId: "font-source",
        artifactRequestIds: ["delivery-font-html"],
        options: { ...baseOptions, strictFonts: true, outputs: ["html"] },
      }],
    };
    const rendered = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      env: {
        ...process.env,
        DOCXODUS_CHROMIUM_PATH: chromium.executablePath(),
        DOCXODUS_FONT_POLICY_PATH: hostFontPolicyPath,
      },
      input: framedProtocolRequest(request, [source, fontSource]),
      maxBuffer: 20 * 1024 * 1024,
    });
    assert.equal(rendered.status, 0, rendered.stderr.toString());
    assert.equal(rendered.stderr.byteLength, 0);
    const { control: response, blobs } = parseFrames(rendered.stdout);
    assert.equal(response.schemaVersion, 1);
    assert.equal(response.batches.length, 3);
    assert.equal(response.batches[0].id, "artifact-1");
    assert.equal(response.batches[1].id, "artifact-2");
    assert.equal(response.batches[2].id, "strict-font-artifact");
    assert.equal(response.artifacts.length, 9);
    const pdfBlob = blobs.find(({ descriptor }) => descriptor.kind === "pdf");
    const htmlBlob = blobs.find(({ descriptor }) => descriptor.kind === "html");
    assert.ok(pdfBlob.bytes.byteLength > 1_000);
    assert.match(htmlBlob.bytes.toString("utf8"), /^<!doctype html>/);
    for (const { descriptor, bytes } of blobs.filter(({ descriptor }) =>
      descriptor.kind === "renderReport")) {
      const report = JSON.parse(bytes.toString("utf8"));
      assert.deepEqual(report.bindings.artifactRequestIds,
        descriptor.batchId === "artifact-1"
          ? ["delivery-pdf"]
          : descriptor.batchId === "artifact-2"
            ? ["delivery-html"]
            : ["delivery-font-html"]);
      if (descriptor.batchId === "strict-font-artifact") {
        assert.ok(report.fonts.every((font) =>
          font.status === "resolved"
          && font.source === "attested"
          && font.resolvedFamily === "Docxodus Canvas Mono"));
        assert.equal(report.environment.verification, "nodeVerified");
      }
    }
  });
});
