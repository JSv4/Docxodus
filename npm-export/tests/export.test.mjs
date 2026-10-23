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
  DocxodusExportError,
  renderDocxArtifacts,
  renderDocxFile,
} from "../dist/index.js";
import { canonicalJson } from "../dist/canonical.js";
import { pdfLineText, joinHyphenatedLines } from "./pdf-line-text.mjs";
import {
  exactRowCell,
  exactRowParagraph,
  generateExactRowOverflowDocx,
  generateLongFootnoteDocx,
  generateMixedSectionDocx,
  generateTrackedRevisionDocx,
} from "./mixed-section-fixture.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const packageRoot = dirname(here);
const repositoryRoot = dirname(packageRoot);
const fixtures = join(repositoryRoot, "TestFiles");
const artifacts = join(packageRoot, "test-artifacts");
const successArtifacts = join(artifacts, "success");
const failureArtifacts = join(artifacts, "failure");
const cliEntry = join(packageRoot, "dist", "cli.js");
const hostEntry = join(packageRoot, "dist", "host.js");
const baseOptions = Object.freeze({
  reviewProfile: "final",
  commentProfile: "hidden",
  timeoutMs: 120_000,
});

let browser;

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function writeJson(path, value) {
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`);
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
      textWithLineBreaks: pdfLineText(content.items),
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

/**
 * One page's text three ways: in content-stream order, in tagged-structure order, and the rectangle
 * of every clip the page sets (as `[left, top, right, bottom]` in the page's content units).
 */
async function inspectPdfPageOrder(bytes, pageNumber) {
  const loading = getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, useSystemFonts: true });
  const pdf = await loading.promise;
  try {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent({ includeMarkedContent: true });
    const textByContentId = new Map();
    const open = [];
    const streamText = [];
    for (const item of content.items) {
      if (item.type === "beginMarkedContentProps" || item.type === "beginMarkedContent") {
        open.push(item.id);
      } else if (item.type === "endMarkedContent") {
        open.pop();
      } else if ("str" in item) {
        streamText.push(item.str);
        const id = open.at(-1);
        if (id !== undefined) textByContentId.set(id, (textByContentId.get(id) ?? "") + item.str);
      }
    }
    const roles = [];
    const structureText = [];
    const walk = (node, ancestors) => {
      if (node.type === "content") {
        const text = textByContentId.get(node.id);
        if (text?.trim()) {
          structureText.push(text);
          roles.push({ text: text.trim(), ancestors });
        }
        return;
      }
      for (const child of node.children ?? []) walk(child, [...ancestors, node.role]);
    };
    walk(await page.getStructTree(), []);
    const operators = await page.getOperatorList();
    const clipRects = [];
    for (let index = 0; index < operators.fnArray.length; index++) {
      if (operators.fnArray[index] !== OPS.clip && operators.fnArray[index] !== OPS.eoClip) continue;
      const path = operators.argsArray.slice(index + 1).find((_, offset) =>
        operators.fnArray[index + 1 + offset] === OPS.constructPath);
      if (path) clipRects.push(Array.from(path.at(-1)));
    }
    const normalize = (parts) => parts.join(" ").replace(/\s+/g, " ").trim();
    return { streamText: normalize(streamText), structureText: normalize(structureText), roles, clipRects };
  } finally {
    await loading.destroy();
  }
}

/**
 * Lay the standalone HTML out twice, with each exact-row SVG rewrite in place and then undone (its
 * clip box moved back with the converter's own overflow clip), and return every element whose
 * page-relative box differs. The rewrite must not move anything.
 */
async function exactRowRewriteLayoutShifts(html) {
  const context = await browser.newContext({ serviceWorkers: "block", bypassCSP: true });
  try {
    const page = await context.newPage();
    await page.setContent(html, { waitUntil: "load" });
    return await page.evaluate(() => {
      const snapshot = () => Array.from(document.querySelectorAll(".page-box")).flatMap((pageBox) => {
        const origin = pageBox.getBoundingClientRect();
        return Array.from(pageBox.querySelectorAll("[data-source-anchor-id], [data-docx-exact-row-cell]"))
          .map((element) => {
            const rect = element.getBoundingClientRect();
            return {
              id: element.getAttribute("data-source-anchor-id") ?? "clip",
              box: [rect.left - origin.left, rect.top - origin.top, rect.width, rect.height],
            };
          });
      });
      const rewritten = snapshot();
      const svgs = Array.from(document.querySelectorAll("svg[data-docx-exact-row-text]"));
      for (const svg of svgs) {
        const clip = svg.querySelector("[data-docx-exact-row-cell]");
        for (const property of ["width", "display", "overflow", "overflow-clip-margin"])
          clip.style.removeProperty(property);
        svg.replaceWith(clip);
      }
      const original = snapshot();
      return {
        undone: svgs.length,
        shifts: original.flatMap((entry, index) => {
          const other = rewritten[index];
          return entry.id === other.id && entry.box.every((value, axis) => value === other.box[axis])
            ? [] : [{ id: entry.id, original: entry.box, rewritten: other.box }];
        }),
      };
    });
  } finally {
    await context.close();
  }
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
  const viewer = `<!doctype html><meta charset="utf-8"><title>Docxodus #439/#440 test artifacts</title>
<style>body{font:15px system-ui;margin:2rem;max-width:70rem}li{margin:.55rem 0}iframe,object,img{width:100%;min-height:48rem;border:1px solid #bbb}</style>
<h1>Docxodus #439/#440 Node/PDF export evidence</h1>
<p>Generated by the end-to-end Node, CLI, Chromium, and PDF-parser test suite.</p>
<ul>
<li><a href="success/generated.pdf">Generated searchable PDF</a></li>
<li><a href="success/standalone.html">Standalone offline HTML</a></li>
<li><a href="success/page-map.json">PageMap</a></li>
<li><a href="success/render-report.json">Complete render report</a></li>
<li><a href="success/pdf-inspection.json">PDF text/tag inspection</a></li>
<li><a href="success/request-log.json">Offline reopen request log</a></li>
<li><a href="success/offline-reopen.png">Offline reopen screenshot</a></li>
<li><a href="success/hyperlinks.pdf">Hyperlink preservation PDF</a></li>
<li><a href="success/chart-vector.pdf">Vector-chart preservation PDF</a></li>
<li><a href="success/cli.pdf">CLI PDF</a></li>
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

// Blob frames carry exact raw bytes, not JSON. Encoding them like the control frame would
// send a {"type":"Buffer","data":[...]} rendering whose length no longer matches the
// declaration the host verifies against.
function framedBlob(bytes) {
  const header = Buffer.alloc(4);
  header.writeUInt32BE(bytes.byteLength);
  return Buffer.concat([header, bytes]);
}

function framedProtocolRequest(value, blobs = []) {
  return Buffer.concat([
    framedRequest(value),
    ...blobs.map((blob) => framedBlob(Buffer.from(blob))),
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
      convertDocxToPdf(source, {
        ...baseOptions,
        limits: { compressedDocxBytes: source.byteLength - 1 },
      }),
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_limit"
        && error.phase === "input_validation",
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
    // The export prints the margin or Word's balloons, which show margin comments only.
    for (const options of [{ commentColumnPlacement: "outside" }, { commentColumnPlacement: "balloons", commentProfile: "inline" }]) {
      await assert.rejects(
        convertDocxToPdf(source, { ...baseOptions, ...options }),
        (error) => error instanceof DocxodusExportError
          && error.code === "invalid_argument"
          && error.phase === "input_validation",
      );
    }
    await assert.rejects(
      convertDocxToPdf(source, {
        ...baseOptions,
        fontDirectories: ["/deployment/fonts"],
        fontLicenseAttestations: [{
          schemaVersion: 1,
          usage: "standalone-document-font-embedding",
          fileSha256: "a".repeat(64),
          embeddingPermitted: true,
          permittedOutputs: ["pdf"],
          subsettingPermitted: false,
          basis: "test-only attestation",
        }],
      }),
      // Before #442 this rejected `unsupported_runtime` because font directories were
      // refused outright. The runtime now attempts discovery, so a well-formed attestation
      // gets as far as resolving the root — and fails closed there, because
      // /deployment/fonts does not exist.
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_policy_failure"
        && error.phase === "font_loading",
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
    assert.equal(result.renderReport.options.reviewProfileAlreadyApplied, true);
    assert.equal(result.renderReport.environment.verification, "browserObserved");
    assert.deepEqual(result.renderReport.bindings.artifactRequestIds, []);
    assert.ok(result.renderReport.readiness.some((entry) =>
      entry.phase === "pdf_print" && entry.status === "complete"));
    // The host owns phases the browser materializer cannot see from inside the page.
    // Without these, a timeout in any of them reports a bare code and no pending work.
    for (const phase of ["browser_launch", "wasm_initialization", "output_verification", "cleanup"]) {
      assert.ok(
        result.renderReport.readiness.some((entry) =>
          entry.phase === phase && entry.status === "complete"),
        `readiness is missing a completed host-owned ${phase} phase`,
      );
    }
    // Host phases that ran before the report existed are prepended, not appended, so the
    // log stays in the order the work actually happened in.
    const readinessOrder = result.renderReport.readiness.map((entry) => entry.phase);
    assert.equal(readinessOrder[0], "browser_launch");
    assert.equal(readinessOrder.at(-1), "cleanup");
    assert.ok(
      readinessOrder.indexOf("docx_conversion") > readinessOrder.indexOf("browser_launch"),
      "browser-reported phases must follow the host launch that produced them",
    );

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
    await writeJson(join(successArtifacts, "pdf-inspection.json"), inspection);
    await writeJson(join(successArtifacts, "request-log.json"), requests);
    await writeFile(join(successArtifacts, "offline-reopen.png"), screenshot);
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
    });
    await writeFile(join(successArtifacts, "mixed-sections-scaled-viewer.png"), screenshot);
  });

  test("preserves one converter-produced footnote paragraph across HTML and PDF pages", {
    timeout: 180_000,
  }, async () => {
    // Long enough to continue across at least three pages even when note-only continuation pages
    // use 95% of the body band (issue #1119); 600 words fitted two.
    const words = 1200;
    const source = generateLongFootnoteDocx(words);
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

    // Match on word boundaries, never substrings: `footnote-1-1-1` occurs inside
    // `footnote-1-1-10`, so a bare indexOf reports every low-numbered word twice.
    // PDF.js reports a hard hyphen at a line end in a separate item from the
    // remainder of the word. Join only actual hasEOL boundaries; ordinary spaces,
    // missing glyphs, and duplicates must still fail the complete inventory check.
    const noteText = joinHyphenatedLines(inspection.pages.map((page) => page.textWithLineBreaks).join("\n"));
    let priorPdfOffset = -1;
    for (let index = 1; index <= words; index++) {
      const token = `footnote-1-1-${index}`;
      assert.equal(result.html.match(new RegExp(`\\b${token}\\b`, "g"))?.length, 1);
      const pdfMatches = [...noteText.matchAll(new RegExp(`\\b${token}\\b`, "g"))];
      assert.equal(pdfMatches.length, 1, `${token} must appear exactly once in the PDF text`);
      assert.ok(pdfMatches[0].index > priorPdfOffset, `${token} out of order in PDF text`);
      priorPdfOffset = pdfMatches[0].index;
    }

    await writeFile(join(successArtifacts, "long-footnote.pdf"), result.pdf);
    await writeFile(join(successArtifacts, "long-footnote.html"), result.html);
    await writeJson(join(successArtifacts, "long-footnote-page-map.json"), result.pageMap);
    await writeJson(join(successArtifacts, "long-footnote-pdf-inspection.json"), inspection);
  });

  // Issue #1140: Chromium's PDF writer keeps a glyph only when it meets the bounding rectangle of the
  // clip it is drawn under, so a line that an exact-height row cuts through used to survive as the
  // letters whose ascenders reach the row ("fl h l"). The clipped line must instead appear whole,
  // once, in reading order, in both the content stream and the tagged structure.
  test("keeps text an exact-height row clips whole and in order in the PDF text layer", {
    timeout: 180_000,
  }, async () => {
    const expected = "Opening ledger paragraph. Birch exact record. Overflow Birch continuation line. "
      + "Neighbour cell. Following row text. Second column. Closing ledger paragraph.";
    const result = await renderDocxArtifacts(generateExactRowOverflowDocx(), {
      ...baseOptions,
      browser,
      outputs: ["html", "pdf"],
    });
    const order = await inspectPdfPageOrder(result.pdf, 1);
    assert.equal(order.streamText, expected);
    assert.equal(order.structureText, expected);
    const hidden = order.roles.find((entry) => entry.text === "Overflow Birch continuation line.");
    assert.ok(hidden, "the clipped paragraph is tagged");
    assert.equal(hidden.ancestors.at(-2), "P", `clipped paragraph tagged as ${hidden.ancestors.join(" > ")}`);
    assert.ok(hidden.ancestors.includes("TD") && !hidden.ancestors.includes("Figure"));
    // Only the cell whose content overflows its row is rewritten.
    assert.equal(result.html.match(/data-docx-exact-row-text/g)?.length, 1);

    const fitting = await renderDocxArtifacts(generateExactRowOverflowDocx({ rowHeight: 720 }), {
      ...baseOptions,
      browser,
      outputs: ["html", "pdf"],
    });
    assert.equal(fitting.html.match(/data-docx-exact-row-text/g), null);
    assert.equal((await inspectPdfPageOrder(fitting.pdf, 1)).streamText, expected);

    // At-least line spacing leaves every run unpositioned, so all of the page's text paints with the
    // in-flow content; the rewritten cell must paint there too, not after the rest of the page.
    const unpositioned = await renderDocxArtifacts(
      generateExactRowOverflowDocx({ spacing: 'w:line="300" w:lineRule="atLeast"' }),
      { ...baseOptions, browser, outputs: ["html", "pdf"] },
    );
    assert.equal(unpositioned.html.match(/data-docx-exact-row-text/g)?.length, 1);
    const unpositionedOrder = await inspectPdfPageOrder(unpositioned.pdf, 1);
    assert.equal(unpositionedOrder.streamText, expected);
    assert.equal(unpositionedOrder.structureText, expected);

    // A cell with text in both paint phases (an exactly spaced first line, which is offset, over an
    // unpositioned hidden line) on a page whose text is unpositioned paints where its positioned text
    // did, after the page's in-flow text, but reads in tree order inside.
    const exactly = 'w:line="280" w:lineRule="exact"';
    const atLeast = 'w:line="300" w:lineRule="atLeast"';
    const mixed = await renderDocxArtifacts(generateExactRowOverflowDocx({
      spacing: atLeast,
      firstRow: `<w:tr><w:trPr><w:trHeight w:val="360" w:hRule="exact"/></w:trPr>`
        + exactRowCell(exactRowParagraph("Birch exact record.", "", exactly)
          + exactRowParagraph("Overflow Birch continuation line.", "", atLeast))
        + exactRowCell(exactRowParagraph("Neighbour cell.", "", atLeast)) + "</w:tr>",
    }), { ...baseOptions, browser, outputs: ["html", "pdf"] });
    const mixedOrder = await inspectPdfPageOrder(mixed.pdf, 1);
    assert.equal(mixedOrder.structureText, expected);
    assert.match(mixedOrder.streamText, /Birch exact record\. Overflow Birch continuation line\./);
    assert.equal(mixedOrder.streamText.match(/Overflow Birch continuation line\./g)?.length, 1);

    await writeFile(join(successArtifacts, "exact-row-overflow.pdf"), result.pdf);
    await writeFile(join(successArtifacts, "exact-row-overflow.html"), result.html);
  });

  // Layout overflow is not hidden text: a line box includes the font's whole ascent and descent and a
  // paragraph its spacing, so a one-line row a little shorter than its line, or space after, overflows
  // the clip box while every glyph still shows.
  test("leaves exact-height rows that hide no text as they were", { timeout: 180_000 }, async () => {
    const spacedAfter = '<w:p><w:pPr><w:spacing w:before="0" w:after="200"/></w:pPr>'
      + "<w:r><w:t>Birch exact record.</w:t></w:r></w:p>";
    for (const [name, rowHeight, body] of [
      ["one line in a 13pt row", 260, exactRowParagraph("Birch exact record.")],
      ["one line in a 14pt row", 280, exactRowParagraph("Birch exact record.")],
      ["one line with space after", 400, spacedAfter],
    ]) {
      const result = await renderDocxArtifacts(generateExactRowOverflowDocx({
        firstRow: `<w:tr><w:trPr><w:trHeight w:val="${rowHeight}" w:hRule="exact"/></w:trPr>`
          + exactRowCell(body) + exactRowCell(exactRowParagraph("Neighbour cell.")) + "</w:tr>",
      }), { ...baseOptions, browser, outputs: ["html", "pdf"] });
      assert.equal(result.html.match(/data-docx-exact-row-text/g), null, `${name}: rewritten`);
      assert.equal((await inspectPdfPageOrder(result.pdf, 1)).streamText,
        "Opening ledger paragraph. Birch exact record. Neighbour cell. Following row text. Second column. "
          + "Closing ledger paragraph.", name);
    }
  });

  // A Word bookmark's name becomes an HTML id verbatim, and url(#id) resolves to the first element with
  // that id: the clip the rewrite mints must not be one the document already uses.
  test("mints an exact-row clip id no bookmark already holds", { timeout: 180_000 }, async () => {
    const result = await renderDocxArtifacts(generateExactRowOverflowDocx({
      firstRow: "<w:tr>" + exactRowCell('<w:p><w:bookmarkStart w:id="5" w:name="docx-exact-row-clip-0"/>'
        + '<w:r><w:t>Bookmarked row.</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>')
        + exactRowCell(exactRowParagraph("Bookmark neighbour.")) + "</w:tr>"
        + `<w:tr><w:trPr><w:trHeight w:val="360" w:hRule="exact"/></w:trPr>`
        + exactRowCell(exactRowParagraph("Birch exact record.")
          + exactRowParagraph("Overflow Birch continuation line."))
        + exactRowCell(exactRowParagraph("Neighbour cell.")) + "</w:tr>",
    }), { ...baseOptions, browser, outputs: ["html", "pdf"] });
    const reference = /clip-path="url\(#([^)]+)\)"/.exec(result.html)?.[1];
    assert.ok(reference, "the rewritten cell is clipped by reference");
    const holders = [...result.html.matchAll(new RegExp(`<([a-zA-Z]+)[^>]*\\sid="${reference}"`, "g"))];
    assert.deepEqual(holders.map((match) => match[1]), ["clipPath"]);
    assert.match((await inspectPdfPageOrder(result.pdf, 1)).streamText,
      /Birch exact record\. Overflow Birch continuation line\. Neighbour cell\./);
  });

  // The rewrite that keeps clipped exact-row text whole moves the clipped content into an inline SVG,
  // which Chromium paints at a whole pixel; the export takes the fraction back so the content and the
  // row's clip land exactly where the row's own overflow clip put them. A cell beside it that is not
  // rewritten is the reference: same row, same clip, same first line. The margins and the A4 page 2
  // put the row at a fractional pixel on both axes; a pinned Chromium that snapped differently would
  // move the rewritten cell by that fraction.
  test("places clipped exact-row content where the row's own clip put it", {
    timeout: 180_000,
  }, async () => {
    const nestedTable = `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders>`
      + `<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/>`
      + `</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>`
      + `<w:tr>${exactRowCell(exactRowParagraph("Nested first line."))}</w:tr>`
      + `<w:tr>${exactRowCell(exactRowParagraph("Nested second line."))}</w:tr></w:tbl>`;
    const geometries = [
      { name: "letter", options: {}, page: 1 },
      {
        name: "nested table",
        options: {
          // 32pt holds the nested table's two lines and borders but not the trailing line, so the cell
          // is clipped. (At 43pt it only overflowed while a table first in a cell took a 7.5pt top
          // margin of its own, which Word does not give it: #1244.)
          firstRow: `<w:tr><w:trPr><w:trHeight w:val="640" w:hRule="exact"/></w:trPr>`
            + exactRowCell(nestedTable + exactRowParagraph("Trailing nested cell line."))
            + exactRowCell(exactRowParagraph("Neighbour cell.")) + "</w:tr>",
        },
        page: 1,
        // The nested table's border puts its first line lower than the neighbour's.
        sameFirstBaseline: false,
      },
      { name: "letter, middle-aligned", options: { vAlign: "center" }, page: 1 },
      {
        name: "A4 page 2, odd margins, bottom-aligned",
        options: {
          pageWidth: 11906,
          pageHeight: 16838,
          margins: { top: 1001, right: 1117, bottom: 1237, left: 1333 },
          leadingParagraphs: 61,
          vAlign: "bottom",
        },
        page: 2,
      },
    ];
    for (const { name, options, page, sameFirstBaseline = true } of geometries) {
      const result = await renderDocxArtifacts(generateExactRowOverflowDocx(options), {
        ...baseOptions,
        browser,
        outputs: ["html", "pdf"],
      });
      assert.equal(result.pageCount, page, `${name}: the table lands on page ${page}`);
      if (sameFirstBaseline) {
        const inspection = await inspectPdf(result.pdf);
        const run = (text) => inspection.pages[page - 1].textRuns.find((entry) => entry.text === text);
        const clipped = run("Birch exact record.");
        const neighbour = run("Neighbour cell.");
        assert.ok(clipped && neighbour, `${name}: both first lines are in the PDF`);
        // 3600 twips separate the two cells' text; their baselines are the row's.
        assert.ok(Math.abs(clipped.transform[5] - neighbour.transform[5]) <= 0.01,
          `${name}: baseline ${clipped.transform[5]} != ${neighbour.transform[5]}`);
        assert.ok(Math.abs(neighbour.transform[4] - clipped.transform[4] - 180) <= 0.01,
          `${name}: cell offset ${neighbour.transform[4] - clipped.transform[4]} != 180pt`);
      }

      // The rewritten cell's clip is the neighbour's, 180pt (750 device units at 300 dpi) to the left.
      const { clipRects } = await inspectPdfPageOrder(result.pdf, page);
      const height = (rect) => rect[3] - rect[1];
      const neighbourClip = clipRects.find((rect) =>
        clipRects.some((other) => Math.abs(rect[0] - other[0] - 750) <= 0.01
          && Math.abs(rect[1] - other[1]) <= 0.01 && Math.abs(height(rect) - height(other)) <= 0.01));
      assert.ok(neighbourClip, `${name}: no clip matches the neighbour cell's: ${JSON.stringify(clipRects)}`);

      // Nothing on the page moves, not even by a layout unit: undoing the rewrite lays out the same.
      const { undone, shifts } = await exactRowRewriteLayoutShifts(result.html);
      assert.equal(undone, 1, `${name}: one cell was rewritten`);
      assert.deepEqual(shifts, [], `${name}: the rewrite moved content`);
    }
  });

  test("PDF text extraction is predictable for each review profile", { timeout: 180_000 }, async () => {
    // The markup profile is the only one that draws inserted and deleted content
    // side by side, so extraction order is part of its contract: tokens come out
    // in document order, deleted content included. The derived profiles must
    // instead prove the losing side of each revision never reaches the PDF text.
    const source = generateTrackedRevisionDocx();
    const expectations = [
      { reviewProfile: "markup", present: ["Before", "removed", "added", "after."], absent: [] },
      { reviewProfile: "final", present: ["Before", "added", "after."], absent: ["removed"] },
      { reviewProfile: "original", present: ["Before", "removed", "after."], absent: ["added"] },
    ];
    for (const { reviewProfile, present, absent } of expectations) {
      const result = await convertDocxToPdf(source, {
        reviewProfile,
        commentProfile: "hidden",
        timeoutMs: 120_000,
        browser,
      });
      const text = (await inspectPdf(result.pdf)).searchableText;
      let priorOffset = -1;
      for (const token of present) {
        const pattern = new RegExp(`\\b${token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`, "g");
        const matches = [...text.matchAll(pattern)];
        assert.equal(matches.length, 1, `${token} must appear exactly once under ${reviewProfile}`);
        assert.ok(matches[0].index > priorOffset, `${token} out of order under ${reviewProfile}`);
        priorOffset = matches[0].index;
      }
      for (const token of absent) {
        assert.doesNotMatch(text, new RegExp(`\\b${token}\\b`), `${token} must not extract under ${reviewProfile}`);
      }
    }
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
    assert.deepEqual(await currentRuntimeDirectories(), runtimeBefore);
  });

  test("framed host rejects local executable authority and returns keyed PDF results", { timeout: 180_000 }, async () => {
    const source = await readFile(join(fixtures, "CA", "CA001-Plain.docx"));
    const sourceDescriptor = {
      id: "source-1",
      byteLength: source.byteLength,
      sha256: digest(source),
      mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    };
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
          },
        }],
      }, [source]),
      maxBuffer: 10 * 1024 * 1024,
    });
    assert.equal(forbidden.status, 0);
    assert.match(parseFrames(forbidden.stdout).control.fatal.message,
      /unknown fields: browserExecutablePath/);

    const request = {
      schemaVersion: 1,
      sources: [sourceDescriptor],
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
      }],
    };
    const rendered = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      env: { ...process.env, DOCXODUS_CHROMIUM_PATH: chromium.executablePath() },
      input: framedProtocolRequest(request, [source]),
      maxBuffer: 20 * 1024 * 1024,
    });
    assert.equal(rendered.status, 0, rendered.stderr.toString());
    assert.equal(rendered.stderr.byteLength, 0);
    const { control: response, blobs } = parseFrames(rendered.stdout);
    assert.equal(response.schemaVersion, 1);
    // A failed render answers with a fatal envelope instead of batches. Surface that message
    // rather than dereferencing undefined, which would hide why the host gave up.
    assert.ok(response.batches, `host returned no batches: ${JSON.stringify(response.fatal)}`);
    assert.equal(response.batches.length, 2);
    assert.equal(response.batches[0].id, "artifact-1");
    assert.equal(response.batches[1].id, "artifact-2");
    assert.equal(response.artifacts.length, 6);
    const pdfBlob = blobs.find(({ descriptor }) => descriptor.kind === "pdf");
    const htmlBlob = blobs.find(({ descriptor }) => descriptor.kind === "html");
    assert.ok(pdfBlob.bytes.byteLength > 1_000);
    assert.match(htmlBlob.bytes.toString("utf8"), /^<!doctype html>/);
    for (const { descriptor, bytes } of blobs.filter(({ descriptor }) =>
      descriptor.kind === "renderReport")) {
      const report = JSON.parse(bytes.toString("utf8"));
      assert.deepEqual(report.bindings.artifactRequestIds,
        descriptor.batchId === "artifact-1" ? ["delivery-pdf"] : ["delivery-html"]);
    }
  });
});
