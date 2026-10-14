import { test, expect, Page } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, "../../TestFiles");

/**
 * Performance profiling tests for WASM document conversion.
 *
 * Run with: npx playwright test profiling.spec.ts --project=chromium
 */

// Helper to format milliseconds nicely
function formatMs(ms: number): string {
  if (ms < 1) return `${(ms * 1000).toFixed(1)}μs`;
  if (ms < 1000) return `${ms.toFixed(2)}ms`;
  return `${(ms / 1000).toFixed(2)}s`;
}

// Read a test file as Uint8Array
function readTestFile(filename: string): Uint8Array {
  const filePath = path.join(TEST_FILES_DIR, filename);

  if (!fs.existsSync(filePath)) {
    throw new Error(`Test file not found: ${filePath}`);
  }

  return new Uint8Array(fs.readFileSync(filePath));
}

// Wait for WASM to be ready
async function waitForWasm(page: Page, timeout = 30000): Promise<void> {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout,
  });
}

test.describe("WASM Performance Profiling", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/test-harness.html");
    await waitForWasm(page);
  });

  test("profile document conversion", async ({ page }) => {
    const bytes = readTestFile("HC007-Test-02.docx");

    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const results = await page.evaluate(async () => {
      const bytes = (window as any).testDocxBytes;
      const Docxodus = (window as any).Docxodus;
      const timings: Record<string, number> = {};

      // 1. GetDocumentMetadata
      let start = performance.now();
      const metadataResult = Docxodus.DocumentConverter.GetDocumentMetadata(bytes);
      timings["GetDocumentMetadata"] = performance.now() - start;

      const metadata = JSON.parse(metadataResult);

      // 2. ConvertDocxToHtml (no pagination)
      start = performance.now();
      Docxodus.DocumentConverter.ConvertDocxToHtml(bytes);
      timings["ConvertDocxToHtml"] = performance.now() - start;

      // 3. ConvertDocxToHtmlWithPagination
      start = performance.now();
      const paginatedHtml =
        Docxodus.DocumentConverter.ConvertDocxToHtmlWithPagination(
          bytes,
          "Document",
          "docx-",
          true,
          "",
          -1,
          "comment-",
          1,
          1.0,
          "page-"
        );
      timings["ConvertDocxToHtmlWithPagination"] = performance.now() - start;

      // 4. ConvertDocxToHtmlComplete
      start = performance.now();
      Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        bytes,
        "Document",
        "docx-",
        true,
        "",
        -1,
        "comment-",
        1,
        1.0,
        "page-",
        false,
        0,
        "annot-",
        true,
        true,
        false,
        true,
        true
      );
      timings["ConvertDocxToHtmlComplete"] = performance.now() - start;

      // 5. DOM insertion measurement
      const container = document.createElement("div");
      container.style.position = "absolute";
      container.style.left = "-9999px";
      document.body.appendChild(container);

      start = performance.now();
      container.innerHTML = paginatedHtml;
      timings["DOM_insertion"] = performance.now() - start;

      document.body.removeChild(container);

      return {
        documentSize: bytes.length,
        paragraphCount: metadata.TotalParagraphs || 0,
        tableCount: metadata.TotalTables || 0,
        estimatedPageCount: metadata.EstimatedPageCount || 1,
        timings,
      };
    });

    console.log("\n" + "═".repeat(70));
    console.log(`PROFILING RESULTS: HC007-Test-02.docx`);
    console.log("═".repeat(70));
    console.log(`Document size:     ${(results.documentSize / 1024).toFixed(1)} KB`);
    console.log(`Paragraphs:        ${results.paragraphCount}`);
    console.log(`Tables:            ${results.tableCount}`);
    console.log(`Estimated pages:   ${results.estimatedPageCount}`);
    console.log("─".repeat(70));

    for (const [name, time] of Object.entries(results.timings)) {
      console.log(`${name.padEnd(35)} ${formatMs(time as number).padStart(12)}`);
    }

    expect(results.timings.GetDocumentMetadata).toBeGreaterThan(0);
    expect(results.timings.ConvertDocxToHtml).toBeGreaterThan(0);
  });

  test("content-visibility:auto approach (browser-native optimization)", async ({ page }) => {
    const bytes = readTestFile("HC007-Test-02.docx");

    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const results = await page.evaluate(async () => {
      const bytes = (window as any).testDocxBytes;
      const Docxodus = (window as any).Docxodus;
      const timings: Record<string, number> = {};

      // Render full document
      let start = performance.now();
      const fullHtml = Docxodus.DocumentConverter.ConvertDocxToHtmlWithPagination(
        bytes, "Document", "docx-", true, "", -1, "comment-", 1, 1.0, "page-"
      );
      timings["WASM render"] = performance.now() - start;

      // Normal DOM insertion
      const normalContainer = document.createElement('div');
      normalContainer.style.cssText = 'position:absolute;left:-9999px;width:816px;height:600px;overflow:auto;';
      document.body.appendChild(normalContainer);

      start = performance.now();
      normalContainer.innerHTML = fullHtml;
      timings["Normal: DOM insertion"] = performance.now() - start;

      start = performance.now();
      void normalContainer.offsetHeight;
      timings["Normal: Layout"] = performance.now() - start;

      // With content-visibility: auto
      const cvContainer = document.createElement('div');
      cvContainer.style.cssText = 'position:absolute;left:-9999px;width:816px;height:600px;overflow:auto;';
      document.body.appendChild(cvContainer);

      const style = document.createElement('style');
      style.textContent = `
        .cv-container p, .cv-container table, .cv-container div {
          content-visibility: auto;
          contain-intrinsic-size: 0 50px;
        }
      `;
      document.head.appendChild(style);
      cvContainer.classList.add('cv-container');

      start = performance.now();
      cvContainer.innerHTML = fullHtml;
      timings["ContentVis: DOM insertion"] = performance.now() - start;

      start = performance.now();
      void cvContainer.offsetHeight;
      timings["ContentVis: Layout"] = performance.now() - start;

      document.body.removeChild(normalContainer);
      document.body.removeChild(cvContainer);
      document.head.removeChild(style);

      return { timings };
    });

    console.log("\n" + "═".repeat(70));
    console.log("CONTENT-VISIBILITY: AUTO TEST");
    console.log("═".repeat(70));
    console.log(`WASM render:              ${formatMs(results.timings["WASM render"])}`);
    console.log("");
    console.log("NORMAL DOM:");
    console.log(`  DOM insertion:          ${formatMs(results.timings["Normal: DOM insertion"])}`);
    console.log(`  Layout:                 ${formatMs(results.timings["Normal: Layout"])}`);
    console.log("");
    console.log("WITH content-visibility: auto:");
    console.log(`  DOM insertion:          ${formatMs(results.timings["ContentVis: DOM insertion"])}`);
    console.log(`  Layout:                 ${formatMs(results.timings["ContentVis: Layout"])}`);
    console.log("═".repeat(70));
  });

  test("measure WASM initialization overhead", async ({ page }) => {
    const initTimes: number[] = [];

    for (let i = 0; i < 3; i++) {
      const startTime = Date.now();
      await page.goto("/test-harness.html");
      await waitForWasm(page);
      initTimes.push(Date.now() - startTime);
    }

    console.log("\n" + "═".repeat(50));
    console.log("WASM INITIALIZATION TIME");
    console.log("═".repeat(50));
    console.log(`Run 1: ${formatMs(initTimes[0])}`);
    console.log(`Run 2: ${formatMs(initTimes[1])}`);
    console.log(`Run 3: ${formatMs(initTimes[2])}`);
    console.log(`Average: ${formatMs(initTimes.reduce((a, b) => a + b, 0) / 3)}`);
    console.log("═".repeat(50));
  });
});
