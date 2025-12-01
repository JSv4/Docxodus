/**
 * Web Worker Tests for Docxodus
 *
 * These tests verify that the Web Worker implementation provides
 * non-blocking WASM operations while maintaining functional correctness.
 */

import { test, expect } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

// ESM compatibility: derive __dirname equivalent
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read test files from the TestFiles directory
const testFilesDir = path.join(__dirname, "../../TestFiles");

function readTestFile(relativePath: string): Uint8Array {
  const fullPath = path.join(testFilesDir, relativePath);
  return new Uint8Array(fs.readFileSync(fullPath));
}

test.describe("Docxodus Web Worker Tests", () => {
  test.beforeEach(async ({ page }) => {
    // Navigate to worker test harness
    await page.goto("/worker-test-harness.html");

    // Wait for test harness to load
    await page.waitForFunction(() => (window as any).DocxodusWorkerTests !== undefined, {
      timeout: 10000,
    });
  });

  test.describe("Worker Initialization", () => {
    test("isWorkerSupported returns true in browser", async ({ page }) => {
      const isSupported = await page.evaluate(() => {
        return (window as any).DocxodusWorkerTests.isSupported();
      });
      expect(isSupported).toBe(true);
    });

    test("worker can be created and initialized", async ({ page }) => {
      // Create worker - this may take a while as it loads WASM
      const result = await page.evaluate(async () => {
        try {
          await (window as any).createDocxodusWorker();
          return {
            ready: (window as any).DocxodusWorkerReady,
            active: (window as any).DocxodusWorkerTests.isActive(),
          };
        } catch (error: any) {
          return { error: error.message };
        }
      });

      expect(result.error).toBeUndefined();
      expect(result.ready).toBe(true);
      expect(result.active).toBe(true);

      console.log("Worker initialized successfully");
    }, { timeout: 60000 });

    test("worker can be terminated", async ({ page }) => {
      // Create and then terminate worker
      const result = await page.evaluate(async () => {
        await (window as any).createDocxodusWorker();
        const activeBefore = (window as any).DocxodusWorkerTests.isActive();
        (window as any).DocxodusWorkerTests.terminate();
        const activeAfter = (window as any).DocxodusWorkerTests.isActive();
        return { activeBefore, activeAfter };
      });

      expect(result.activeBefore).toBe(true);
      expect(result.activeAfter).toBe(false);

      console.log("Worker terminated successfully");
    }, { timeout: 60000 });
  });

  test.describe("Non-blocking Behavior", () => {
    test("UI remains responsive during conversion", async ({ page }) => {
      const bytes = readTestFile("HC006-Test-01.docx");

      // This test verifies that the main thread is not blocked during worker operations
      const result = await page.evaluate(async (bytesArray) => {
        // Create worker
        await (window as any).createDocxodusWorker();

        // Track UI responsiveness
        let animationFrameCount = 0;
        let conversionComplete = false;

        // Start counting animation frames
        const countFrames = () => {
          if (!conversionComplete) {
            animationFrameCount++;
            requestAnimationFrame(countFrames);
          }
        };
        requestAnimationFrame(countFrames);

        // Start conversion in worker
        const startTime = performance.now();
        const conversionPromise = (window as any).DocxodusWorkerTests.convertToHtml(bytesArray);

        // Wait for conversion
        const result = await conversionPromise;
        conversionComplete = true;

        const endTime = performance.now();
        const duration = endTime - startTime;

        return {
          success: !result.error,
          htmlLength: result.html?.length || 0,
          duration,
          animationFrameCount,
          // If frames were counted, the main thread wasn't blocked
          mainThreadResponsive: animationFrameCount > 0,
        };
      }, Array.from(bytes));

      expect(result.success).toBe(true);
      expect(result.htmlLength).toBeGreaterThan(100);
      expect(result.mainThreadResponsive).toBe(true);

      console.log(
        `Conversion took ${result.duration.toFixed(0)}ms, ` +
        `${result.animationFrameCount} animation frames fired (main thread responsive)`
      );
    }, { timeout: 60000 });

    test("multiple operations can be queued", async ({ page }) => {
      const bytes = readTestFile("HC006-Test-01.docx");

      const result = await page.evaluate(async (bytesArray) => {
        await (window as any).createDocxodusWorker();

        const startTime = performance.now();

        // Queue multiple operations
        const promises = [
          (window as any).DocxodusWorkerTests.convertToHtml(bytesArray),
          (window as any).DocxodusWorkerTests.getVersion(),
        ];

        const results = await Promise.all(promises);
        const endTime = performance.now();

        return {
          conversionSuccess: !results[0].error,
          htmlLength: results[0].html?.length || 0,
          versionSuccess: !!results[1].library,
          duration: endTime - startTime,
        };
      }, Array.from(bytes));

      expect(result.conversionSuccess).toBe(true);
      expect(result.htmlLength).toBeGreaterThan(100);
      expect(result.versionSuccess).toBe(true);

      console.log(`Multiple operations completed in ${result.duration.toFixed(0)}ms`);
    }, { timeout: 60000 });
  });

  test.describe("Conversion Operations", () => {
    test("convertDocxToHtml produces valid HTML", async ({ page }) => {
      const bytes = readTestFile("HC006-Test-01.docx");

      const result = await page.evaluate(async (bytesArray) => {
        await (window as any).createDocxodusWorker();
        return await (window as any).DocxodusWorkerTests.convertToHtml(bytesArray);
      }, Array.from(bytes));

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();
      expect(result.html.length).toBeGreaterThan(100);
      expect(result.html).toContain("<html");
      expect(result.html).toContain("</html>");

      console.log(`Converted document to ${result.html.length} bytes of HTML`);
    }, { timeout: 60000 });

    test("getVersion returns library info", async ({ page }) => {
      const result = await page.evaluate(async () => {
        await (window as any).createDocxodusWorker();
        return await (window as any).DocxodusWorkerTests.getVersion();
      });

      expect(result.error).toBeUndefined();
      expect(result.library).toBeDefined();
      expect(result.dotnetVersion).toBeDefined();
      expect(result.platform).toBeDefined();

      console.log(`Worker version: ${result.library}`);
    }, { timeout: 60000 });
  });

  test.describe("Comparison Operations", () => {
    test("compareDocuments produces valid redlined document", async ({ page }) => {
      const originalBytes = readTestFile("WC/WC001-Digits.docx");
      const modifiedBytes = readTestFile("WC/WC001-Digits-Mod.docx");

      const result = await page.evaluate(async ([original, modified]) => {
        await (window as any).createDocxodusWorker();
        return await (window as any).DocxodusWorkerTests.compareDocuments(original, modified);
      }, [Array.from(originalBytes), Array.from(modifiedBytes)]);

      expect(result.error).toBeUndefined();
      expect(result.docxBytes).toBeDefined();
      expect(result.docxBytes.length).toBeGreaterThan(0);

      // Verify it's a valid DOCX (starts with PK zip signature)
      expect(result.docxBytes[0]).toBe(0x50); // P
      expect(result.docxBytes[1]).toBe(0x4B); // K

      console.log(`Comparison produced ${result.docxBytes.length} byte redlined document`);
    }, { timeout: 60000 });

    test("compareDocumentsToHtml produces HTML with tracked changes", async ({ page }) => {
      const originalBytes = readTestFile("WC/WC001-Digits.docx");
      const modifiedBytes = readTestFile("WC/WC001-Digits-Mod.docx");

      const result = await page.evaluate(async ([original, modified]) => {
        await (window as any).createDocxodusWorker();
        return await (window as any).DocxodusWorkerTests.compareToHtml(original, modified);
      }, [Array.from(originalBytes), Array.from(modifiedBytes)]);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();
      expect(result.html.length).toBeGreaterThan(100);

      // Should contain tracked changes markup
      expect(result.html).toMatch(/<(ins|del)/);

      console.log(`Comparison HTML with tracked changes: ${result.html.length} bytes`);
    }, { timeout: 60000 });

    test("getRevisions extracts revisions from compared document", async ({ page }) => {
      const originalBytes = readTestFile("WC/WC001-Digits.docx");
      const modifiedBytes = readTestFile("WC/WC001-Digits-Mod.docx");

      // First compare to get a document with tracked changes
      const compareResult = await page.evaluate(async ([original, modified]) => {
        await (window as any).createDocxodusWorker();
        return await (window as any).DocxodusWorkerTests.compareDocuments(original, modified);
      }, [Array.from(originalBytes), Array.from(modifiedBytes)]);

      expect(compareResult.error).toBeUndefined();

      // Then extract revisions
      const result = await page.evaluate(async (docxBytes) => {
        return await (window as any).DocxodusWorkerTests.getRevisions(docxBytes);
      }, compareResult.docxBytes);

      expect(result.error).toBeUndefined();
      expect(result.revisions).toBeDefined();
      expect(result.revisions.length).toBeGreaterThan(0);

      // Check revision structure
      const firstRevision = result.revisions[0];
      expect(firstRevision.author).toBeDefined();
      expect(firstRevision.revisionType).toBeDefined();

      console.log(`Extracted ${result.revisions.length} revisions`);
    }, { timeout: 60000 });
  });

  test.describe("Error Handling", () => {
    test("handles invalid document gracefully", async ({ page }) => {
      const invalidBytes = new Uint8Array([0, 1, 2, 3, 4, 5]); // Not a valid DOCX

      const result = await page.evaluate(async (bytesArray) => {
        await (window as any).createDocxodusWorker();
        return await (window as any).DocxodusWorkerTests.convertToHtml(bytesArray);
      }, Array.from(invalidBytes));

      expect(result.error).toBeDefined();
      expect(result.error.message).toBeTruthy();

      console.log(`Error handling works: ${result.error.message}`);
    }, { timeout: 60000 });

    test("rejects requests after termination", async ({ page }) => {
      const bytes = readTestFile("HC006-Test-01.docx");

      const result = await page.evaluate(async (bytesArray) => {
        await (window as any).createDocxodusWorker();
        (window as any).DocxodusWorkerTests.terminate();

        // Try to convert after termination
        return await (window as any).DocxodusWorkerTests.convertToHtml(bytesArray);
      }, Array.from(bytes));

      expect(result.error).toBeDefined();

      console.log("Correctly rejects requests after termination");
    }, { timeout: 60000 });
  });
});
