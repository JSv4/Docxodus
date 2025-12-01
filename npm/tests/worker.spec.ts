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

  // ============================================================
  // Visual Lazy Loading Tests (Issue #44 Phase 3)
  // ============================================================
  test.describe("Visual Lazy Loading", () => {
    test("demonstrates lazy loading with visual page placeholders", async ({ page }) => {
      const bytes = readTestFile("HC001-5DayTourPlanTemplate.docx");

      // This test visually demonstrates the lazy loading workflow:
      // 1. Get metadata first (fast) to create page shells
      // 2. Show loading placeholders with correct dimensions
      // 3. Load full content and render into the page shells

      const result = await page.evaluate(async (bytesArray) => {
        const docBytes = new Uint8Array(bytesArray);
        const timeline: string[] = [];
        const startTime = performance.now();

        // Create visual container for the document viewer
        const viewer = document.createElement('div');
        viewer.id = 'doc-viewer';
        viewer.style.cssText = `
          width: 100%;
          max-width: 800px;
          margin: 20px auto;
          background: #f5f5f5;
          padding: 20px;
          font-family: Arial, sans-serif;
        `;
        document.body.appendChild(viewer);

        // Status display
        const status = document.createElement('div');
        status.id = 'load-status';
        status.style.cssText = 'padding: 10px; background: #e3f2fd; margin-bottom: 20px; border-radius: 4px;';
        status.textContent = 'Initializing worker...';
        viewer.appendChild(status);

        // Page container
        const pageContainer = document.createElement('div');
        pageContainer.id = 'page-container';
        pageContainer.style.cssText = 'display: flex; flex-direction: column; gap: 20px; align-items: center;';
        viewer.appendChild(pageContainer);

        // Step 1: Initialize worker
        await (window as any).createDocxodusWorker();
        timeline.push(`${(performance.now() - startTime).toFixed(0)}ms: Worker initialized`);
        status.textContent = 'Getting document metadata...';

        // Step 2: Get metadata (fast operation)
        const metaStart = performance.now();
        const metaResult = await (window as any).DocxodusWorkerTests.getDocumentMetadata(Array.from(docBytes));
        const metaTime = performance.now() - metaStart;
        timeline.push(`${(performance.now() - startTime).toFixed(0)}ms: Metadata received (${metaTime.toFixed(0)}ms)`);

        if (metaResult.error) {
          return { error: metaResult.error, timeline };
        }

        const metadata = metaResult.metadata;
        status.textContent = `Found ${metadata.sections.length} section(s), ~${metadata.estimatedPageCount} pages. Creating placeholders...`;

        // Step 3: Create page placeholders based on metadata
        const placeholders: HTMLElement[] = [];
        for (let i = 0; i < metadata.estimatedPageCount; i++) {
          const sectionIndex = Math.min(i, metadata.sections.length - 1);
          const section = metadata.sections[sectionIndex];

          const placeholder = document.createElement('div');
          placeholder.className = 'page-placeholder';
          placeholder.dataset.pageNumber = String(i + 1);
          placeholder.dataset.sectionIndex = String(sectionIndex);

          // Use metadata dimensions (scale down for display)
          const scale = 0.5;
          const widthPx = section.pageWidthPt * scale * (96 / 72);
          const heightPx = section.pageHeightPt * scale * (96 / 72);

          placeholder.style.cssText = `
            width: ${widthPx}px;
            height: ${heightPx}px;
            background: linear-gradient(135deg, #e0e0e0 25%, #f0f0f0 50%, #e0e0e0 75%);
            background-size: 20px 20px;
            animation: shimmer 1.5s infinite;
            border: 1px solid #ccc;
            border-radius: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #666;
            font-size: 14px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          `;
          placeholder.innerHTML = `<span>Page ${i + 1}<br><small>Loading...</small></span>`;

          pageContainer.appendChild(placeholder);
          placeholders.push(placeholder);
        }

        // Add shimmer animation
        const style = document.createElement('style');
        style.textContent = `
          @keyframes shimmer {
            0% { background-position: -20px 0; }
            100% { background-position: 20px 0; }
          }
        `;
        document.head.appendChild(style);

        timeline.push(`${(performance.now() - startTime).toFixed(0)}ms: ${placeholders.length} placeholders created`);

        // Allow a brief moment for placeholders to be visible
        await new Promise(r => setTimeout(r, 500));

        // Capture placeholder state
        const placeholderState = {
          count: placeholders.length,
          firstPlaceholder: {
            width: placeholders[0]?.offsetWidth,
            height: placeholders[0]?.offsetHeight,
            visible: placeholders[0]?.offsetParent !== null
          }
        };

        status.textContent = 'Loading full document content...';

        // Step 4: Load full HTML content via worker
        const convStart = performance.now();
        const htmlResult = await (window as any).DocxodusWorkerTests.convertToHtmlWithPagination(
          Array.from(docBytes), 1, 0.5
        );
        const convTime = performance.now() - convStart;
        timeline.push(`${(performance.now() - startTime).toFixed(0)}ms: HTML conversion complete (${convTime.toFixed(0)}ms)`);

        if (htmlResult.error) {
          return { error: htmlResult.error, timeline, placeholderState };
        }

        // Step 5: Replace placeholders with actual content
        status.textContent = 'Rendering content...';

        // Parse the HTML
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlResult.html, 'text/html');

        // Clear placeholders and insert actual content
        pageContainer.innerHTML = '';

        // Find the document body content
        const docBody = doc.querySelector('.page-section') || doc.body;

        // Create a styled container for the actual content
        const contentWrapper = document.createElement('div');
        contentWrapper.id = 'rendered-content';
        contentWrapper.style.cssText = `
          background: white;
          padding: 20px;
          border-radius: 4px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.15);
          max-width: 100%;
          overflow: auto;
        `;

        // Copy styles from parsed document
        const styles = doc.querySelectorAll('style');
        styles.forEach(s => contentWrapper.appendChild(s.cloneNode(true)));

        // Copy body content
        contentWrapper.innerHTML += docBody.innerHTML;
        pageContainer.appendChild(contentWrapper);

        timeline.push(`${(performance.now() - startTime).toFixed(0)}ms: Content rendered`);
        status.textContent = `Document loaded successfully! (${metadata.totalParagraphs} paragraphs)`;
        status.style.background = '#c8e6c9';

        // Verify rendered content
        const renderedState = {
          hasContent: contentWrapper.innerHTML.length > 0,
          paragraphCount: contentWrapper.querySelectorAll('p').length,
          hasStyles: contentWrapper.querySelectorAll('style').length > 0,
          contentHeight: contentWrapper.offsetHeight,
          isVisible: contentWrapper.offsetParent !== null
        };

        return {
          success: true,
          timeline,
          metadata: {
            sections: metadata.sections.length,
            totalParagraphs: metadata.totalParagraphs,
            estimatedPages: metadata.estimatedPageCount
          },
          placeholderState,
          renderedState,
          timing: {
            metadataMs: metaTime,
            conversionMs: convTime,
            totalMs: performance.now() - startTime
          }
        };
      }, Array.from(bytes));

      // Assertions
      expect(result.error).toBeUndefined();
      expect(result.success).toBe(true);

      // Verify placeholder phase worked
      expect(result.placeholderState.count).toBeGreaterThan(0);
      expect(result.placeholderState.firstPlaceholder.visible).toBe(true);
      expect(result.placeholderState.firstPlaceholder.width).toBeGreaterThan(0);
      expect(result.placeholderState.firstPlaceholder.height).toBeGreaterThan(0);

      // Verify rendered content
      expect(result.renderedState.hasContent).toBe(true);
      expect(result.renderedState.paragraphCount).toBeGreaterThan(0);
      expect(result.renderedState.isVisible).toBe(true);

      // Metadata should be significantly faster than full conversion
      expect(result.timing.metadataMs).toBeLessThan(result.timing.conversionMs);

      // Log timeline
      console.log('\n=== Lazy Loading Timeline ===');
      result.timeline.forEach((entry: string) => console.log(entry));
      console.log(`\nMetadata: ${result.timing.metadataMs.toFixed(0)}ms`);
      console.log(`Conversion: ${result.timing.conversionMs.toFixed(0)}ms`);
      console.log(`Total: ${result.timing.totalMs.toFixed(0)}ms`);
      console.log(`\nPlaceholders created: ${result.placeholderState.count}`);
      console.log(`Paragraphs rendered: ${result.renderedState.paragraphCount}`);
    }, { timeout: 120000 });

    test("visual verification with screenshot", async ({ page }) => {
      const bytes = readTestFile("HC006-Test-01.docx");

      // Set viewport for consistent screenshots
      await page.setViewportSize({ width: 1024, height: 768 });

      await page.evaluate(async (bytesArray) => {
        const docBytes = new Uint8Array(bytesArray);

        // Initialize worker first (before modifying DOM)
        await (window as any).createDocxodusWorker();

        // Now safe to clear body
        document.body.innerHTML = '';
        document.body.style.cssText = 'margin: 0; padding: 20px; background: #fafafa;';

        const viewer = document.createElement('div');
        viewer.id = 'lazy-viewer';
        viewer.style.cssText = 'max-width: 900px; margin: 0 auto;';
        document.body.appendChild(viewer);

        // Get metadata
        const metaResult = await (window as any).DocxodusWorkerTests.getDocumentMetadata(Array.from(docBytes));
        const metadata = metaResult.metadata;

        // Create header showing metadata
        const header = document.createElement('div');
        header.style.cssText = `
          background: #1976d2;
          color: white;
          padding: 15px 20px;
          border-radius: 8px 8px 0 0;
          margin-bottom: 0;
        `;
        header.innerHTML = `
          <h2 style="margin: 0 0 10px 0; font-size: 18px;">Lazy Loading Demo</h2>
          <div style="font-size: 13px; opacity: 0.9;">
            Sections: ${metadata.sections.length} |
            Paragraphs: ${metadata.totalParagraphs} |
            Est. Pages: ${metadata.estimatedPageCount} |
            Page Size: ${metadata.sections[0]?.pageWidthPt?.toFixed(0)}×${metadata.sections[0]?.pageHeightPt?.toFixed(0)}pt
          </div>
        `;
        viewer.appendChild(header);

        // Create page preview area
        const previewArea = document.createElement('div');
        previewArea.style.cssText = `
          background: white;
          border: 1px solid #ddd;
          border-top: none;
          border-radius: 0 0 8px 8px;
          padding: 20px;
          min-height: 400px;
        `;
        viewer.appendChild(previewArea);

        // Get and render actual content
        const htmlResult = await (window as any).DocxodusWorkerTests.convertToHtml(Array.from(docBytes));

        if (!htmlResult.error) {
          const parser = new DOMParser();
          const doc = parser.parseFromString(htmlResult.html, 'text/html');

          // Copy styles
          doc.querySelectorAll('style').forEach(s => previewArea.appendChild(s.cloneNode(true)));

          // Create content container with proper styling
          const content = document.createElement('div');
          content.style.cssText = 'max-height: 500px; overflow-y: auto; padding-right: 10px;';
          content.innerHTML = doc.body.innerHTML;
          previewArea.appendChild(content);
        }

        // Add footer with status
        const footer = document.createElement('div');
        footer.style.cssText = `
          margin-top: 15px;
          padding: 10px;
          background: #e8f5e9;
          border-radius: 4px;
          font-size: 13px;
          color: #2e7d32;
        `;
        footer.textContent = '✓ Document loaded successfully via Web Worker';
        viewer.appendChild(footer);

      }, Array.from(bytes));

      // Take screenshot for visual verification
      const screenshot = await page.screenshot({
        fullPage: false,
        clip: { x: 0, y: 0, width: 1024, height: 768 }
      });

      // Verify screenshot was captured (non-empty)
      expect(screenshot.length).toBeGreaterThan(1000);

      // Verify DOM elements exist
      const viewerVisible = await page.locator('#lazy-viewer').isVisible();
      expect(viewerVisible).toBe(true);

      const headerText = await page.locator('#lazy-viewer h2').textContent();
      expect(headerText).toContain('Lazy Loading');

      console.log('Screenshot captured successfully');
      console.log(`Screenshot size: ${screenshot.length} bytes`);
    }, { timeout: 120000 });
  });
});
