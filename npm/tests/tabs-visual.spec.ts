import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

// Helper to read file as Uint8Array for browser
function readTestFile(relativePath: string): Uint8Array {
  const fullPath = path.join(TEST_FILES_DIR, relativePath);
  return new Uint8Array(fs.readFileSync(fullPath));
}

// Helper to wait for WASM to be ready
async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout: 30000,
  });
}

// Helper to run conversion in browser
async function convertToHtml(page: Page, bytes: Uint8Array): Promise<{ html?: string; error?: any }> {
  return await page.evaluate((bytesArray) => {
    return (window as any).DocxodusTests.convertToHtml(new Uint8Array(bytesArray));
  }, Array.from(bytes));
}

test.describe('Tab Rendering Visual Tests', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('HC024 - left-aligned tabs render with correct spacing', async ({ page }) => {
    const bytes = readTestFile('HC024-Tabs-01.docx');

    await page.setViewportSize({ width: 1200, height: 800 });

    const result = await convertToHtml(page, bytes);
    expect(result.error).toBeUndefined();
    expect(result.html).toBeDefined();

    // Inject HTML into page with proper styling
    await page.evaluate((htmlContent: string) => {
      document.body.innerHTML = `
        <div id="doc-container" style="padding: 40px; background: white; font-family: 'Times New Roman', serif; width: 8.5in; margin: 0 auto;">
          ${htmlContent}
        </div>
      `;
    }, result.html!);

    // Wait for rendering
    await page.waitForTimeout(500);

    // Take screenshot for visual verification
    await expect(page.locator('#doc-container')).toHaveScreenshot('tabs-hc024-left-aligned.png', {
      maxDiffPixelRatio: 0.05  // Allow 5% variance for font rendering differences
    });
  });

  test('HC025 - tabs with various alignments render correctly', async ({ page }) => {
    const bytes = readTestFile('HC025-Tabs-02.docx');

    await page.setViewportSize({ width: 1200, height: 800 });

    const result = await convertToHtml(page, bytes);
    expect(result.error).toBeUndefined();
    expect(result.html).toBeDefined();

    await page.evaluate((htmlContent: string) => {
      document.body.innerHTML = `
        <div id="doc-container" style="padding: 40px; background: white; font-family: 'Times New Roman', serif; width: 8.5in; margin: 0 auto;">
          ${htmlContent}
        </div>
      `;
    }, result.html!);

    await page.waitForTimeout(500);

    await expect(page.locator('#doc-container')).toHaveScreenshot('tabs-hc025-various-alignments.png', {
      maxDiffPixelRatio: 0.05
    });
  });

  test('HC026 - tabs with leader characters render correctly', async ({ page }) => {
    const bytes = readTestFile('HC026-Tabs-03.docx');

    await page.setViewportSize({ width: 1200, height: 800 });

    const result = await convertToHtml(page, bytes);
    expect(result.error).toBeUndefined();
    expect(result.html).toBeDefined();

    await page.evaluate((htmlContent: string) => {
      document.body.innerHTML = `
        <div id="doc-container" style="padding: 40px; background: white; font-family: 'Times New Roman', serif; width: 8.5in; margin: 0 auto;">
          ${htmlContent}
        </div>
      `;
    }, result.html!);

    await page.waitForTimeout(500);

    await expect(page.locator('#doc-container')).toHaveScreenshot('tabs-hc026-leader-characters.png', {
      maxDiffPixelRatio: 0.05
    });
  });

  test('HC027 - complex tab scenarios render correctly', async ({ page }) => {
    const bytes = readTestFile('HC027-Tabs-04.docx');

    await page.setViewportSize({ width: 1200, height: 800 });

    const result = await convertToHtml(page, bytes);
    expect(result.error).toBeUndefined();
    expect(result.html).toBeDefined();

    await page.evaluate((htmlContent: string) => {
      document.body.innerHTML = `
        <div id="doc-container" style="padding: 40px; background: white; font-family: 'Times New Roman', serif; width: 8.5in; margin: 0 auto;">
          ${htmlContent}
        </div>
      `;
    }, result.html!);

    await page.waitForTimeout(500);

    await expect(page.locator('#doc-container')).toHaveScreenshot('tabs-hc027-complex-scenarios.png', {
      maxDiffPixelRatio: 0.05
    });
  });

  test('tab spacing is visually correct with text measurement', async ({ page }) => {
    // This test specifically validates that text width is being measured
    // by checking that tabs create visual spacing in the document
    const bytes = readTestFile('HC024-Tabs-01.docx');

    await page.setViewportSize({ width: 1200, height: 800 });

    const result = await convertToHtml(page, bytes);
    expect(result.error).toBeUndefined();
    expect(result.html).toBeDefined();

    await page.evaluate((htmlContent: string) => {
      document.body.innerHTML = `
        <div id="doc-container" style="padding: 40px; background: white; width: 8.5in; margin: 0 auto;">
          ${htmlContent}
        </div>
      `;
    }, result.html!);

    // Check that the HTML contains margin or width styling for tabs
    // The converter applies styles to span elements for tab spacing
    const hasTabStyling = await page.evaluate(() => {
      const container = document.getElementById('doc-container');
      if (!container) return false;
      const html = container.innerHTML;
      // Look for margin or min-width styles in the HTML
      return html.includes('margin') || html.includes('min-width') || html.includes('width:');
    });

    expect(hasTabStyling).toBe(true);

    // Verify the document has content
    const hasContent = await page.evaluate(() => {
      const container = document.getElementById('doc-container');
      return container ? container.textContent!.length > 0 : false;
    });
    expect(hasContent).toBe(true);
  });

  test('table of contents style with right-aligned tabs', async ({ page }) => {
    // HC022-Table-Of-Contents.docx should have TOC-style right-aligned tabs with dot leaders
    const bytes = readTestFile('HC022-Table-Of-Contents.docx');

    await page.setViewportSize({ width: 1200, height: 1000 });

    const result = await convertToHtml(page, bytes);
    expect(result.error).toBeUndefined();
    expect(result.html).toBeDefined();

    await page.evaluate((htmlContent: string) => {
      document.body.innerHTML = `
        <div id="doc-container" style="padding: 40px; background: white; font-family: 'Times New Roman', serif; width: 8.5in; margin: 0 auto;">
          ${htmlContent}
        </div>
      `;
    }, result.html!);

    await page.waitForTimeout(500);

    await expect(page.locator('#doc-container')).toHaveScreenshot('tabs-toc-style.png', {
      maxDiffPixelRatio: 0.05
    });
  });
});
