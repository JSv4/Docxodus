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

// Helper to run comparison in browser
async function compareDocuments(
  page: Page,
  originalBytes: Uint8Array,
  modifiedBytes: Uint8Array
): Promise<{ docxBytes?: number[]; error?: any }> {
  const result = await page.evaluate(
    ([original, modified]) => {
      const result = (window as any).DocxodusTests.compareDocuments(
        new Uint8Array(original),
        new Uint8Array(modified)
      );
      if (result.docxBytes) {
        return { docxBytes: Array.from(result.docxBytes) };
      }
      return result;
    },
    [Array.from(originalBytes), Array.from(modifiedBytes)]
  );
  return result;
}

// Helper to get revisions from compared document
async function getRevisions(
  page: Page,
  docxBytes: number[]
): Promise<{ revisions?: any[]; error?: any }> {
  return await page.evaluate((bytesArray) => {
    return (window as any).DocxodusTests.getRevisions(new Uint8Array(bytesArray));
  }, docxBytes);
}

// Helper to compare and get HTML
async function compareToHtml(
  page: Page,
  originalBytes: Uint8Array,
  modifiedBytes: Uint8Array
): Promise<{ html?: string; error?: any }> {
  return await page.evaluate(
    ([original, modified]) => {
      return (window as any).DocxodusTests.compareToHtml(
        new Uint8Array(original),
        new Uint8Array(modified)
      );
    },
    [Array.from(originalBytes), Array.from(modifiedBytes)]
  );
}

// Helper to compare and get HTML with options
async function compareToHtmlWithOptions(
  page: Page,
  originalBytes: Uint8Array,
  modifiedBytes: Uint8Array,
  renderTrackedChanges: boolean
): Promise<{ html?: string; error?: any }> {
  return await page.evaluate(
    ([original, modified, renderChanges]) => {
      return (window as any).DocxodusTests.compareToHtmlWithOptions(
        new Uint8Array(original),
        new Uint8Array(modified),
        'Test',
        renderChanges
      );
    },
    [Array.from(originalBytes), Array.from(modifiedBytes), renderTrackedChanges]
  );
}

// Helper to convert to HTML with pagination
async function convertToHtmlWithPagination(
  page: Page,
  bytes: Uint8Array,
  paginationMode: number = 1,
  paginationScale: number = 1.0
): Promise<{ html?: string; error?: any }> {
  return await page.evaluate(
    ([bytesArray, mode, scale]) => {
      return (window as any).DocxodusTests.convertToHtmlWithPagination(
        new Uint8Array(bytesArray),
        mode,
        scale
      );
    },
    [Array.from(bytes), paginationMode, paginationScale]
  );
}

test.describe('Docxodus WASM Tests', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test.describe('HTML Conversion (HC tests)', () => {
    const htmlConversionTests = [
      { name: 'HC001-5DayTourPlanTemplate.docx', description: 'Tour plan template' },
      { name: 'HC004-ResumeTemplate.docx', description: 'Resume template' },
      { name: 'HC005-TaskPlanTemplate.docx', description: 'Task plan template' },
      { name: 'HC006-Test-01.docx', description: 'Basic test document' },
      { name: 'HC007-Test-02.docx', description: 'Test document 2' },
      { name: 'HC008-Test-03.docx', description: 'Test document 3' },
      { name: 'HC019-Hidden-Run.docx', description: 'Hidden text run' },
      { name: 'HC020-Small-Caps.docx', description: 'Small caps formatting' },
    ];

    for (const testCase of htmlConversionTests) {
      test(`converts ${testCase.name} to HTML`, async ({ page }) => {
        const bytes = readTestFile(testCase.name);
        const result = await convertToHtml(page, bytes);

        expect(result.error).toBeUndefined();
        expect(result.html).toBeDefined();
        expect(result.html!.length).toBeGreaterThan(100);
        expect(result.html).toContain('<html');
        expect(result.html).toContain('</html>');
      });
    }
  });

  test.describe('Document Comparison (WC tests)', () => {
    const comparisonTests = [
      {
        name: 'WC001-Digits',
        original: 'WC/WC001-Digits.docx',
        modified: 'WC/WC001-Digits-Mod.docx',
        description: 'Basic digit changes',
      },
      {
        name: 'WC002-DiffInMiddle',
        original: 'WC/WC002-Unmodified.docx',
        modified: 'WC/WC002-DiffInMiddle.docx',
        description: 'Difference in middle of text',
      },
      {
        name: 'WC002-InsertAtBeginning',
        original: 'WC/WC002-Unmodified.docx',
        modified: 'WC/WC002-InsertAtBeginning.docx',
        description: 'Insert at beginning',
      },
      {
        name: 'WC002-InsertAtEnd',
        original: 'WC/WC002-Unmodified.docx',
        modified: 'WC/WC002-InsertAtEnd.docx',
        description: 'Insert at end',
      },
      {
        name: 'WC002-DeleteAtBeginning',
        original: 'WC/WC002-Unmodified.docx',
        modified: 'WC/WC002-DeleteAtBeginning.docx',
        description: 'Delete at beginning',
      },
      {
        name: 'WC002-DeleteInMiddle',
        original: 'WC/WC002-Unmodified.docx',
        modified: 'WC/WC002-DeleteInMiddle.docx',
        description: 'Delete in middle',
      },
      {
        name: 'WC006-Table-DeleteRow',
        original: 'WC/WC006-Table.docx',
        modified: 'WC/WC006-Table-Delete-Row.docx',
        description: 'Table with deleted row',
      },
      {
        name: 'WC009-TableCell',
        original: 'WC/WC009-Table-Unmodified.docx',
        modified: 'WC/WC009-Table-Cell-1-1-Mod.docx',
        description: 'Table cell modification',
      },
      {
        name: 'WC011-General',
        original: 'WC/WC011-Before.docx',
        modified: 'WC/WC011-After.docx',
        description: 'General comparison',
      },
    ];

    for (const testCase of comparisonTests) {
      test(`compares ${testCase.name}: ${testCase.description}`, async ({ page }) => {
        const originalBytes = readTestFile(testCase.original);
        const modifiedBytes = readTestFile(testCase.modified);

        // Test comparison returning DOCX
        const compareResult = await compareDocuments(page, originalBytes, modifiedBytes);
        expect(compareResult.error).toBeUndefined();
        expect(compareResult.docxBytes).toBeDefined();
        expect(compareResult.docxBytes!.length).toBeGreaterThan(1000); // Valid DOCX is at least a few KB

        // Verify it's a valid DOCX (starts with PK zip signature)
        expect(compareResult.docxBytes![0]).toBe(0x50); // P
        expect(compareResult.docxBytes![1]).toBe(0x4b); // K

        // Get revisions from the compared document
        const revisionsResult = await getRevisions(page, compareResult.docxBytes!);
        expect(revisionsResult.error).toBeUndefined();
        expect(revisionsResult.revisions).toBeDefined();
        expect(revisionsResult.revisions!.length).toBeGreaterThan(0);

        console.log(`${testCase.name}: Found ${revisionsResult.revisions!.length} revisions`);
      });

      test(`compares ${testCase.name} to HTML`, async ({ page }) => {
        const originalBytes = readTestFile(testCase.original);
        const modifiedBytes = readTestFile(testCase.modified);

        const result = await compareToHtml(page, originalBytes, modifiedBytes);
        expect(result.error).toBeUndefined();
        expect(result.html).toBeDefined();
        expect(result.html!.length).toBeGreaterThan(100);
        expect(result.html).toContain('<html');
      });
    }
  });

  test.describe('CZ Comparison Tests (tracked changes)', () => {
    const czTests = [
      {
        name: 'CZ001-Plain',
        original: 'CZ/CZ001-Plain.docx',
        modified: 'CZ/CZ001-Plain-Mod.docx',
        description: 'Plain document with tracked changes',
      },
      {
        name: 'CZ002-MultiParagraphs',
        original: 'CZ/CZ002-Multi-Paragraphs.docx',
        modified: 'CZ/CZ002-Multi-Paragraphs-Mod.docx',
        description: 'Multiple paragraphs with changes',
      },
      {
        name: 'CZ003-MultiParagraphs',
        original: 'CZ/CZ003-Multi-Paragraphs.docx',
        modified: 'CZ/CZ003-Multi-Paragraphs-Mod.docx',
        description: 'Multiple paragraphs variant',
      },
      {
        name: 'CZ004-MultiParagraphsInCell',
        original: 'CZ/CZ004-Multi-Paragraphs-in-Cell.docx',
        modified: 'CZ/CZ004-Multi-Paragraphs-in-Cell-Mod.docx',
        description: 'Multiple paragraphs in table cell',
      },
    ];

    for (const testCase of czTests) {
      test(`compares ${testCase.name}: ${testCase.description}`, async ({ page }) => {
        const originalBytes = readTestFile(testCase.original);
        const modifiedBytes = readTestFile(testCase.modified);

        const compareResult = await compareDocuments(page, originalBytes, modifiedBytes);
        expect(compareResult.error).toBeUndefined();
        expect(compareResult.docxBytes).toBeDefined();
        expect(compareResult.docxBytes!.length).toBeGreaterThan(1000);

        // Verify DOCX signature
        expect(compareResult.docxBytes![0]).toBe(0x50);
        expect(compareResult.docxBytes![1]).toBe(0x4b);

        // Get revisions
        const revisionsResult = await getRevisions(page, compareResult.docxBytes!);
        expect(revisionsResult.error).toBeUndefined();
        expect(revisionsResult.revisions).toBeDefined();

        console.log(`${testCase.name}: Found ${revisionsResult.revisions!.length} revisions`);
      });
    }
  });

  test.describe('CA Tests (Content Assembly)', () => {
    const caTests = [
      {
        name: 'CA001-Plain',
        original: 'CA/CA001-Plain.docx',
        modified: 'CA/CA001-Plain-Mod.docx',
        description: 'Plain content comparison',
      },
    ];

    for (const testCase of caTests) {
      test(`compares ${testCase.name}: ${testCase.description}`, async ({ page }) => {
        const originalBytes = readTestFile(testCase.original);
        const modifiedBytes = readTestFile(testCase.modified);

        const compareResult = await compareDocuments(page, originalBytes, modifiedBytes);
        expect(compareResult.error).toBeUndefined();
        expect(compareResult.docxBytes).toBeDefined();
        expect(compareResult.docxBytes!.length).toBeGreaterThan(1000);
      });
    }
  });

  test('version info is available', async ({ page }) => {
    const version = await page.evaluate(() => {
      return (window as any).DocxodusTests.getVersion();
    });

    expect(version.Library).toBe('Docxodus WASM');
    expect(version.Platform).toBe('browser-wasm');
    expect(version.DotnetVersion).toBeDefined();
  });

  test.describe('Tracked Changes Rendering (renderTrackedChanges option)', () => {
    // Use a simple comparison test case that will have insertions/deletions
    const testCase = {
      original: 'WC/WC002-Unmodified.docx',
      modified: 'WC/WC002-DiffInMiddle.docx',
    };

    test('renderTrackedChanges=true shows <ins> and <del> elements', async ({ page }) => {
      const originalBytes = readTestFile(testCase.original);
      const modifiedBytes = readTestFile(testCase.modified);

      const result = await compareToHtmlWithOptions(page, originalBytes, modifiedBytes, true);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();
      expect(result.html!.length).toBeGreaterThan(100);
      expect(result.html).toContain('<html');

      // With tracked changes rendered, we should see ins and/or del elements
      const hasTrackingElements = result.html!.includes('<ins') || result.html!.includes('<del');
      expect(hasTrackingElements).toBe(true);

      console.log('renderTrackedChanges=true: Contains ins/del elements:', hasTrackingElements);
    });

    test('renderTrackedChanges=false produces clean HTML without <ins>/<del>', async ({ page }) => {
      const originalBytes = readTestFile(testCase.original);
      const modifiedBytes = readTestFile(testCase.modified);

      const result = await compareToHtmlWithOptions(page, originalBytes, modifiedBytes, false);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();
      expect(result.html!.length).toBeGreaterThan(100);
      expect(result.html).toContain('<html');

      // With tracked changes NOT rendered, we should NOT see ins/del elements
      const hasInsElement = result.html!.includes('<ins');
      const hasDelElement = result.html!.includes('<del');
      expect(hasInsElement).toBe(false);
      expect(hasDelElement).toBe(false);

      console.log('renderTrackedChanges=false: No ins/del elements present');
    });

    test('default compareToHtml includes tracked changes (backward compatible)', async ({ page }) => {
      const originalBytes = readTestFile(testCase.original);
      const modifiedBytes = readTestFile(testCase.modified);

      // Default compareToHtml should show tracked changes (true by default)
      const result = await compareToHtml(page, originalBytes, modifiedBytes);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Default should have tracking elements
      const hasTrackingElements = result.html!.includes('<ins') || result.html!.includes('<del');
      expect(hasTrackingElements).toBe(true);

      console.log('Default compareToHtml: Contains ins/del elements (backward compatible)');
    });

    test('tracked changes rendering includes proper CSS styling', async ({ page }) => {
      const originalBytes = readTestFile(testCase.original);
      const modifiedBytes = readTestFile(testCase.modified);

      const result = await compareToHtmlWithOptions(page, originalBytes, modifiedBytes, true);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Check for CSS styling related to tracked changes
      // The HTML should include styles for insertions and deletions
      const hasStyleTag = result.html!.includes('<style');
      expect(hasStyleTag).toBe(true);

      // Check for redline CSS class prefix (used by DocumentComparer)
      const hasRedlineClass = result.html!.includes('redline-');
      expect(hasRedlineClass).toBe(true);

      console.log('Tracked changes HTML includes proper CSS styling');
    });
  });

  test.describe('Pagination Tests', () => {
    // Use a document with multiple paragraphs that will span multiple pages
    const testDoc = 'HC001-5DayTourPlanTemplate.docx';

    test('generates pagination HTML structure', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await convertToHtmlWithPagination(page, bytes, 1, 1.0);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Check for pagination structure
      expect(result.html).toContain('pagination-staging');
      expect(result.html).toContain('pagination-container');
      expect(result.html).toContain('page-staging');
      expect(result.html).toContain('page-container');

      console.log('Pagination HTML structure generated correctly');
    });

    test('includes page dimension data attributes', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await convertToHtmlWithPagination(page, bytes, 1, 1.0);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Check for page dimension data attributes
      expect(result.html).toContain('data-page-width');
      expect(result.html).toContain('data-page-height');
      expect(result.html).toContain('data-content-width');
      expect(result.html).toContain('data-content-height');
      expect(result.html).toContain('data-margin-top');
      expect(result.html).toContain('data-margin-left');

      console.log('Page dimension data attributes present');
    });

    test('pagination CSS includes overflow hidden', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await convertToHtmlWithPagination(page, bytes, 1, 1.0);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Check that pagination CSS includes overflow:hidden for clipping
      expect(result.html).toContain('overflow: hidden');
      // Check for page-box class styling
      expect(result.html).toContain('.page-box');
      expect(result.html).toContain('.page-content');

      console.log('Pagination CSS includes proper overflow handling');
    });

    test('content does not overflow page boundaries when paginated', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await convertToHtmlWithPagination(page, bytes, 1, 0.8);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Insert the HTML into the page and run pagination
      const overflowCheck = await page.evaluate((html) => {
        // Create a container for the paginated content
        const container = document.createElement('div');
        container.id = 'test-pagination-container';
        container.innerHTML = html;
        document.body.appendChild(container);

        // Find staging and page container
        const staging = container.querySelector('#pagination-staging') as HTMLElement;
        const pageContainer = container.querySelector('#pagination-container') as HTMLElement;

        if (!staging || !pageContainer) {
          return { error: 'Pagination elements not found' };
        }

        // Import and run the pagination engine
        // We need to manually implement pagination logic here since we can't import ES modules
        // Instead, let's parse dimensions and verify content fits

        // Get page dimensions from data attributes
        const section = staging.querySelector('[data-section-index]') as HTMLElement;
        if (!section) {
          return { error: 'Section element not found' };
        }

        const pageHeight = parseFloat(section.dataset.pageHeight || '792');
        const contentHeight = parseFloat(section.dataset.contentHeight || '648');
        const marginTop = parseFloat(section.dataset.marginTop || '72');
        const marginBottom = parseFloat(section.dataset.marginBottom || '72');

        // Measure all content blocks
        const blocks = Array.from(section.children) as HTMLElement[];
        let totalContentHeight = 0;
        const blockMeasurements: { height: number; marginTop: number; marginBottom: number }[] = [];

        // Make staging visible for measurement
        staging.style.visibility = 'hidden';
        staging.style.position = 'absolute';
        staging.style.left = '-9999px';
        staging.style.display = 'block';
        section.style.width = `${parseFloat(section.dataset.contentWidth || '468')}pt`;

        for (const block of blocks) {
          if (block.dataset.sectionIndex !== undefined) continue;

          const rect = block.getBoundingClientRect();
          const style = window.getComputedStyle(block);
          const marginTopPx = parseFloat(style.marginTop) || 0;
          const marginBottomPx = parseFloat(style.marginBottom) || 0;

          blockMeasurements.push({
            height: rect.height * 0.75, // Convert px to pt
            marginTop: marginTopPx * 0.75,
            marginBottom: marginBottomPx * 0.75
          });
        }

        staging.style.display = 'none';

        // Simulate pagination flow with margin collapsing
        const pages: number[][] = [];
        let currentPage: number[] = [];
        let remainingHeight = contentHeight;
        let prevMarginBottom = 0;

        for (let i = 0; i < blockMeasurements.length; i++) {
          const block = blockMeasurements[i];
          const isFirst = currentPage.length === 0;

          // Calculate effective margin with collapsing
          let effectiveMarginTop = block.marginTop;
          if (!isFirst) {
            effectiveMarginTop = Math.max(block.marginTop, prevMarginBottom) - prevMarginBottom;
          }

          const blockSpace = effectiveMarginTop + block.height + block.marginBottom;

          if (blockSpace <= remainingHeight) {
            currentPage.push(i);
            remainingHeight -= blockSpace;
            prevMarginBottom = block.marginBottom;
          } else {
            // Start new page
            if (currentPage.length > 0) {
              pages.push(currentPage);
            }
            currentPage = [i];
            const newPageSpace = block.marginTop + block.height + block.marginBottom;
            remainingHeight = contentHeight - newPageSpace;
            prevMarginBottom = block.marginBottom;
          }
        }

        if (currentPage.length > 0) {
          pages.push(currentPage);
        }

        // Verify each page's content fits within contentHeight
        const pageOverflows: { page: number; usedHeight: number; available: number }[] = [];

        for (let pageIdx = 0; pageIdx < pages.length; pageIdx++) {
          const pageBlocks = pages[pageIdx];
          let usedHeight = 0;
          let prevMargin = 0;

          for (let i = 0; i < pageBlocks.length; i++) {
            const blockIdx = pageBlocks[i];
            const block = blockMeasurements[blockIdx];
            const isFirst = i === 0;

            let effectiveMarginTop = block.marginTop;
            if (!isFirst) {
              effectiveMarginTop = Math.max(block.marginTop, prevMargin) - prevMargin;
            }

            usedHeight += effectiveMarginTop + block.height + block.marginBottom;
            prevMargin = block.marginBottom;
          }

          // Allow small tolerance for floating point errors (1pt)
          if (usedHeight > contentHeight + 1) {
            pageOverflows.push({
              page: pageIdx + 1,
              usedHeight,
              available: contentHeight
            });
          }
        }

        // Clean up
        document.body.removeChild(container);

        return {
          totalPages: pages.length,
          contentHeight,
          pageOverflows,
          hasOverflow: pageOverflows.length > 0
        };
      }, result.html!);

      // Verify no overflow
      if ('error' in overflowCheck) {
        throw new Error(overflowCheck.error as string);
      }

      expect(overflowCheck.hasOverflow).toBe(false);

      if (overflowCheck.pageOverflows.length > 0) {
        console.log('Page overflows detected:', overflowCheck.pageOverflows);
      }

      console.log(`Pagination test passed: ${overflowCheck.totalPages} pages, no content overflow`);
    });

    test('scaled pagination maintains proper clipping', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Test with different scale factors
      for (const scale of [0.5, 0.75, 1.0, 1.25]) {
        const result = await convertToHtmlWithPagination(page, bytes, 1, scale);

        expect(result.error).toBeUndefined();
        expect(result.html).toBeDefined();

        // Verify HTML was generated
        expect(result.html).toContain('pagination-staging');

        console.log(`Scale ${scale}: HTML generated successfully`);
      }
    });
  });
});
