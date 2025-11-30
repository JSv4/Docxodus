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

// Helper to convert to HTML with annotations
async function convertToHtmlWithAnnotations(
  page: Page,
  bytes: Uint8Array,
  renderAnnotations: boolean = true,
  annotationLabelMode: number = 0
): Promise<{ html?: string; error?: any }> {
  return await page.evaluate(
    ([bytesArray, render, labelMode]) => {
      return (window as any).DocxodusTests.convertToHtmlWithAnnotations(
        new Uint8Array(bytesArray),
        render,
        labelMode
      );
    },
    [Array.from(bytes), renderAnnotations, annotationLabelMode]
  );
}

// Helper to get annotations from a document
async function getAnnotationsFromDoc(
  page: Page,
  bytes: Uint8Array
): Promise<{ annotations?: any[]; error?: any }> {
  return await page.evaluate((bytesArray) => {
    return (window as any).DocxodusTests.getAnnotations(new Uint8Array(bytesArray));
  }, Array.from(bytes));
}

// Helper to add an annotation to a document
async function addAnnotationToDoc(
  page: Page,
  bytes: Uint8Array,
  request: any
): Promise<{ success?: boolean; documentBytes?: number[]; annotation?: any; error?: any }> {
  const result = await page.evaluate(
    ([bytesArray, req]) => {
      const result = (window as any).DocxodusTests.addAnnotation(
        new Uint8Array(bytesArray),
        req
      );
      if (result.documentBytes) {
        return {
          success: result.success,
          documentBytes: Array.from(result.documentBytes),
          annotation: result.annotation
        };
      }
      return result;
    },
    [Array.from(bytes), request]
  );
  return result;
}

// Helper to remove an annotation from a document
async function removeAnnotationFromDoc(
  page: Page,
  bytes: number[],
  annotationId: string
): Promise<{ success?: boolean; documentBytes?: number[]; error?: any }> {
  const result = await page.evaluate(
    ([bytesArray, id]) => {
      const result = (window as any).DocxodusTests.removeAnnotation(
        new Uint8Array(bytesArray),
        id
      );
      if (result.documentBytes) {
        return {
          success: result.success,
          documentBytes: Array.from(result.documentBytes)
        };
      }
      return result;
    },
    [bytes, annotationId]
  );
  return result;
}

// Helper to check if a document has annotations
async function hasAnnotationsInDoc(
  page: Page,
  bytes: Uint8Array | number[]
): Promise<{ hasAnnotations?: boolean; error?: any }> {
  return await page.evaluate((bytesArray) => {
    return (window as any).DocxodusTests.hasAnnotations(new Uint8Array(bytesArray));
  }, Array.from(bytes as any));
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

      // Load the real pagination engine bundle
      await page.addScriptTag({ path: 'dist/pagination.bundle.js' });

      // Insert the HTML into the page and run pagination using the real PaginationEngine
      const paginationResult = await page.evaluate((html) => {
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

        // Use the real PaginationEngine from the bundle
        const { PaginationEngine } = (window as any).DocxodusPagination;

        try {
          const engine = new PaginationEngine(staging, pageContainer, {
            scale: 0.8,
            showPageNumbers: true
          });

          const result = engine.paginate();

          // Now verify that content doesn't overflow in the rendered pages
          const pageBoxes = pageContainer.querySelectorAll('.page-box');
          const overflows: { page: number; contentBottom: number; pageBottom: number }[] = [];

          pageBoxes.forEach((pageBox, idx) => {
            const pageContent = pageBox.querySelector('.page-content') as HTMLElement;
            if (!pageContent) return;

            const pageBoxRect = pageBox.getBoundingClientRect();
            const contentRect = pageContent.getBoundingClientRect();

            // Check if content bottom exceeds page box bottom (accounting for transform/zoom)
            // We need to check the actual children inside page-content
            const children = pageContent.children;
            if (children.length > 0) {
              const lastChild = children[children.length - 1] as HTMLElement;
              const lastChildRect = lastChild.getBoundingClientRect();
              const style = window.getComputedStyle(lastChild);
              const marginBottom = parseFloat(style.marginBottom) || 0;

              // Content should not extend beyond the page-content container
              const contentBottom = lastChildRect.bottom + marginBottom;
              const containerBottom = contentRect.bottom;

              // Allow 1px tolerance for rounding
              if (contentBottom > containerBottom + 1) {
                overflows.push({
                  page: idx + 1,
                  contentBottom: contentBottom,
                  pageBottom: containerBottom
                });
              }
            }
          });

          // Clean up
          document.body.removeChild(container);

          return {
            totalPages: result.totalPages,
            pageOverflows: overflows,
            hasOverflow: overflows.length > 0
          };
        } catch (e) {
          document.body.removeChild(container);
          return { error: (e as Error).message };
        }
      }, result.html!);

      // Verify no errors
      if ('error' in paginationResult) {
        throw new Error(paginationResult.error as string);
      }

      expect(paginationResult.hasOverflow).toBe(false);

      if (paginationResult.pageOverflows && paginationResult.pageOverflows.length > 0) {
        console.log('Page overflows detected:', paginationResult.pageOverflows);
      }

      console.log(`Pagination test passed: ${paginationResult.totalPages} pages, no content overflow`);
    });

    test('scaled pagination maintains proper clipping', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Load the real pagination engine bundle
      await page.addScriptTag({ path: 'dist/pagination.bundle.js' });

      // Test with different scale factors
      for (const scale of [0.5, 0.75, 1.0, 1.25]) {
        const result = await convertToHtmlWithPagination(page, bytes, 1, scale);

        expect(result.error).toBeUndefined();
        expect(result.html).toBeDefined();

        // Run pagination with the real engine and verify no overflow
        const paginationResult = await page.evaluate(({ html, scale }) => {
          // Create a container for the paginated content
          const container = document.createElement('div');
          container.id = `test-pagination-container-${scale}`;
          container.innerHTML = html;
          document.body.appendChild(container);

          const staging = container.querySelector('#pagination-staging') as HTMLElement;
          const pageContainer = container.querySelector('#pagination-container') as HTMLElement;

          if (!staging || !pageContainer) {
            document.body.removeChild(container);
            return { error: 'Pagination elements not found' };
          }

          const { PaginationEngine } = (window as any).DocxodusPagination;

          try {
            const engine = new PaginationEngine(staging, pageContainer, {
              scale: scale,
              showPageNumbers: true
            });

            const result = engine.paginate();

            // Verify page boxes were created
            const pageBoxes = pageContainer.querySelectorAll('.page-box');

            // Check overflow on each page
            let hasOverflow = false;
            pageBoxes.forEach((pageBox) => {
              const pageContent = pageBox.querySelector('.page-content') as HTMLElement;
              if (!pageContent) return;

              const contentRect = pageContent.getBoundingClientRect();
              const children = pageContent.children;

              if (children.length > 0) {
                const lastChild = children[children.length - 1] as HTMLElement;
                const lastChildRect = lastChild.getBoundingClientRect();
                const style = window.getComputedStyle(lastChild);
                const marginBottom = parseFloat(style.marginBottom) || 0;

                if (lastChildRect.bottom + marginBottom > contentRect.bottom + 1) {
                  hasOverflow = true;
                }
              }
            });

            document.body.removeChild(container);

            return {
              totalPages: result.totalPages,
              pageBoxCount: pageBoxes.length,
              hasOverflow: hasOverflow
            };
          } catch (e) {
            document.body.removeChild(container);
            return { error: (e as Error).message };
          }
        }, { html: result.html!, scale });

        if ('error' in paginationResult) {
          throw new Error(`Scale ${scale}: ${paginationResult.error}`);
        }

        expect(paginationResult.totalPages).toBeGreaterThan(0);
        expect(paginationResult.hasOverflow).toBe(false);

        console.log(`Scale ${scale}: ${paginationResult.totalPages} pages, no overflow`);
      }
    });
  });

  test.describe('Annotation Tests', () => {
    // Use a simple document for annotation tests
    const testDoc = 'HC006-Test-01.docx';

    test('document initially has no annotations', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await hasAnnotationsInDoc(page, bytes);

      expect(result.error).toBeUndefined();
      expect(result.hasAnnotations).toBe(false);

      console.log('Document has no initial annotations');
    });

    test('can add an annotation using text search', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'test-annot-1',
        LabelId: 'CLAUSE_A',
        Label: 'Test Clause',
        Color: '#FFEB3B',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();
      expect(addResult.success).toBe(true);
      expect(addResult.documentBytes).toBeDefined();
      expect(addResult.documentBytes!.length).toBeGreaterThan(1000);
      expect(addResult.annotation).toBeDefined();
      expect(addResult.annotation.Id).toBe('test-annot-1');

      console.log('Added annotation:', addResult.annotation);
    });

    test('can retrieve annotations from a document', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'test-annot-retrieve',
        LabelId: 'SECTION_1',
        Label: 'Section One',
        Color: '#4CAF50',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();
      expect(addResult.documentBytes).toBeDefined();

      // Retrieve annotations
      const getResult = await getAnnotationsFromDoc(
        page,
        new Uint8Array(addResult.documentBytes!)
      );

      expect(getResult.error).toBeUndefined();
      expect(getResult.annotations).toBeDefined();
      expect(getResult.annotations!.length).toBe(1);
      expect(getResult.annotations![0].Id).toBe('test-annot-retrieve');
      expect(getResult.annotations![0].LabelId).toBe('SECTION_1');
      expect(getResult.annotations![0].Label).toBe('Section One');
      expect(getResult.annotations![0].Color).toBe('#4CAF50');

      console.log('Retrieved annotations:', getResult.annotations);
    });

    test('can add multiple annotations', async ({ page }) => {
      let bytes = readTestFile(testDoc);

      // Add first annotation
      const addResult1 = await addAnnotationToDoc(page, bytes, {
        Id: 'multi-annot-1',
        LabelId: 'CLAUSE_A',
        Label: 'First',
        Color: '#FFEB3B',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult1.error).toBeUndefined();
      expect(addResult1.success).toBe(true);

      // Add second annotation to modified document
      const addResult2 = await addAnnotationToDoc(
        page,
        new Uint8Array(addResult1.documentBytes!),
        {
          Id: 'multi-annot-2',
          LabelId: 'CLAUSE_B',
          Label: 'Second',
          Color: '#4CAF50',
          SearchText: 'and',
          Occurrence: 1
        }
      );

      expect(addResult2.error).toBeUndefined();
      expect(addResult2.success).toBe(true);

      // Verify both annotations exist
      const getResult = await getAnnotationsFromDoc(
        page,
        new Uint8Array(addResult2.documentBytes!)
      );

      expect(getResult.error).toBeUndefined();
      expect(getResult.annotations).toBeDefined();
      expect(getResult.annotations!.length).toBe(2);

      const ids = getResult.annotations!.map((a: any) => a.Id);
      expect(ids).toContain('multi-annot-1');
      expect(ids).toContain('multi-annot-2');

      console.log('Added multiple annotations:', getResult.annotations!.length);
    });

    test('can remove an annotation', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'test-annot-remove',
        LabelId: 'REMOVE_ME',
        Label: 'To Remove',
        Color: '#F44336',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();
      expect(addResult.documentBytes).toBeDefined();

      // Verify it was added
      const hasResult1 = await hasAnnotationsInDoc(page, addResult.documentBytes!);
      expect(hasResult1.hasAnnotations).toBe(true);

      // Remove the annotation
      const removeResult = await removeAnnotationFromDoc(
        page,
        addResult.documentBytes!,
        'test-annot-remove'
      );

      expect(removeResult.error).toBeUndefined();
      expect(removeResult.success).toBe(true);
      expect(removeResult.documentBytes).toBeDefined();

      // Verify it was removed
      const hasResult2 = await hasAnnotationsInDoc(page, removeResult.documentBytes!);
      expect(hasResult2.hasAnnotations).toBe(false);

      console.log('Successfully removed annotation');
    });

    test('annotation rendering generates highlight spans', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'render-test',
        LabelId: 'HIGHLIGHT',
        Label: 'Highlighted Text',
        Color: '#FFEB3B',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();

      // Convert to HTML with annotations enabled
      const htmlResult = await convertToHtmlWithAnnotations(
        page,
        new Uint8Array(addResult.documentBytes!),
        true,
        0  // AnnotationLabelMode.Above
      );

      expect(htmlResult.error).toBeUndefined();
      expect(htmlResult.html).toBeDefined();

      // Check for annotation highlight elements
      expect(htmlResult.html).toContain('annot-highlight');
      expect(htmlResult.html).toContain('data-annotation-id="render-test"');
      expect(htmlResult.html).toContain('Highlighted Text');  // The label text

      console.log('Annotation rendering generates highlight spans with labels');
    });

    test('annotation CSS includes highlight styles', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation with a specific color
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'css-test',
        LabelId: 'STYLE_CHECK',
        Label: 'Styled',
        Color: '#4CAF50',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();

      // Convert to HTML with annotations
      const htmlResult = await convertToHtmlWithAnnotations(
        page,
        new Uint8Array(addResult.documentBytes!),
        true,
        0
      );

      expect(htmlResult.error).toBeUndefined();
      expect(htmlResult.html).toBeDefined();

      // Check for annotation CSS
      expect(htmlResult.html).toContain('<style');
      expect(htmlResult.html).toContain('.annot-highlight');

      // Check that the annotation color is used
      expect(htmlResult.html).toContain('#4CAF50');

      console.log('Annotation CSS includes highlight styles with proper color');
    });

    test('annotation label modes render differently', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'label-mode-test',
        LabelId: 'MODE_TEST',
        Label: 'Test Label',
        Color: '#2196F3',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();
      const annotatedBytes = new Uint8Array(addResult.documentBytes!);

      // Test different label modes
      const modes = [
        { mode: 0, name: 'Above', checkFor: 'annot-label' },
        { mode: 1, name: 'Inline', checkFor: 'annot-label' },
        { mode: 2, name: 'Tooltip', checkFor: 'annot-highlight' },
        { mode: 3, name: 'None', checkFor: 'annot-highlight' }
      ];

      for (const { mode, name, checkFor } of modes) {
        const htmlResult = await convertToHtmlWithAnnotations(
          page,
          annotatedBytes,
          true,
          mode
        );

        expect(htmlResult.error).toBeUndefined();
        expect(htmlResult.html).toBeDefined();
        expect(htmlResult.html).toContain(checkFor);

        console.log(`Label mode ${name} (${mode}): Rendered correctly`);
      }
    });

    test('annotation metadata is preserved', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation with metadata
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'metadata-test',
        LabelId: 'META',
        Label: 'With Metadata',
        Color: '#9C27B0',
        SearchText: 'the',
        Occurrence: 1,
        Author: 'Test Author',
        Metadata: {
          customKey: 'customValue',
          priority: 'high'
        }
      });

      expect(addResult.error).toBeUndefined();

      // Retrieve and verify metadata
      const getResult = await getAnnotationsFromDoc(
        page,
        new Uint8Array(addResult.documentBytes!)
      );

      expect(getResult.error).toBeUndefined();
      expect(getResult.annotations).toBeDefined();
      expect(getResult.annotations!.length).toBe(1);

      const annot = getResult.annotations![0];
      expect(annot.Author).toBe('Test Author');
      expect(annot.Metadata).toBeDefined();
      expect(annot.Metadata.customKey).toBe('customValue');
      expect(annot.Metadata.priority).toBe('high');

      console.log('Annotation metadata preserved:', annot.Metadata);
    });

    test('disabling annotation rendering produces clean HTML', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add an annotation
      const addResult = await addAnnotationToDoc(page, bytes, {
        Id: 'disable-test',
        LabelId: 'HIDDEN',
        Label: 'Should Be Hidden',
        Color: '#FF5722',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();

      // Convert to HTML with annotations DISABLED
      const htmlResult = await convertToHtmlWithAnnotations(
        page,
        new Uint8Array(addResult.documentBytes!),
        false,  // renderAnnotations = false
        0
      );

      expect(htmlResult.error).toBeUndefined();
      expect(htmlResult.html).toBeDefined();

      // Verify no annotation elements
      expect(htmlResult.html).not.toContain('annot-highlight');
      expect(htmlResult.html).not.toContain('data-annotation-id');

      console.log('Disabled annotation rendering produces clean HTML');
    });
  });
});
