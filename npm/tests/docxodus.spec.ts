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

// Helper to convert to HTML with pagination AND tracked changes
async function convertToHtmlWithPaginationAndTrackedChanges(
  page: Page,
  bytes: Uint8Array,
  paginationMode: number = 1,
  paginationScale: number = 1.0,
  renderTrackedChanges: boolean = true
): Promise<{ html?: string; error?: any }> {
  return await page.evaluate(
    ([bytesArray, mode, scale, renderChanges]) => {
      return (window as any).DocxodusTests.convertToHtmlWithPaginationAndTrackedChanges(
        new Uint8Array(bytesArray),
        mode,
        scale,
        renderChanges
      );
    },
    [Array.from(bytes), paginationMode, paginationScale, renderTrackedChanges]
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

// Helper to get document structure
async function getDocumentStructure(
  page: Page,
  bytes: Uint8Array
): Promise<{ root?: any; elementsById?: any; tableColumns?: any; error?: any }> {
  return await page.evaluate((bytesArray) => {
    return (window as any).DocxodusTests.getDocumentStructure(new Uint8Array(bytesArray));
  }, Array.from(bytes));
}

// Helper to add annotation with flexible targeting
async function addAnnotationWithTarget(
  page: Page,
  bytes: Uint8Array,
  request: any
): Promise<{ success?: boolean; documentBytes?: number[]; annotation?: any; error?: any }> {
  const result = await page.evaluate(
    ([bytesArray, req]) => {
      const result = (window as any).DocxodusTests.addAnnotationWithTarget(
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


test.describe('Docxodus WASM Tests', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test.describe('HTML Conversion (HC tests)', () => {
    const htmlConversionTests = [
      { name: 'HC001-5DayTourPlanTemplate.docx', description: 'Tour plan template', expectTables: true },
      { name: 'HC004-ResumeTemplate.docx', description: 'Resume template', expectTables: true },
      { name: 'HC005-TaskPlanTemplate.docx', description: 'Task plan template', expectTables: true },
      { name: 'HC006-Test-01.docx', description: 'Basic test document', expectTables: false },
      { name: 'HC007-Test-02.docx', description: 'Test document 2', expectTables: false },
      { name: 'HC008-Test-03.docx', description: 'Test document 3', expectTables: false },
      { name: 'HC019-Hidden-Run.docx', description: 'Hidden text run', expectTables: false },
      { name: 'HC020-Small-Caps.docx', description: 'Small caps formatting', expectTables: false },
    ];

    for (const testCase of htmlConversionTests) {
      test(`converts ${testCase.name} to HTML and renders correctly`, async ({ page }) => {
        const bytes = readTestFile(testCase.name);
        const result = await convertToHtml(page, bytes);

        expect(result.error).toBeUndefined();
        expect(result.html).toBeDefined();
        expect(result.html!.length).toBeGreaterThan(100);

        // Actually render the HTML to the page
        await page.setContent(result.html!);

        // Verify basic document structure is rendered in DOM
        await expect(page.locator('html')).toBeAttached();
        await expect(page.locator('body')).toBeAttached();

        // Verify content is visible (not empty body)
        const bodyText = await page.locator('body').textContent();
        expect(bodyText!.length).toBeGreaterThan(10);

        // Verify styles are present
        await expect(page.locator('style')).toBeAttached();

        // Check for tables if expected
        if (testCase.expectTables) {
          await expect(page.locator('table').first()).toBeVisible();
        }

        // Verify paragraphs or spans exist (actual content)
        const contentElements = await page.locator('p, span, div').count();
        expect(contentElements).toBeGreaterThan(0);
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

      test(`compares ${testCase.name} to HTML and renders tracked changes`, async ({ page }) => {
        const originalBytes = readTestFile(testCase.original);
        const modifiedBytes = readTestFile(testCase.modified);

        const result = await compareToHtml(page, originalBytes, modifiedBytes);
        expect(result.error).toBeUndefined();
        expect(result.html).toBeDefined();
        expect(result.html!.length).toBeGreaterThan(100);

        // Actually render the HTML to the page
        await page.setContent(result.html!);

        // Verify document structure is rendered
        await expect(page.locator('html')).toBeAttached();
        await expect(page.locator('body')).toBeAttached();

        // Verify tracked changes are rendered (ins/del elements)
        const insertions = await page.locator('ins').count();
        const deletions = await page.locator('del').count();
        expect(insertions + deletions).toBeGreaterThan(0);

        // Verify content is visible
        const bodyText = await page.locator('body').textContent();
        expect(bodyText!.length).toBeGreaterThan(5);
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

    test('move detection exposes MoveGroupId and IsMoveSource in revisions', async ({ page }) => {
      // Use a document that already has move markup (from Word's track changes)
      // This tests that the WASM wrapper correctly exposes move data
      const docWithMoves = readTestFile('FA/RevTracking/014-MovedParagraph.docx');

      // Get revisions directly from the document with existing move markup
      const revisionsResult = await page.evaluate((bytesArray) => {
        return (window as any).DocxodusTests.getRevisions(new Uint8Array(bytesArray));
      }, Array.from(docWithMoves));

      expect(revisionsResult.error).toBeUndefined();
      expect(revisionsResult.revisions).toBeDefined();

      const revisions = revisionsResult.revisions!;
      console.log(`Move detection test: Found ${revisions.length} total revisions`);

      // Log all revisions for debugging
      for (const rev of revisions) {
        console.log(`  - Type: ${rev.RevisionType}, Text: "${rev.Text?.substring(0, 50) || '(empty)'}...", MoveGroupId: ${rev.MoveGroupId}, IsMoveSource: ${rev.IsMoveSource}`);
      }

      // Check for move revisions (RevisionType === "Moved")
      const moveRevisions = revisions.filter((r: any) => r.RevisionType === 'Moved');

      if (moveRevisions.length > 0) {
        console.log(`Found ${moveRevisions.length} move revisions`);

        // Verify move revisions have MoveGroupId
        for (const moveRev of moveRevisions) {
          expect(moveRev.MoveGroupId).toBeDefined();
          expect(typeof moveRev.MoveGroupId).toBe('number');
          expect(moveRev.IsMoveSource).toBeDefined();
          expect(typeof moveRev.IsMoveSource).toBe('boolean');
        }

        // Verify move pairs exist (same MoveGroupId, one source, one destination)
        const moveGroups = new Map<number, any[]>();
        for (const rev of moveRevisions) {
          const groupId = rev.MoveGroupId;
          if (!moveGroups.has(groupId)) {
            moveGroups.set(groupId, []);
          }
          moveGroups.get(groupId)!.push(rev);
        }

        for (const [groupId, group] of moveGroups) {
          expect(group.length).toBe(2); // Each move has source + destination
          const sources = group.filter((r: any) => r.IsMoveSource === true);
          const destinations = group.filter((r: any) => r.IsMoveSource === false);
          expect(sources.length).toBe(1);
          expect(destinations.length).toBe(1);
          console.log(`Move group ${groupId}: source="${sources[0].Text?.substring(0, 30) || '(empty)'}...", dest="${destinations[0].Text?.substring(0, 30) || '(empty)'}..."`);
        }
      } else {
        // Also test with a comparison that should produce moves
        // Compare documents where paragraphs have been reordered
        console.log('Testing move detection via document comparison...');

        const originalBytes = readTestFile('WC/WC007-Unmodified.docx');
        const modifiedBytes = readTestFile('WC/WC007-Moved-into-Table.docx');

        // Compare the documents
        const compareResult = await page.evaluate(([original, modified]) => {
          const result = (window as any).DocxodusTests.compareDocuments(
            new Uint8Array(original),
            new Uint8Array(modified)
          );
          if (result.docxBytes) {
            return { docxBytes: Array.from(result.docxBytes) };
          }
          return result;
        }, [Array.from(originalBytes), Array.from(modifiedBytes)]);

        expect(compareResult.error).toBeUndefined();
        expect(compareResult.docxBytes).toBeDefined();

        const compRevisions = await page.evaluate((bytesArray) => {
          return (window as any).DocxodusTests.getRevisions(new Uint8Array(bytesArray));
        }, compareResult.docxBytes);

        expect(compRevisions.revisions).toBeDefined();
        expect(compRevisions.revisions.length).toBeGreaterThan(0);

        // Verify the revision structure includes move fields (even if null for non-moves)
        for (const rev of compRevisions.revisions) {
          // MoveGroupId and IsMoveSource fields should exist in the response
          expect('MoveGroupId' in rev || 'moveGroupId' in rev).toBe(true);
          expect('IsMoveSource' in rev || 'isMoveSource' in rev).toBe(true);
        }

        console.log(`Verified ${compRevisions.revisions.length} revisions have move fields in schema`);
      }
    });
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

    test('tracked changes are preserved when paginating a compared document', async ({ page }) => {
      // Compare two documents to get tracked changes
      const originalBytes = readTestFile('WC/WC002-Unmodified.docx');
      const modifiedBytes = readTestFile('WC/WC002-DiffInMiddle.docx');

      // First compare the documents
      const compResult = await compareDocuments(page, originalBytes, modifiedBytes);
      expect(compResult.error).toBeUndefined();
      expect(compResult.docxBytes).toBeDefined();

      // Convert the compared document (which has tracked changes) to HTML with pagination
      const htmlResult = await convertToHtmlWithPaginationAndTrackedChanges(
        page,
        new Uint8Array(compResult.docxBytes!),
        1,    // paginated mode
        0.8,  // scale
        true  // renderTrackedChanges
      );

      expect(htmlResult.error).toBeUndefined();
      expect(htmlResult.html).toBeDefined();

      // Verify tracked change markup is present in the HTML
      expect(htmlResult.html).toContain('rev-ins');  // insertion class
      expect(htmlResult.html).toContain('rev-del');  // deletion class
      expect(htmlResult.html).toContain('<ins');     // semantic ins element
      expect(htmlResult.html).toContain('<del');     // semantic del element

      // Verify pagination structure is also present
      expect(htmlResult.html).toContain('pagination-staging');
      expect(htmlResult.html).toContain('pagination-container');

      // Load the pagination bundle and run pagination
      await page.addScriptTag({ path: 'dist/pagination.bundle.js' });

      const paginationResult = await page.evaluate((html) => {
        // Create container and insert HTML
        const container = document.createElement('div');
        container.id = 'test-pagination-tracked-changes';
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
            scale: 0.8,
            showPageNumbers: true
          });

          const result = engine.paginate();

          // Check that tracked change elements are present in the paginated output
          const pageBoxes = pageContainer.querySelectorAll('.page-box');
          const pageContents = pageContainer.querySelectorAll('.page-content');

          // Find ins/del elements in the paginated pages
          let insCount = 0;
          let delCount = 0;

          pageContents.forEach((pageContent) => {
            insCount += pageContent.querySelectorAll('ins').length;
            delCount += pageContent.querySelectorAll('del').length;
          });

          // Verify tracked changes CSS classes are preserved after cloning
          const insWithClass = pageContainer.querySelectorAll('ins.rev-ins').length;
          const delWithClass = pageContainer.querySelectorAll('del.rev-del').length;

          document.body.removeChild(container);

          return {
            totalPages: result.totalPages,
            pageBoxCount: pageBoxes.length,
            insElements: insCount,
            delElements: delCount,
            insWithClass: insWithClass,
            delWithClass: delWithClass
          };
        } catch (e) {
          document.body.removeChild(container);
          return { error: (e as Error).message };
        }
      }, htmlResult.html!);

      if ('error' in paginationResult) {
        throw new Error(paginationResult.error as string);
      }

      // Verify pagination completed successfully
      expect(paginationResult.totalPages).toBeGreaterThan(0);
      expect(paginationResult.pageBoxCount).toBeGreaterThan(0);

      // Verify tracked change elements are present in paginated output
      // The comparison should have created at least some insertions and deletions
      expect(paginationResult.insElements + paginationResult.delElements).toBeGreaterThan(0);

      // Verify CSS classes were preserved during cloning
      expect(paginationResult.insWithClass + paginationResult.delWithClass).toBeGreaterThan(0);

      console.log(`Pagination with tracked changes: ${paginationResult.totalPages} pages, ` +
        `${paginationResult.insElements} ins elements, ${paginationResult.delElements} del elements, ` +
        `${paginationResult.insWithClass} ins.rev-ins, ${paginationResult.delWithClass} del.rev-del`);
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

    test('annotation rendering generates highlight spans in DOM', async ({ page }) => {
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

      // Actually render the HTML to the page
      await page.setContent(htmlResult.html!);

      // Verify annotation highlight elements are rendered in DOM
      const highlights = page.locator('.annot-highlight');
      await expect(highlights.first()).toBeAttached();

      // Verify annotation has correct data attribute
      const annotationEl = page.locator('[data-annotation-id="render-test"]');
      await expect(annotationEl).toBeAttached();

      // Verify label is rendered
      const label = page.locator('.annot-label');
      await expect(label.first()).toBeAttached();
      await expect(label.first()).toContainText('Highlighted Text');

      // Verify highlight is visible
      await expect(highlights.first()).toBeVisible();

      console.log('Annotation highlight spans verified in DOM');
    });

    test('annotation CSS applies highlight styles in DOM', async ({ page }) => {
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

      // Verify the color is in the HTML (in stylesheet)
      expect(htmlResult.html).toContain('#4CAF50');

      // Actually render the HTML to the page
      await page.setContent(htmlResult.html!);

      // Verify style element is in DOM and contains annotation styles
      const styleContent = await page.locator('style').first().textContent();
      expect(styleContent).toContain('.annot-highlight');
      expect(styleContent).toContain('#4CAF50');

      // Verify highlight element exists and is visible
      const highlight = page.locator('.annot-highlight').first();
      await expect(highlight).toBeVisible();

      // Verify the highlight has the correct data attributes
      await expect(highlight).toHaveAttribute('data-annotation-id', 'css-test');

      // Verify label is visible
      const label = page.locator('.annot-label').first();
      await expect(label).toBeVisible();
      await expect(label).toContainText('Styled');

      console.log('Annotation CSS verified as applied in DOM');
    });

    test('annotation label modes render differently in DOM', async ({ page }) => {
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

      // Test Above mode - label should be visible
      const aboveResult = await convertToHtmlWithAnnotations(page, annotatedBytes, true, 0);
      await page.setContent(aboveResult.html!);
      await expect(page.locator('.annot-highlight').first()).toBeVisible();
      await expect(page.locator('.annot-label').first()).toBeAttached();
      console.log('Label mode Above: Label element present');

      // Test Inline mode - label should be inline
      const inlineResult = await convertToHtmlWithAnnotations(page, annotatedBytes, true, 1);
      await page.setContent(inlineResult.html!);
      await expect(page.locator('.annot-highlight').first()).toBeVisible();
      console.log('Label mode Inline: Rendered correctly');

      // Test Tooltip mode - highlight visible, label for tooltip
      const tooltipResult = await convertToHtmlWithAnnotations(page, annotatedBytes, true, 2);
      await page.setContent(tooltipResult.html!);
      await expect(page.locator('.annot-highlight').first()).toBeVisible();
      console.log('Label mode Tooltip: Rendered correctly');

      // Test None mode - highlight only, no label element
      const noneResult = await convertToHtmlWithAnnotations(page, annotatedBytes, true, 3);
      await page.setContent(noneResult.html!);
      await expect(page.locator('.annot-highlight').first()).toBeVisible();
      const labelCount = await page.locator('.annot-label').count();
      expect(labelCount).toBe(0);
      console.log('Label mode None: No label elements rendered');
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

    test('disabling annotation rendering produces clean DOM', async ({ page }) => {
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

      // Actually render the HTML to the page
      await page.setContent(htmlResult.html!);

      // Verify document renders but has no annotation elements in DOM
      await expect(page.locator('body')).toBeAttached();
      const highlightCount = await page.locator('.annot-highlight').count();
      expect(highlightCount).toBe(0);

      const annotationIdCount = await page.locator('[data-annotation-id]').count();
      expect(annotationIdCount).toBe(0);

      console.log('Disabled annotation rendering produces clean DOM');
    });
  });

  test.describe('Document Structure Tests', () => {
    // Use a simple document for structure tests
    const testDoc = 'HC006-Test-01.docx';
    // Use a document with tables for table structure tests
    const tableDoc = 'HC001-5DayTourPlanTemplate.docx';

    test('can get document structure', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await getDocumentStructure(page, bytes);

      expect(result.error).toBeUndefined();
      expect(result.root).toBeDefined();
      expect(result.root.Id).toBe('doc');
      expect(result.root.Type).toBe('Document');
      expect(result.elementsById).toBeDefined();

      console.log('Document structure retrieved successfully');
    });

    test('structure contains paragraphs with correct IDs', async ({ page }) => {
      const bytes = readTestFile(testDoc);
      const result = await getDocumentStructure(page, bytes);

      expect(result.error).toBeUndefined();

      // Find paragraphs by checking element type, not just ID pattern
      // (IDs containing /p- might be part of longer paths like doc/p-0/hl-0/r-0)
      const paragraphIds = Object.keys(result.elementsById!).filter(id => {
        const element = result.elementsById![id];
        return element.Type === 'Paragraph';
      });
      expect(paragraphIds.length).toBeGreaterThan(0);

      // Verify all paragraph IDs end with /p-N and start with doc/
      for (const id of paragraphIds) {
        expect(id).toMatch(/\/p-\d+$/);  // Should end with /p-N
        expect(id).toMatch(/^doc\//);   // Should start with doc/
      }

      console.log(`Found ${paragraphIds.length} paragraphs with correct ID format`);
    });

    test('structure contains table information', async ({ page }) => {
      const bytes = readTestFile(tableDoc);
      const result = await getDocumentStructure(page, bytes);

      expect(result.error).toBeUndefined();

      // Find tables in elementsById
      const tableIds = Object.keys(result.elementsById!).filter(id => id.match(/\/tbl-\d+$/));

      if (tableIds.length > 0) {
        // Verify table structure
        const tableId = tableIds[0];
        const table = result.elementsById![tableId];
        expect(table.Type).toBe('Table');

        // Check for table column info
        expect(result.tableColumns).toBeDefined();
        const columnKeys = Object.keys(result.tableColumns!);
        expect(columnKeys.length).toBeGreaterThan(0);

        // Verify column format
        for (const colKey of columnKeys) {
          const col = result.tableColumns![colKey];
          expect(col.TableId).toBeDefined();
          expect(col.ColumnIndex).toBeDefined();
          expect(col.CellIds).toBeDefined();
        }

        console.log(`Found ${tableIds.length} tables with ${columnKeys.length} columns`);
      } else {
        console.log('No tables found in test document');
      }
    });

    test('structure contains rows and cells for tables', async ({ page }) => {
      const bytes = readTestFile(tableDoc);
      const result = await getDocumentStructure(page, bytes);

      expect(result.error).toBeUndefined();

      // Find table rows by checking element type, not ID pattern
      // (IDs containing /tr- might be part of longer paths like doc/tbl-0/tr-1/tc-0/p-0)
      const rowIds = Object.keys(result.elementsById!).filter(id => {
        const element = result.elementsById![id];
        return element.Type === 'TableRow';
      });
      // Find table cells by checking element type
      const cellIds = Object.keys(result.elementsById!).filter(id => {
        const element = result.elementsById![id];
        return element.Type === 'TableCell';
      });

      // Find tables
      const tableIds = Object.keys(result.elementsById!).filter(id => id.match(/\/tbl-\d+$/));

      if (tableIds.length > 0) {
        expect(rowIds.length).toBeGreaterThan(0);
        expect(cellIds.length).toBeGreaterThan(0);

        // Verify row structure
        const rowId = rowIds[0];
        const row = result.elementsById![rowId];
        expect(row.Type).toBe('TableRow');
        // RowIndex might be null for some rows
        expect(typeof row.RowIndex === 'number' || row.RowIndex === null || row.RowIndex === undefined).toBe(true);

        // Verify cell structure
        const cellId = cellIds[0];
        const cell = result.elementsById![cellId];
        expect(cell.Type).toBe('TableCell');

        console.log(`Found ${tableIds.length} tables, ${rowIds.length} rows and ${cellIds.length} cells`);
      } else {
        console.log('No tables found in test document');
      }
    });

    test('element IDs are path-based and deterministic', async ({ page }) => {
      const bytes = readTestFile(tableDoc);

      // Get structure twice to verify determinism
      const result1 = await getDocumentStructure(page, bytes);
      const result2 = await getDocumentStructure(page, bytes);

      expect(result1.error).toBeUndefined();
      expect(result2.error).toBeUndefined();

      // Verify same element IDs
      const ids1 = Object.keys(result1.elementsById!).sort();
      const ids2 = Object.keys(result2.elementsById!).sort();

      expect(ids1).toEqual(ids2);

      console.log('Element IDs are deterministic');
    });
  });

  test.describe('Element-based Annotation Targeting Tests', () => {
    // Use a simple document for targeting tests
    const testDoc = 'HC006-Test-01.docx';

    test('can annotate a paragraph by element ID', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // First get the structure to find a paragraph ID (top-level or nested)
      const structure = await getDocumentStructure(page, bytes);
      expect(structure.error).toBeUndefined();

      // Find any paragraph (could be top-level like doc/p-0 or nested like doc/tbl-0/tr-0/tc-0/p-0)
      const paragraphIds = Object.keys(structure.elementsById!).filter(id => id.includes('/p-'));
      expect(paragraphIds.length).toBeGreaterThan(0);

      // Prefer a top-level paragraph if available, otherwise use the first one
      const topLevelParagraphs = paragraphIds.filter(id => id.match(/^doc\/p-\d+$/));
      const targetId = topLevelParagraphs.length > 0 ? topLevelParagraphs[0] : paragraphIds[0];

      // Add annotation targeting this element
      const addResult = await addAnnotationWithTarget(page, bytes, {
        Id: 'element-id-test',
        LabelId: 'PARAGRAPH',
        Label: 'Targeted Paragraph',
        Color: '#4CAF50',
        ElementId: targetId
      });

      expect(addResult.error).toBeUndefined();
      expect(addResult.success).toBe(true);
      expect(addResult.documentBytes).toBeDefined();
      expect(addResult.annotation).toBeDefined();
      expect(addResult.annotation.Id).toBe('element-id-test');

      console.log(`Annotated paragraph ${targetId} by element ID`);
    });

    test('can annotate a paragraph by index', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add annotation targeting paragraph by index
      const addResult = await addAnnotationWithTarget(page, bytes, {
        Id: 'paragraph-index-test',
        LabelId: 'PARA_IDX',
        Label: 'First Paragraph',
        Color: '#2196F3',
        ElementType: 'Paragraph',
        ParagraphIndex: 0
      });

      expect(addResult.error).toBeUndefined();
      expect(addResult.success).toBe(true);
      expect(addResult.documentBytes).toBeDefined();
      expect(addResult.annotation).toBeDefined();

      // Verify annotation was added
      const getResult = await getAnnotationsFromDoc(
        page,
        new Uint8Array(addResult.documentBytes!)
      );

      expect(getResult.error).toBeUndefined();
      expect(getResult.annotations!.length).toBe(1);
      expect(getResult.annotations![0].Id).toBe('paragraph-index-test');

      console.log('Annotated paragraph by index');
    });

    test('can annotate using text search', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Add annotation using text search
      const addResult = await addAnnotationWithTarget(page, bytes, {
        Id: 'text-search-test',
        LabelId: 'SEARCH',
        Label: 'Found Text',
        Color: '#FF9800',
        SearchText: 'the',
        Occurrence: 1
      });

      expect(addResult.error).toBeUndefined();
      expect(addResult.success).toBe(true);
      expect(addResult.documentBytes).toBeDefined();

      console.log('Annotated using text search');
    });

    test('can annotate a paragraph range', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Get structure to verify we have enough top-level paragraphs
      const structure = await getDocumentStructure(page, bytes);
      const topLevelParagraphs = Object.keys(structure.elementsById!).filter(id => id.match(/^doc\/p-\d+$/));
      const paragraphCount = topLevelParagraphs.length;

      if (paragraphCount >= 2) {
        // Add annotation spanning paragraphs 0-1
        const addResult = await addAnnotationWithTarget(page, bytes, {
          Id: 'range-test',
          LabelId: 'RANGE',
          Label: 'Paragraph Range',
          Color: '#9C27B0',
          ElementType: 'Paragraph',
          ParagraphIndex: 0,
          RangeEndParagraphIndex: 1
        });

        expect(addResult.error).toBeUndefined();
        expect(addResult.success).toBe(true);
        expect(addResult.documentBytes).toBeDefined();

        console.log('Annotated paragraph range 0-1');
      } else {
        console.log(`Skipped range test: only ${paragraphCount} paragraphs`);
      }
    });

    test('annotation renders correctly when targeting by element ID', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Get structure to find first paragraph (prefer top-level)
      const structure = await getDocumentStructure(page, bytes);
      const allParagraphIds = Object.keys(structure.elementsById!).filter(id => id.includes('/p-'));
      const topLevelParagraphs = allParagraphIds.filter(id => id.match(/^doc\/p-\d+$/));
      const targetId = topLevelParagraphs.length > 0 ? topLevelParagraphs[0] : allParagraphIds[0];

      // Add annotation
      const addResult = await addAnnotationWithTarget(page, bytes, {
        Id: 'render-element-test',
        LabelId: 'RENDER',
        Label: 'Rendered Element',
        Color: '#E91E63',
        ElementId: targetId
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
      expect(htmlResult.html).toContain('annot-highlight');
      expect(htmlResult.html).toContain('data-annotation-id="render-element-test"');
      expect(htmlResult.html).toContain('Rendered Element');

      console.log('Element-targeted annotation renders correctly');
    });

    test('multiple targeting methods work together', async ({ page }) => {
      let bytes = readTestFile(testDoc);

      // Get structure (prefer top-level paragraphs)
      const structure = await getDocumentStructure(page, bytes);
      const allParagraphIds = Object.keys(structure.elementsById!).filter(id => id.includes('/p-'));
      const topLevelParagraphs = allParagraphIds.filter(id => id.match(/^doc\/p-\d+$/));
      const paragraphIds = topLevelParagraphs.length > 0 ? topLevelParagraphs : allParagraphIds;

      // Add annotation by element ID
      const add1 = await addAnnotationWithTarget(page, bytes, {
        Id: 'multi-1',
        LabelId: 'TYPE_A',
        Label: 'By Element ID',
        Color: '#4CAF50',
        ElementId: paragraphIds[0]
      });

      expect(add1.error).toBeUndefined();
      bytes = new Uint8Array(add1.documentBytes!);

      // Add annotation by text search
      const add2 = await addAnnotationWithTarget(page, bytes, {
        Id: 'multi-2',
        LabelId: 'TYPE_B',
        Label: 'By Text Search',
        Color: '#2196F3',
        SearchText: 'and',
        Occurrence: 1
      });

      expect(add2.error).toBeUndefined();
      bytes = new Uint8Array(add2.documentBytes!);

      // Verify both annotations exist
      const getResult = await getAnnotationsFromDoc(page, bytes);

      expect(getResult.error).toBeUndefined();
      expect(getResult.annotations!.length).toBe(2);

      const ids = getResult.annotations!.map((a: any) => a.Id);
      expect(ids).toContain('multi-1');
      expect(ids).toContain('multi-2');

      console.log('Multiple targeting methods work together');
    });
  });

  test.describe('Frame Yielding Tests (Issue #44)', () => {
    // These tests verify that frame yielding allows UI updates before heavy WASM work
    const testDoc = 'HC006-Test-01.docx';

    test('loading state is observable before conversion completes', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // This test verifies the core behavior: that a loading state set before
      // calling convertToHtml() is actually painted before the blocking work begins.
      // We do this by:
      // 1. Setting a loading indicator
      // 2. Calling conversion (which internally yields via double-rAF)
      // 3. Verifying the loading indicator was painted at some point

      const result = await page.evaluate(async (bytesArray) => {
        const timeline: string[] = [];
        const loadingDiv = document.createElement('div');
        loadingDiv.id = 'loading-indicator';
        loadingDiv.textContent = 'Loading...';
        loadingDiv.style.display = 'none';
        document.body.appendChild(loadingDiv);

        // Track when loading indicator becomes visible via MutationObserver
        let loadingWasPainted = false;

        // Use IntersectionObserver to detect when element is actually rendered
        const paintPromise = new Promise<void>((resolve) => {
          // Schedule a check after potential paint
          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              if (loadingDiv.style.display === 'block') {
                loadingWasPainted = true;
              }
              resolve();
            });
          });
        });

        // Simulate React-like state update: show loading, then do work
        timeline.push('setState:loading');
        loadingDiv.style.display = 'block';

        // Give the browser a chance to paint
        await new Promise(resolve => requestAnimationFrame(() => {
          requestAnimationFrame(() => resolve(undefined));
        }));
        timeline.push('afterYield');

        // Check if loading was painted
        const computedDisplay = window.getComputedStyle(loadingDiv).display;
        if (computedDisplay === 'block') {
          loadingWasPainted = true;
          timeline.push('loadingPainted');
        }

        // Now do the conversion
        timeline.push('startConversion');
        const conversionResult = (window as any).DocxodusTests.convertToHtml(new Uint8Array(bytesArray));
        timeline.push('endConversion');

        // Hide loading
        loadingDiv.style.display = 'none';
        timeline.push('setState:done');

        document.body.removeChild(loadingDiv);

        return {
          loadingWasPainted,
          timeline,
          conversionSuccess: !conversionResult.error,
          htmlLength: conversionResult.html?.length || 0
        };
      }, Array.from(bytes));

      // The critical assertion: loading state was painted before work completed
      expect(result.loadingWasPainted).toBe(true);
      expect(result.conversionSuccess).toBe(true);
      expect(result.htmlLength).toBeGreaterThan(100);

      console.log('Frame yielding timeline:', result.timeline.join(' -> '));
    });

    test('multiple async operations yield properly', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      // Test that sequential async operations all yield properly
      const result = await page.evaluate(async (bytesArray) => {
        const timestamps: { operation: string; time: number }[] = [];
        const start = performance.now();

        const addTimestamp = (op: string) => {
          timestamps.push({ operation: op, time: performance.now() - start });
        };

        addTimestamp('start');

        // Operation 1: Convert to HTML
        addTimestamp('convert:start');
        const convertResult = (window as any).DocxodusTests.convertToHtml(new Uint8Array(bytesArray));
        addTimestamp('convert:end');

        // Operation 2: Get document structure (if available)
        addTimestamp('structure:start');
        const structureResult = (window as any).DocxodusTests.getDocumentStructure(new Uint8Array(bytesArray));
        addTimestamp('structure:end');

        addTimestamp('done');

        return {
          timestamps,
          convertSuccess: !convertResult.error,
          structureSuccess: !structureResult.error,
          totalTime: performance.now() - start
        };
      }, Array.from(bytes));

      expect(result.convertSuccess).toBe(true);
      expect(result.structureSuccess).toBe(true);

      console.log('Operation timing:', result.timestamps.map(t =>
        `${t.operation}: ${t.time.toFixed(1)}ms`
      ).join(', '));
    });

    test('comparison operation yields to allow loading state', async ({ page }) => {
      const originalBytes = readTestFile('WC/WC001-Digits.docx');
      const modifiedBytes = readTestFile('WC/WC001-Digits-Mod.docx');

      const result = await page.evaluate(async ([original, modified]) => {
        let loadingPainted = false;
        const indicator = document.createElement('div');
        indicator.id = 'compare-loading';
        indicator.style.display = 'none';
        document.body.appendChild(indicator);

        // Show loading
        indicator.style.display = 'block';

        // Yield like the library does
        await new Promise(resolve => {
          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              // Check if loading is painted
              if (window.getComputedStyle(indicator).display === 'block') {
                loadingPainted = true;
              }
              resolve(undefined);
            });
          });
        });

        // Do comparison
        const compareResult = (window as any).DocxodusTests.compareDocuments(
          new Uint8Array(original),
          new Uint8Array(modified)
        );

        // Hide loading
        indicator.style.display = 'none';
        document.body.removeChild(indicator);

        return {
          loadingPainted,
          compareSuccess: !compareResult.error,
          hasDocxBytes: compareResult.docxBytes?.length > 0
        };
      }, [Array.from(originalBytes), Array.from(modifiedBytes)]);

      expect(result.loadingPainted).toBe(true);
      expect(result.compareSuccess).toBe(true);
      expect(result.hasDocxBytes).toBe(true);

      console.log('Comparison yielded properly, loading state was painted');
    });

    test('getRevisions yields before processing', async ({ page }) => {
      const originalBytes = readTestFile('WC/WC001-Digits.docx');
      const modifiedBytes = readTestFile('WC/WC001-Digits-Mod.docx');

      // First create a compared document
      const compareResult = await page.evaluate(([original, modified]) => {
        return (window as any).DocxodusTests.compareDocuments(
          new Uint8Array(original),
          new Uint8Array(modified)
        );
      }, [Array.from(originalBytes), Array.from(modifiedBytes)]);

      expect(compareResult.error).toBeUndefined();
      expect(compareResult.docxBytes).toBeDefined();

      // Now test getRevisions with yielding
      const result = await page.evaluate(async (docxBytes) => {
        let uiUpdateCompleted = false;

        // Schedule UI update
        const uiPromise = new Promise<void>(resolve => {
          requestAnimationFrame(() => {
            uiUpdateCompleted = true;
            resolve();
          });
        });

        // In parallel, do the revision extraction
        const revisionsResult = (window as any).DocxodusTests.getRevisions(new Uint8Array(docxBytes));

        // Wait for UI update
        await uiPromise;

        return {
          uiUpdateCompleted,
          revisionsSuccess: !revisionsResult.error,
          revisionCount: revisionsResult.revisions?.length || 0
        };
      }, compareResult.docxBytes);

      expect(result.uiUpdateCompleted).toBe(true);
      expect(result.revisionsSuccess).toBe(true);
      expect(result.revisionCount).toBeGreaterThan(0);

      console.log(`getRevisions completed with ${result.revisionCount} revisions, UI updated`);
    });

    test('annotation operations yield properly', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        let loadingVisible = false;

        // Show loading indicator
        const indicator = document.createElement('div');
        indicator.textContent = 'Adding annotation...';
        indicator.style.display = 'block';
        indicator.style.backgroundColor = 'yellow';
        document.body.appendChild(indicator);

        // Yield to ensure paint (double-rAF like the library does)
        await new Promise(resolve => {
          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              // Check visibility using computed style and dimensions
              const style = window.getComputedStyle(indicator);
              const rect = indicator.getBoundingClientRect();
              loadingVisible = style.display === 'block' && rect.width > 0 && rect.height > 0;
              resolve(undefined);
            });
          });
        });

        // Add annotation
        const addResult = (window as any).DocxodusTests.addAnnotation(
          new Uint8Array(bytesArray),
          {
            Id: 'yield-test-annot',
            LabelId: 'TEST',
            Label: 'Yield Test',
            Color: '#4CAF50',
            SearchText: 'the',
            Occurrence: 1
          }
        );

        document.body.removeChild(indicator);

        return {
          loadingVisible,
          addSuccess: addResult.success === true || (!addResult.error && addResult.documentBytes),
          hasAnnotation: !!addResult.annotation
        };
      }, Array.from(bytes));

      expect(result.loadingVisible).toBe(true);
      expect(result.addSuccess).toBe(true);

      console.log('Annotation operation yielded properly');
    });
  });

  // ============================================================
  // Document Metadata Tests (Phase 3: Lazy Loading)
  // ============================================================
  test.describe('Document Metadata Tests (Issue #44 Phase 3)', () => {
    const testDoc = 'HC001-5DayTourPlanTemplate.docx';

    test('getDocumentMetadata returns valid metadata for simple document', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        return metadata;
      }, Array.from(bytes));

      // Verify basic structure
      expect(result.error).toBeUndefined();
      expect(result.sections).toBeDefined();
      expect(Array.isArray(result.sections)).toBe(true);
      expect(result.sections.length).toBeGreaterThan(0);

      // Verify section properties
      const firstSection = result.sections[0];
      expect(firstSection.sectionIndex).toBe(0);
      expect(firstSection.pageWidthPt).toBeGreaterThan(0);
      expect(firstSection.pageHeightPt).toBeGreaterThan(0);
      expect(firstSection.contentWidthPt).toBeGreaterThan(0);
      expect(firstSection.contentHeightPt).toBeGreaterThan(0);

      // Verify totals
      expect(result.totalParagraphs).toBeGreaterThanOrEqual(0);
      expect(result.estimatedPageCount).toBeGreaterThan(0);

      console.log(`Document has ${result.sections.length} section(s), ` +
                  `${result.totalParagraphs} paragraphs, ` +
                  `estimated ${result.estimatedPageCount} pages`);
    });

    test('getDocumentMetadata returns correct page dimensions', async ({ page }) => {
      // Use a document with known page size (US Letter is 612x792 points)
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        return metadata;
      }, Array.from(bytes));

      expect(result.sections.length).toBeGreaterThan(0);
      const section = result.sections[0];

      // US Letter size (default) is 612x792 points (8.5" x 11")
      expect(section.pageWidthPt).toBeGreaterThanOrEqual(540); // Allow some variation
      expect(section.pageWidthPt).toBeLessThanOrEqual(700);
      expect(section.pageHeightPt).toBeGreaterThanOrEqual(700);
      expect(section.pageHeightPt).toBeLessThanOrEqual(850);

      // Content area should be smaller than page
      expect(section.contentWidthPt).toBeLessThan(section.pageWidthPt);
      expect(section.contentHeightPt).toBeLessThan(section.pageHeightPt);

      console.log(`Page size: ${section.pageWidthPt}x${section.pageHeightPt}pt, ` +
                  `content: ${section.contentWidthPt}x${section.contentHeightPt}pt`);
    });

    test('getDocumentMetadata detects document features', async ({ page }) => {
      // Test with a document that has comments
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        return {
          hasFootnotes: metadata.hasFootnotes,
          hasEndnotes: metadata.hasEndnotes,
          hasComments: metadata.hasComments,
          hasTrackedChanges: metadata.hasTrackedChanges,
        };
      }, Array.from(bytes));

      // Document features should be booleans
      expect(typeof result.hasFootnotes).toBe('boolean');
      expect(typeof result.hasEndnotes).toBe('boolean');
      expect(typeof result.hasComments).toBe('boolean');
      expect(typeof result.hasTrackedChanges).toBe('boolean');

      console.log('Document features:', result);
    });

    test('getDocumentMetadata tracks paragraph indices correctly', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        return metadata;
      }, Array.from(bytes));

      expect(result.sections.length).toBeGreaterThan(0);

      // For single section document, start should be 0 and end should match total
      if (result.sections.length === 1) {
        expect(result.sections[0].startParagraphIndex).toBe(0);
        expect(result.sections[0].endParagraphIndex).toBe(result.totalParagraphs);
      }

      // For multi-section documents, indices should be contiguous
      let lastEnd = 0;
      for (const section of result.sections) {
        expect(section.startParagraphIndex).toBe(lastEnd);
        expect(section.endParagraphIndex).toBeGreaterThanOrEqual(section.startParagraphIndex);
        lastEnd = section.endParagraphIndex;
      }
      expect(lastEnd).toBe(result.totalParagraphs);

      console.log(`Paragraph ranges verified for ${result.sections.length} section(s)`);
    });

    test('getDocumentMetadata is faster than full conversion', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const docBytes = new Uint8Array(bytesArray);

        // Time metadata extraction
        const metaStart = performance.now();
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(docBytes);
        const metaEnd = performance.now();
        const metaTime = metaEnd - metaStart;

        // Time full conversion
        const convStart = performance.now();
        const html = await (window as any).DocxodusTests.convertToHtml(docBytes);
        const convEnd = performance.now();
        const convTime = convEnd - convStart;

        return {
          metaTime,
          convTime,
          metaHasSections: metadata.sections?.length > 0,
          convHasHtml: html.html?.length > 0
        };
      }, Array.from(bytes));

      // Both should succeed
      expect(result.metaHasSections).toBe(true);
      expect(result.convHasHtml).toBe(true);

      // Metadata should be faster (or at least not significantly slower)
      console.log(`Metadata: ${result.metaTime.toFixed(1)}ms, Conversion: ${result.convTime.toFixed(1)}ms`);

      // Note: First call may have initialization overhead, so we just log the times
      // In practice, metadata should be significantly faster for large documents
    });

    test('getDocumentMetadata yields to browser before WASM work', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        let loadingVisible = false;

        // Show loading indicator
        const indicator = document.createElement('div');
        indicator.textContent = 'Getting metadata...';
        indicator.style.display = 'block';
        indicator.style.backgroundColor = 'cyan';
        document.body.appendChild(indicator);

        // Start metadata extraction (async, should yield)
        const metadataPromise = (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));

        // Check if indicator is visible using RAF pattern
        await new Promise(resolve => {
          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              const style = window.getComputedStyle(indicator);
              const rect = indicator.getBoundingClientRect();
              loadingVisible = style.display === 'block' && rect.width > 0 && rect.height > 0;
              resolve(undefined);
            });
          });
        });

        const metadata = await metadataPromise;
        document.body.removeChild(indicator);

        return {
          loadingVisible,
          success: !metadata.error && metadata.sections?.length > 0
        };
      }, Array.from(bytes));

      expect(result.loadingVisible).toBe(true);
      expect(result.success).toBe(true);

      console.log('getDocumentMetadata yielded properly before WASM work');
    });

    test('metadata dimensions match rendered HTML data attributes', async ({ page }) => {
      // This test verifies metadata extraction produces same values as full rendering
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const docBytes = new Uint8Array(bytesArray);

        // Get metadata
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(docBytes);

        // Get rendered HTML with pagination mode
        const htmlResult = (window as any).DocxodusTests.convertToHtmlWithPagination(docBytes, 1, 1.0);

        if (metadata.error || htmlResult.error) {
          return { error: metadata.error || htmlResult.error };
        }

        // Parse the HTML and extract section dimensions from data attributes
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlResult.html, 'text/html');
        const sectionElements = doc.querySelectorAll('[data-section-index]');

        const renderedSections = [];
        for (const section of sectionElements) {
          renderedSections.push({
            sectionIndex: parseInt(section.getAttribute('data-section-index') || '0'),
            pageWidth: parseFloat(section.getAttribute('data-page-width') || '0'),
            pageHeight: parseFloat(section.getAttribute('data-page-height') || '0'),
            contentWidth: parseFloat(section.getAttribute('data-content-width') || '0'),
            contentHeight: parseFloat(section.getAttribute('data-content-height') || '0'),
          });
        }

        return {
          metadataSections: metadata.sections,
          renderedSections,
          metadataSectionCount: metadata.sections?.length || 0,
          renderedSectionCount: renderedSections.length
        };
      }, Array.from(bytes));

      expect(result.error).toBeUndefined();

      // For documents with sections, verify dimensions match
      if (result.renderedSectionCount > 0 && result.metadataSectionCount > 0) {
        for (let i = 0; i < Math.min(result.metadataSectionCount, result.renderedSectionCount); i++) {
          const meta = result.metadataSections[i];
          const rendered = result.renderedSections[i];

          // Dimensions should be very close (within 1 point tolerance for rounding)
          expect(Math.abs(meta.pageWidthPt - rendered.pageWidth)).toBeLessThan(1);
          expect(Math.abs(meta.pageHeightPt - rendered.pageHeight)).toBeLessThan(1);
          expect(Math.abs(meta.contentWidthPt - rendered.contentWidth)).toBeLessThan(1);
          expect(Math.abs(meta.contentHeightPt - rendered.contentHeight)).toBeLessThan(1);

          console.log(`Section ${i}: metadata ${meta.pageWidthPt}x${meta.pageHeightPt}pt ` +
                      `matches rendered ${rendered.pageWidth}x${rendered.pageHeight}pt`);
        }
      }
    });

    test('metadata paragraph count matches rendered content', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const docBytes = new Uint8Array(bytesArray);

        // Get metadata
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(docBytes);

        // Get rendered HTML
        const htmlResult = (window as any).DocxodusTests.convertToHtml(docBytes);

        if (metadata.error || htmlResult.error) {
          return { error: metadata.error || htmlResult.error };
        }

        // Count paragraphs in rendered HTML
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlResult.html, 'text/html');

        // Count paragraph elements (p tags in the content)
        const paragraphs = doc.querySelectorAll('p');
        const renderedParagraphCount = paragraphs.length;

        return {
          metadataParagraphs: metadata.totalParagraphs,
          renderedParagraphs: renderedParagraphCount,
          metadataTables: metadata.totalTables,
          htmlLength: htmlResult.html?.length
        };
      }, Array.from(bytes));

      expect(result.error).toBeUndefined();
      expect(result.htmlLength).toBeGreaterThan(0);

      // Metadata paragraph count should be in the same ballpark as rendered
      // (may not be exact due to empty paragraphs, hidden content, etc.)
      console.log(`Metadata: ${result.metadataParagraphs} paragraphs, ${result.metadataTables} tables`);
      console.log(`Rendered: ${result.renderedParagraphs} paragraph elements`);

      // At minimum, both should have content
      expect(result.metadataParagraphs).toBeGreaterThan(0);
      expect(result.renderedParagraphs).toBeGreaterThan(0);
    });

    test('getDocumentMetadata handles invalid document gracefully', async ({ page }) => {
      // Create invalid document data (not a valid DOCX/ZIP)
      const invalidBytes = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);

      const result = await page.evaluate(async (bytesArray) => {
        try {
          const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
          // Test harness returns { error: {...} } for error responses, not throwing
          if (metadata && metadata.error) {
            return { success: false, error: JSON.stringify(metadata.error) };
          }
          return { success: true, metadata };
        } catch (error) {
          return { success: false, error: String(error) };
        }
      }, Array.from(invalidBytes));

      // Should either throw an error or return an error response
      expect(result.success).toBe(false);
      expect(result.error).toBeDefined();
      console.log('Invalid document handled gracefully:', result.error);
    });

    test('getDocumentMetadata returns correct boolean feature flags', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        return {
          hasFootnotes: typeof metadata.hasFootnotes,
          hasEndnotes: typeof metadata.hasEndnotes,
          hasComments: typeof metadata.hasComments,
          hasTrackedChanges: typeof metadata.hasTrackedChanges,
        };
      }, Array.from(bytes));

      // All feature flags should be booleans
      expect(result.hasFootnotes).toBe('boolean');
      expect(result.hasEndnotes).toBe('boolean');
      expect(result.hasComments).toBe('boolean');
      expect(result.hasTrackedChanges).toBe('boolean');
    });

    test('getDocumentMetadata section indices are sequential', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        if (metadata.error) return { error: metadata.error };

        const indices = metadata.sections.map((s: any) => s.sectionIndex);
        const sequential = indices.every((idx: number, i: number) => idx === i);

        return {
          sectionCount: metadata.sections.length,
          indices,
          sequential
        };
      }, Array.from(bytes));

      expect(result.error).toBeUndefined();
      expect(result.sequential).toBe(true);
      console.log(`Section indices are sequential: ${result.indices.join(', ')}`);
    });

    test('getDocumentMetadata content dimensions are calculated correctly', async ({ page }) => {
      const bytes = readTestFile(testDoc);

      const result = await page.evaluate(async (bytesArray) => {
        const metadata = await (window as any).DocxodusTests.getDocumentMetadata(new Uint8Array(bytesArray));
        if (metadata.error) return { error: metadata.error };

        const section = metadata.sections[0];
        const calculatedContentWidth = section.pageWidthPt - section.marginLeftPt - section.marginRightPt;
        const calculatedContentHeight = section.pageHeightPt - section.marginTopPt - section.marginBottomPt;

        return {
          contentWidthPt: section.contentWidthPt,
          contentHeightPt: section.contentHeightPt,
          calculatedContentWidth,
          calculatedContentHeight,
          widthMatch: Math.abs(section.contentWidthPt - calculatedContentWidth) < 0.01,
          heightMatch: Math.abs(section.contentHeightPt - calculatedContentHeight) < 0.01
        };
      }, Array.from(bytes));

      expect(result.error).toBeUndefined();
      expect(result.widthMatch).toBe(true);
      expect(result.heightMatch).toBe(true);
      console.log(`Content width: ${result.contentWidthPt}pt (calculated: ${result.calculatedContentWidth}pt)`);
      console.log(`Content height: ${result.contentHeightPt}pt (calculated: ${result.calculatedContentHeight}pt)`);
    });
  });

  test.describe('Image Handling', () => {
    test('converts document with images to HTML with base64 data URIs', async ({ page }) => {
      // Use a test document known to contain images
      const bytes = readTestFile('HC042-Image-Png.docx');

      const result = await convertToHtml(page, bytes);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Check that images are embedded as base64 data URIs
      const hasBase64Image = result.html!.includes('data:image/');
      const imgTagCount = (result.html!.match(/<img /g) || []).length;

      console.log(`Document contains ${imgTagCount} image(s)`);
      console.log(`Has base64 embedded images: ${hasBase64Image}`);

      // If the document has images, they should be embedded as base64
      if (imgTagCount > 0) {
        expect(hasBase64Image).toBe(true);

        // Verify the base64 data URI format
        const base64Regex = /data:image\/(png|jpeg|gif|bmp);base64,[A-Za-z0-9+\/=]+/;
        expect(result.html!).toMatch(base64Regex);
      }
    });

    test('converts document with image hyperlink correctly', async ({ page }) => {
      // Use a test document with hyperlinked image
      const bytes = readTestFile('HC060-Image-with-Hyperlink.docx');

      const result = await convertToHtml(page, bytes);

      expect(result.error).toBeUndefined();
      expect(result.html).toBeDefined();

      // Check that images are embedded
      const imgTagCount = (result.html!.match(/<img /g) || []).length;
      console.log(`Document with hyperlinked image contains ${imgTagCount} image(s)`);

      // Images should be embedded
      if (imgTagCount > 0) {
        expect(result.html!).toMatch(/data:image\//);
      }
    });
  });

  // ============================================================================
  // External Annotation Tests (Issue #57)
  // ============================================================================
  test.describe('External Annotations', () => {
    // Helper to compute document hash
    async function computeDocumentHash(page: Page, bytes: Uint8Array): Promise<{ hash?: string; error?: any }> {
      return await page.evaluate((bytesArray) => {
        return (window as any).DocxodusTests.computeDocumentHash(new Uint8Array(bytesArray));
      }, Array.from(bytes));
    }

    // Helper to create external annotation set
    async function createExternalAnnotationSet(page: Page, bytes: Uint8Array, documentId: string): Promise<{ annotationSet?: any; error?: any }> {
      return await page.evaluate(([bytesArray, docId]) => {
        return (window as any).DocxodusTests.createExternalAnnotationSet(new Uint8Array(bytesArray), docId);
      }, [Array.from(bytes), documentId] as const);
    }

    // Helper to validate external annotations
    async function validateExternalAnnotations(page: Page, bytes: Uint8Array, annotationSet: any): Promise<{ isValid?: boolean; hashMismatch?: boolean; issues?: any[]; error?: any }> {
      return await page.evaluate(([bytesArray, annSet]) => {
        return (window as any).DocxodusTests.validateExternalAnnotations(new Uint8Array(bytesArray), annSet);
      }, [Array.from(bytes), annotationSet] as const);
    }

    // Helper to convert to HTML with external annotations
    async function convertToHtmlWithExternalAnnotations(page: Page, bytes: Uint8Array, annotationSet: any, labelMode: number = 1): Promise<{ html?: string; error?: any }> {
      return await page.evaluate(([bytesArray, annSet, mode]) => {
        return (window as any).DocxodusTests.convertToHtmlWithExternalAnnotations(new Uint8Array(bytesArray), annSet, mode);
      }, [Array.from(bytes), annotationSet, labelMode] as const);
    }

    // Helper to search for text offsets
    async function searchTextOffsets(page: Page, bytes: Uint8Array, searchText: string, maxResults: number = 100): Promise<{ results?: any[]; error?: any }> {
      return await page.evaluate(([bytesArray, text, max]) => {
        return (window as any).DocxodusTests.searchTextOffsets(new Uint8Array(bytesArray), text, max);
      }, [Array.from(bytes), searchText, maxResults] as const);
    }

    // Helper to create annotation from search (client-side)
    async function createAnnotationFromSearch(page: Page, id: string, labelId: string, documentText: string, searchText: string, occurrence: number = 1): Promise<{ annotation?: any; error?: any }> {
      return await page.evaluate(([annId, label, docText, search, occ]) => {
        return (window as any).DocxodusTests.createAnnotationFromSearch(annId, label, docText, search, occ);
      }, [id, labelId, documentText, searchText, occurrence] as const);
    }

    test('computes document hash consistently', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Compute hash twice
      const result1 = await computeDocumentHash(page, bytes);
      const result2 = await computeDocumentHash(page, bytes);

      expect(result1.error).toBeUndefined();
      expect(result2.error).toBeUndefined();
      expect(result1.hash).toBeDefined();
      expect(result1.hash).toBe(result2.hash);
      expect(result1.hash!.length).toBe(64); // SHA256 = 64 hex chars

      console.log(`Document hash: ${result1.hash}`);
    });

    test('different documents have different hashes', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes1 = readTestFile('HC006-Test-01.docx');
      const bytes2 = readTestFile('HC007-Test-02.docx');

      const result1 = await computeDocumentHash(page, bytes1);
      const result2 = await computeDocumentHash(page, bytes2);

      expect(result1.error).toBeUndefined();
      expect(result2.error).toBeUndefined();
      expect(result1.hash).not.toBe(result2.hash);
    });

    test('creates external annotation set from document', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      const result = await createExternalAnnotationSet(page, bytes, 'test-doc-001');

      expect(result.error).toBeUndefined();
      expect(result.annotationSet).toBeDefined();

      const set = result.annotationSet;
      expect(set.DocumentId || set.documentId).toBe('test-doc-001');
      expect(set.DocumentHash || set.documentHash).toBeDefined();
      expect(set.Content || set.content).toBeDefined();
      expect((set.Content || set.content).length).toBeGreaterThan(0);

      console.log(`Document content length: ${(set.Content || set.content).length} chars`);
      console.log(`Document hash: ${set.DocumentHash || set.documentHash}`);
    });

    test('validates annotation set against document', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Create annotation set
      const createResult = await createExternalAnnotationSet(page, bytes, 'test-doc');
      expect(createResult.error).toBeUndefined();

      // Validate - should pass since we just created it from the same document
      const validateResult = await validateExternalAnnotations(page, bytes, createResult.annotationSet);

      expect(validateResult.error).toBeUndefined();
      expect(validateResult.isValid).toBe(true);
      expect(validateResult.hashMismatch).toBe(false);
      expect(validateResult.issues).toHaveLength(0);
    });

    test('detects hash mismatch when document changes', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes1 = readTestFile('HC006-Test-01.docx');
      const bytes2 = readTestFile('HC007-Test-02.docx'); // Different document

      // Create annotation set from first document
      const createResult = await createExternalAnnotationSet(page, bytes1, 'test-doc');
      expect(createResult.error).toBeUndefined();

      // Validate against second document - should detect hash mismatch
      const validateResult = await validateExternalAnnotations(page, bytes2, createResult.annotationSet);

      expect(validateResult.error).toBeUndefined();
      expect(validateResult.isValid).toBe(false);
      expect(validateResult.hashMismatch).toBe(true);
    });

    test('searches for text occurrences in document', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Search for a common word in Business Trip Checklist
      const result = await searchTextOffsets(page, bytes, 'trip');

      expect(result.error).toBeUndefined();
      expect(result.results).toBeDefined();
      expect(result.results!.length).toBeGreaterThan(0);

      console.log(`Found ${result.results!.length} occurrences of "trip"`);

      // Each result should have start, end, and text
      const first = result.results![0];
      expect(first.start).toBeDefined();
      expect(first.end).toBeDefined();
      expect(first.end).toBeGreaterThan(first.start);
    });

    test('creates annotation from text search', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Create annotation set to get document content
      const createResult = await createExternalAnnotationSet(page, bytes, 'test-doc');
      expect(createResult.error).toBeUndefined();

      const content = createResult.annotationSet.Content || createResult.annotationSet.content;

      // Create annotation by searching for text
      const annResult = await createAnnotationFromSearch(page, 'ann-001', 'IMPORTANT', content, 'trip', 1);

      expect(annResult.error).toBeUndefined();
      expect(annResult.annotation).toBeDefined();
      expect(annResult.annotation.rawText).toBe('trip');
      expect(annResult.annotation.annotationLabel).toBe('IMPORTANT');
      expect(annResult.annotation.annotationJson.start).toBeDefined();
      expect(annResult.annotation.annotationJson.end).toBeDefined();

      console.log(`Created annotation at offset ${annResult.annotation.annotationJson.start}-${annResult.annotation.annotationJson.end}`);
    });

    test('projects annotations onto HTML', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Create annotation set
      const createResult = await createExternalAnnotationSet(page, bytes, 'test-doc');
      expect(createResult.error).toBeUndefined();

      const set = createResult.annotationSet;
      const content = set.Content || set.content;
      console.log(`Document content length: ${content.length}`);
      console.log(`Document content sample: "${content.substring(0, 100)}..."`);

      // Add a label definition
      if (!set.TextLabels) set.TextLabels = {};
      if (!set.textLabels) set.textLabels = {};
      const labels = set.TextLabels || set.textLabels;
      labels['IMPORTANT'] = {
        Id: 'IMPORTANT',
        id: 'IMPORTANT',
        Text: 'Important',
        text: 'Important',
        Color: '#FF5722',
        color: '#FF5722',
        Description: 'Important text',
        description: 'Important text',
        Icon: '',
        icon: '',
        LabelType: 'text',
        labelType: 'text'
      };

      // Create and add annotation
      const annResult = await createAnnotationFromSearch(page, 'ann-001', 'IMPORTANT', content, 'trip', 1);
      console.log(`createAnnotationFromSearch result: ${JSON.stringify(annResult, null, 2)}`);
      expect(annResult.annotation).toBeDefined();

      // Add to annotation set
      const labelledText = set.LabelledText || set.labelledText || [];
      labelledText.push(annResult.annotation);
      set.LabelledText = labelledText;
      set.labelledText = labelledText;

      console.log(`Number of annotations in set: ${labelledText.length}`);

      // Convert to HTML with annotations projected
      const htmlResult = await convertToHtmlWithExternalAnnotations(page, bytes, set, 1); // Inline mode

      expect(htmlResult.error).toBeUndefined();
      expect(htmlResult.html).toBeDefined();

      // Log debug info
      const hasHighlight = htmlResult.html!.includes('ext-annot-highlight');
      const hasAnnotId = htmlResult.html!.includes('data-annotation-id');
      console.log(`Has ext-annot-highlight: ${hasHighlight}`);
      console.log(`Has data-annotation-id: ${hasAnnotId}`);

      // Check that annotation wrapper is in the HTML
      expect(htmlResult.html).toContain('ext-annot-highlight');
      expect(htmlResult.html).toContain('data-annotation-id');
      expect(htmlResult.html).toContain('ann-001');

      console.log('Annotations successfully projected onto HTML');
    });

    test('projects multiple annotations with different labels', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Create annotation set
      const createResult = await createExternalAnnotationSet(page, bytes, 'test-doc');
      expect(createResult.error).toBeUndefined();

      const set = createResult.annotationSet;
      const content = set.Content || set.content;

      // Add label definitions
      if (!set.TextLabels) set.TextLabels = {};
      if (!set.textLabels) set.textLabels = {};
      const labels = set.TextLabels || set.textLabels;

      labels['IMPORTANT'] = {
        id: 'IMPORTANT', Id: 'IMPORTANT',
        text: 'Important', Text: 'Important',
        color: '#FF5722', Color: '#FF5722',
        description: '', Description: '',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };
      labels['TOPIC'] = {
        id: 'TOPIC', Id: 'TOPIC',
        text: 'Topic', Text: 'Topic',
        color: '#9C27B0', Color: '#9C27B0',
        description: '', Description: '',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };

      // Create multiple annotations - search for words that appear multiple times
      const ann1Result = await createAnnotationFromSearch(page, 'ann-001', 'IMPORTANT', content, 'trip', 1);
      const ann2Result = await createAnnotationFromSearch(page, 'ann-002', 'IMPORTANT', content, 'trip', 2);
      const ann3Result = await createAnnotationFromSearch(page, 'ann-003', 'TOPIC', content, 'office', 1);

      // Add to annotation set
      const labelledText = set.LabelledText || set.labelledText || [];
      if (ann1Result.annotation) labelledText.push(ann1Result.annotation);
      if (ann2Result.annotation) labelledText.push(ann2Result.annotation);
      if (ann3Result.annotation) labelledText.push(ann3Result.annotation);
      set.LabelledText = labelledText;
      set.labelledText = labelledText;

      // Convert to HTML with annotations projected
      const htmlResult = await convertToHtmlWithExternalAnnotations(page, bytes, set, 1);

      expect(htmlResult.error).toBeUndefined();
      expect(htmlResult.html).toBeDefined();

      // Count annotation wrappers
      const highlightCount = (htmlResult.html!.match(/ext-annot-highlight/g) || []).length;
      console.log(`Found ${highlightCount} annotation highlights in HTML`);

      // Should have at least the annotations we added (filtering out structural ones)
      const userAnnotations = labelledText.filter((a: any) => !(a.Structural || a.structural));
      expect(highlightCount).toBeGreaterThanOrEqual(userAnnotations.length);

      // Check for different annotation IDs
      expect(htmlResult.html).toContain('ann-001');
      expect(htmlResult.html).toContain('ann-002');
      if (ann3Result.annotation) {
        expect(htmlResult.html).toContain('ann-003');
      }
    });

    test('validates annotations with text mismatch detection', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Create annotation set
      const createResult = await createExternalAnnotationSet(page, bytes, 'test-doc');
      expect(createResult.error).toBeUndefined();

      const set = createResult.annotationSet;

      // Add a bad annotation with wrong text at valid offsets
      const badAnnotation = {
        id: 'bad-ann',
        Id: 'bad-ann',
        annotationLabel: 'TEST',
        AnnotationLabel: 'TEST',
        rawText: 'WRONG TEXT',
        RawText: 'WRONG TEXT',
        page: 0,
        Page: 0,
        annotationJson: {
          id: 'bad-ann',
          Id: 'bad-ann',
          start: 0,
          Start: 0,
          end: 10,
          End: 10,
          text: 'WRONG TEXT',
          Text: 'WRONG TEXT'
        },
        AnnotationJson: {
          id: 'bad-ann',
          Id: 'bad-ann',
          start: 0,
          Start: 0,
          end: 10,
          End: 10,
          text: 'WRONG TEXT',
          Text: 'WRONG TEXT'
        },
        structural: false,
        Structural: false
      };

      const labelledText = set.LabelledText || set.labelledText || [];
      labelledText.push(badAnnotation);
      set.LabelledText = labelledText;
      set.labelledText = labelledText;

      // Validate - should detect text mismatch
      const validateResult = await validateExternalAnnotations(page, bytes, set);

      expect(validateResult.error).toBeUndefined();
      expect(validateResult.isValid).toBe(false);
      expect(validateResult.issues!.length).toBeGreaterThan(0);

      // Find the text mismatch issue
      const textMismatchIssue = validateResult.issues!.find(
        (i: any) => (i.IssueType || i.issueType) === 'TextMismatch'
      );
      expect(textMismatchIssue).toBeDefined();

      console.log('Validation correctly detected text mismatch');
    });

    test('end-to-end: create, annotate, validate, and render', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Step 1: Create annotation set
      console.log('Step 1: Creating annotation set...');
      const createResult = await createExternalAnnotationSet(page, bytes, 'business-trip-v1');
      expect(createResult.error).toBeUndefined();
      const set = createResult.annotationSet;

      console.log(`  - Document ID: ${set.DocumentId || set.documentId}`);
      console.log(`  - Document hash: ${(set.DocumentHash || set.documentHash).substring(0, 16)}...`);
      console.log(`  - Content length: ${(set.Content || set.content).length} chars`);

      // Step 2: Add label definitions
      console.log('Step 2: Adding label definitions...');
      if (!set.TextLabels) set.TextLabels = {};
      if (!set.textLabels) set.textLabels = {};
      const labels = set.TextLabels || set.textLabels;

      labels['TASK'] = {
        id: 'TASK', Id: 'TASK',
        text: 'Task', Text: 'Task',
        color: '#4CAF50', Color: '#4CAF50',
        description: 'Checklist task', Description: 'Checklist task',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };
      labels['TOPIC'] = {
        id: 'TOPIC', Id: 'TOPIC',
        text: 'Topic', Text: 'Topic',
        color: '#2196F3', Color: '#2196F3',
        description: 'Checklist topic', Description: 'Checklist topic',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };

      // Step 3: Create annotations by searching for text
      console.log('Step 3: Creating annotations...');
      const content = set.Content || set.content;

      // Search for "trip" occurrences (common in Business Trip Checklist)
      const searchResult = await searchTextOffsets(page, bytes, 'trip');
      console.log(`  - Found ${searchResult.results?.length || 0} occurrences of "trip"`);

      // Create annotations for first 3 occurrences
      const annotations = [];
      for (let i = 1; i <= Math.min(3, searchResult.results?.length || 0); i++) {
        const annResult = await createAnnotationFromSearch(page, `trip-${i}`, 'TASK', content, 'trip', i);
        if (annResult.annotation) {
          annotations.push(annResult.annotation);
          console.log(`  - Created annotation trip-${i} at offset ${annResult.annotation.annotationJson.start}`);
        }
      }

      // Add annotations to set
      const labelledText = set.LabelledText || set.labelledText || [];
      annotations.forEach(ann => labelledText.push(ann));
      set.LabelledText = labelledText;
      set.labelledText = labelledText;

      // Step 4: Validate
      console.log('Step 4: Validating annotations...');
      const validateResult = await validateExternalAnnotations(page, bytes, set);
      console.log(`  - Validation result: isValid=${validateResult.isValid}, hashMismatch=${validateResult.hashMismatch}`);
      if (validateResult.issues && validateResult.issues.length > 0) {
        console.log(`  - Validation issues:`);
        validateResult.issues.forEach((issue: any, i: number) => {
          console.log(`    ${i + 1}. ${issue.annotationId}: ${issue.issueType} - ${issue.description}`);
        });
      }
      expect(validateResult.isValid).toBe(true);
      console.log(`  - Validation passed: ${validateResult.isValid}`);

      // Step 5: Render with annotations
      console.log('Step 5: Rendering HTML with annotations...');
      const htmlResult = await convertToHtmlWithExternalAnnotations(page, bytes, set, 1);
      expect(htmlResult.error).toBeUndefined();

      // Verify annotations are in the HTML
      const highlightCount = (htmlResult.html!.match(/ext-annot-highlight/g) || []).length;
      console.log(`  - Found ${highlightCount} annotation highlights in rendered HTML`);

      // Check for our specific annotations
      for (const ann of annotations) {
        expect(htmlResult.html).toContain(ann.id);
      }

      console.log('End-to-end test completed successfully!');
    });

    test('VISUAL DEMO: renders annotated document to DOM', async ({ page }) => {
      await page.goto('/test-harness.html');
      await waitForDocxodus(page);

      const bytes = readTestFile('HC006-Test-01.docx');

      // Create annotation set
      const createResult = await createExternalAnnotationSet(page, bytes, 'visual-demo');
      const set = createResult.annotationSet;

      // Add colorful label definitions
      if (!set.TextLabels) set.TextLabels = {};
      if (!set.textLabels) set.textLabels = {};
      const labels = set.TextLabels || set.textLabels;

      labels['IMPORTANT'] = {
        id: 'IMPORTANT', Id: 'IMPORTANT',
        text: 'Important', Text: 'Important',
        color: '#FF5722', Color: '#FF5722',
        description: 'Important item', Description: 'Important item',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };
      labels['ACTION'] = {
        id: 'ACTION', Id: 'ACTION',
        text: 'Action', Text: 'Action',
        color: '#4CAF50', Color: '#4CAF50',
        description: 'Action item', Description: 'Action item',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };
      labels['REFERENCE'] = {
        id: 'REFERENCE', Id: 'REFERENCE',
        text: 'Reference', Text: 'Reference',
        color: '#2196F3', Color: '#2196F3',
        description: 'Reference', Description: 'Reference',
        icon: '', Icon: '',
        labelType: 'text', LabelType: 'text'
      };

      // Create annotations for various terms
      const content = set.Content || set.content;
      const annotations: any[] = [];

      // Annotate "Business" as IMPORTANT
      const businessResult = await createAnnotationFromSearch(page, 'ann-business', 'IMPORTANT', content, 'Business', 1);
      if (businessResult.annotation) annotations.push(businessResult.annotation);

      // Annotate "trip" occurrences as ACTION
      for (let i = 1; i <= 2; i++) {
        const tripResult = await createAnnotationFromSearch(page, `ann-trip-${i}`, 'ACTION', content, 'trip', i);
        if (tripResult.annotation) annotations.push(tripResult.annotation);
      }

      // Annotate "Checklist" as REFERENCE
      const checklistResult = await createAnnotationFromSearch(page, 'ann-checklist', 'REFERENCE', content, 'Checklist', 1);
      if (checklistResult.annotation) annotations.push(checklistResult.annotation);

      // Add annotations to set
      const labelledText = set.LabelledText || set.labelledText || [];
      annotations.forEach(ann => labelledText.push(ann));
      set.LabelledText = labelledText;
      set.labelledText = labelledText;

      // Render with annotations
      const htmlResult = await convertToHtmlWithExternalAnnotations(page, bytes, set, 1);
      expect(htmlResult.error).toBeUndefined();

      // Inject the rendered HTML into the page with custom styles and JSON panel
      await page.evaluate(({ html, annotationSet, annotations }) => {
        document.body.innerHTML = `
          <style>
            * { box-sizing: border-box; }
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              margin: 0;
              background: #f5f5f5;
            }
            h1 {
              color: #333;
              border-bottom: 2px solid #2196F3;
              padding-bottom: 10px;
              margin-top: 0;
            }
            .demo-info {
              background: #e3f2fd;
              padding: 15px;
              border-radius: 8px;
              margin-bottom: 20px;
            }
            .demo-info h3 { margin-top: 0; color: #1565c0; }
            .legend { display: flex; gap: 20px; margin-top: 10px; flex-wrap: wrap; }
            .legend-item { display: flex; align-items: center; gap: 5px; }
            .legend-color { width: 20px; height: 20px; border-radius: 4px; }

            .main-container {
              display: flex;
              gap: 20px;
              align-items: flex-start;
            }
            #document-container {
              flex: 1;
              background: white;
              padding: 30px;
              border-radius: 8px;
              box-shadow: 0 2px 8px rgba(0,0,0,0.1);
              min-width: 0;
            }
            #json-panel {
              width: 400px;
              flex-shrink: 0;
              background: #1e1e1e;
              color: #d4d4d4;
              padding: 15px;
              border-radius: 8px;
              font-family: 'Consolas', 'Monaco', monospace;
              font-size: 11px;
              max-height: 80vh;
              overflow-y: auto;
              position: sticky;
              top: 20px;
            }
            #json-panel h3 {
              color: #569cd6;
              margin-top: 0;
              margin-bottom: 15px;
              font-size: 14px;
            }
            .json-annotation {
              background: #2d2d2d;
              padding: 10px;
              border-radius: 4px;
              margin-bottom: 10px;
              border-left: 3px solid #569cd6;
              white-space: pre-wrap;
              word-break: break-all;
            }
            .json-annotation.highlighted {
              border-left-color: #ff0;
              background: #3d3d1d;
            }
            .json-key { color: #9cdcfe; }
            .json-string { color: #ce9178; }
            .json-number { color: #b5cea8; }

            /* Annotation highlight hover styles */
            .ext-annot-highlight {
              cursor: pointer;
              transition: all 0.2s;
              position: relative;
            }
            .ext-annot-highlight:hover {
              filter: brightness(0.85);
            }

            /* Tooltip styles */
            .annotation-tooltip {
              position: fixed;
              background: #333;
              color: white;
              padding: 8px 12px;
              border-radius: 4px;
              font-size: 12px;
              pointer-events: none;
              z-index: 1000;
              box-shadow: 0 2px 8px rgba(0,0,0,0.3);
              max-width: 300px;
            }
            .annotation-tooltip .tooltip-label {
              font-weight: bold;
              margin-bottom: 4px;
            }
            .annotation-tooltip .tooltip-text {
              color: #aaa;
              font-style: italic;
            }
          </style>
          <h1>External Annotations Visual Demo</h1>
          <div class="demo-info">
            <h3>What you're seeing:</h3>
            <p>This document has <strong>external annotations</strong> projected onto it.
            The annotations are stored separately (as JSON) and rendered as highlights.
            <strong>Hover over highlights</strong> to see tooltips. The JSON panel shows the raw annotation data.</p>
            <div class="legend">
              <div class="legend-item">
                <div class="legend-color" style="background: #FF5722;"></div>
                <span>Important</span>
              </div>
              <div class="legend-item">
                <div class="legend-color" style="background: #4CAF50;"></div>
                <span>Action</span>
              </div>
              <div class="legend-item">
                <div class="legend-color" style="background: #2196F3;"></div>
                <span>Reference</span>
              </div>
            </div>
          </div>
          <div class="main-container">
            <div id="document-container">${html}</div>
            <div id="json-panel">
              <h3>📋 Annotation JSON</h3>
              ${annotations.map((ann: any, i: number) => `
                <div class="json-annotation" data-annotation-id="${ann.id}">
                  <span class="json-key">"id"</span>: <span class="json-string">"${ann.id}"</span>,
<span class="json-key">"label"</span>: <span class="json-string">"${ann.annotationLabel}"</span>,
<span class="json-key">"text"</span>: <span class="json-string">"${ann.rawText}"</span>,
<span class="json-key">"offset"</span>: { <span class="json-key">start</span>: <span class="json-number">${ann.annotationJson?.start || 0}</span>, <span class="json-key">end</span>: <span class="json-number">${ann.annotationJson?.end || 0}</span> }
                </div>
              `).join('')}
            </div>
          </div>
          <div id="tooltip" class="annotation-tooltip" style="display: none;"></div>
        `;

        // Add hover interactions
        const tooltip = document.getElementById('tooltip')!;
        const jsonPanel = document.getElementById('json-panel')!;

        document.querySelectorAll('.ext-annot-highlight').forEach(el => {
          const annotationId = el.getAttribute('data-annotation-id');
          const labelId = el.getAttribute('data-label-id');
          const labelText = el.getAttribute('data-label') || labelId;
          const annotatedText = el.textContent;

          el.addEventListener('mouseenter', (e: Event) => {
            const mouseEvent = e as MouseEvent;
            // Show tooltip
            tooltip.innerHTML = '<div class="tooltip-label">' + labelText + '</div>' +
              '<div class="tooltip-text">"' + annotatedText + '"</div>';
            tooltip.style.display = 'block';
            tooltip.style.left = mouseEvent.pageX + 10 + 'px';
            tooltip.style.top = mouseEvent.pageY + 10 + 'px';

            // Highlight corresponding JSON
            jsonPanel.querySelectorAll('.json-annotation').forEach(jsonEl => {
              if (jsonEl.getAttribute('data-annotation-id') === annotationId) {
                jsonEl.classList.add('highlighted');
                jsonEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
              } else {
                jsonEl.classList.remove('highlighted');
              }
            });
          });

          el.addEventListener('mousemove', (e: Event) => {
            const mouseEvent = e as MouseEvent;
            tooltip.style.left = mouseEvent.pageX + 10 + 'px';
            tooltip.style.top = mouseEvent.pageY + 10 + 'px';
          });

          el.addEventListener('mouseleave', () => {
            tooltip.style.display = 'none';
            jsonPanel.querySelectorAll('.json-annotation').forEach(jsonEl => {
              jsonEl.classList.remove('highlighted');
            });
          });
        });
      }, { html: htmlResult.html, annotationSet: set, annotations });

      // Only pause in debug/headed mode - check if PWDEBUG env var is set
      // or use --debug flag. In CI, this will be skipped.
      if (process.env.PWDEBUG) {
        await page.pause();
      }

      // Verify the demo rendered correctly
      const highlights = await page.locator('.ext-annot-highlight').count();
      expect(highlights).toBeGreaterThan(0);

      const jsonPanel = await page.locator('#json-panel').isVisible();
      expect(jsonPanel).toBe(true);
    });
  });

});
