import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

// Issue #961: the browser's typed compare exports (DocumentComparer) route through the shared
// DocxDiffOps facade. Before, they built their own settings: author "Docxodus" where .NET said
// "Open-Xml-PowerTools", DateTime.UtcNow dates, and stack traces in error JSON.

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

const read = (name: string) => Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, name)));
const ORIGINAL = read('WC/WC001-Digits.docx');
const MODIFIED = read('WC/WC001-Digits-Mod.docx');

async function ready(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('Browser compare goes through the shared facade (#961)', () => {
  test.beforeEach(async ({ page }) => ready(page));

  test('typed compare exports are byte-identical to the DocxDiff facade and repeatable', async ({ page }) => {
    const result = await page.evaluate(([a, b]) => {
      const D = (window as any).Docxodus;
      const left = () => new Uint8Array(a);
      const right = () => new Uint8Array(b);
      const hex = (bytes: Uint8Array) => Array.from(bytes).join(',');
      // The front door is DocxDiff with the inputs' own revisions pre-accepted.
      const facade = (json: string) => hex(D.DocxDiffBridge.Compare(left(), right(), json));
      return {
        noAuthor: hex(D.DocumentComparer.CompareDocuments(left(), right(), null)),
        noAuthorAgain: hex(D.DocumentComparer.CompareDocuments(left(), right(), null)),
        facadeDefault: facade('{"preAcceptInputRevisions":true}'),
        named: hex(D.DocumentComparer.CompareDocumentsWithOptions(left(), right(), 'Reviewer', true)),
        facadeNamed: facade('{"preAcceptInputRevisions":true,"authorForRevisions":"Reviewer","caseInsensitive":true}'),
      };
    }, [ORIGINAL, MODIFIED] as const);

    expect(result.noAuthor.length).toBeGreaterThan(1000);
    expect(result.noAuthorAgain).toBe(result.noAuthor);
    expect(result.noAuthor).toBe(result.facadeDefault);
    expect(result.named).toBe(result.facadeNamed);
  });

  test('a compare naming no author stamps the core default', async ({ page }) => {
    const authors = await page.evaluate(([a, b]) => {
      const D = (window as any).Docxodus;
      const redline = D.DocumentComparer.CompareDocuments(new Uint8Array(a), new Uint8Array(b), null);
      const parsed = JSON.parse(D.DocumentComparer.GetRevisionsJson(redline));
      const list = Array.isArray(parsed) ? parsed : parsed.revisions;
      return Array.from(new Set(list.map((r: any) => r.author ?? r.Author)));
    }, [ORIGINAL, MODIFIED] as const);

    expect(authors).toEqual(['Docxodus']);
  });

  test('compare-to-HTML is deterministic and outlines the comparison author', async ({ page }) => {
    const result = await page.evaluate(([a, b]) => {
      const D = (window as any).Docxodus;
      const html = () => D.DocumentComparer.CompareDocumentsToHtmlWithOptions(new Uint8Array(a), new Uint8Array(b), null, true);
      return { first: html() as string, second: html() as string };
    }, [ORIGINAL, MODIFIED] as const);

    expect(result.second).toBe(result.first);
    expect(result.first).toContain('[data-author="Docxodus"]');
    expect(result.first).toContain('<ins');
  });

  test('error JSON carries no stack trace', async ({ page }) => {
    const error = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      return JSON.parse(D.DocumentComparer.CompareDocumentsToHtml(new Uint8Array(0), new Uint8Array(0), null));
    });

    expect(error.Error).toBeTruthy();
    expect(error).not.toHaveProperty('StackTrace');
  });
});
