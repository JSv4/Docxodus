import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// semanticLists is ConvertDocxToHtmlComplete's last argument (after revisionPresentation).
// Parsing the output with the browser's HTML parser checks the tree a consumer actually gets.
test.describe('ConvertDocxToHtmlComplete semanticLists (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('emits ol/li for list paragraphs only when asked', async ({ page }) => {
    // Nine list paragraphs across three w:num instances.
    const bytes = readTestFile('DB012-Lists-With-Different-Numberings.docx');

    const result = await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const dc = (window as any).Docxodus.DocumentConverter;
      const render = (...semanticLists: boolean[]): string => dc.ConvertDocxToHtmlComplete(
        bin, 'Document', 'docx-', true, '', -1, 'comment-', 0, 1.0, 'page-',
        false, 0, 'annot-', true, false, false, true, true, false, null, /*stampAnchors*/ true,
        /*revisionPresentation*/ 0, ...semanticLists);
      const parse = (html: string) => new DOMParser().parseFromString(html, 'text/html');

      const on = parse(render(true));
      const items = Array.from(on.querySelectorAll('li'));
      return {
        omitted: render(),
        off: render(false),
        offItems: parse(render(false)).querySelectorAll('li').length,
        lists: on.querySelectorAll('ol').length,
        items: items.length,
        itemsOutsideLists: items.filter((li) => !/^(OL|UL)$/.test(li.parentElement?.tagName ?? '')).length,
        anchoredItems: items.filter((li) => li.hasAttribute('data-anchor')).length,
      };
    }, Array.from(bytes));

    expect(result.omitted).toBe(result.off);
    expect(result.offItems).toBe(0);
    expect(result.lists).toBe(3);
    expect(result.items).toBe(9);
    expect(result.itemsOutsideLists).toBe(0);
    expect(result.anchoredItems).toBe(9);
  });
});
