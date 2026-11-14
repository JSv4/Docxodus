import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { bookmarkedTableDocx } from './docx-bookmarked-table-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const ranges of [false, true]) {
  test(`a bookmarked table ${ranges ? 'with tracked ranges between its cells ' : ''}continues with all rows and one working bookmark`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(bookmarkedTableDocx(ranges), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
      commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, pages: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(pages.join(' ').match(/Birch\d+[AB]/g)).toEqual(Array.from({ length: 61 }, (_, i) =>
        ['A', 'B'].map(column => `Birch${String(i + 1).padStart(2, '0')}${column}`)).flat());
      expect(pages.join(' ').match(/Birch ledger complete\./g)).toHaveLength(1);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await expect(page.locator('.page-content #BirchLedger')).toHaveCount(1);
    // Each range's two boundary marks stay in the cell after them (issue #1737).
    await expect(page.locator('.page-content .rev-cxml-marker')).toHaveCount(ranges ? 122 : 0);
    await expect(page.locator('.page-content td .rev-cxml-marker')).toHaveCount(ranges ? 122 : 0);
    const geometry = await page.locator('.page-content table').evaluateAll(tables => tables.map(table => {
      const r = table.getBoundingClientRect(), band = table.closest('.page-content')!.getBoundingClientRect();
      return { overflow: r.bottom - band.bottom,
        rows: Array.from(table.querySelectorAll('tr'), row => row.getAttribute('data-source-anchor-id')) };
    }));
    expect(geometry.length).toBeGreaterThan(1);
    expect(geometry.every(table => table.overflow <= 1)).toBe(true);
    // A row may break across pages (it has no w:cantSplit), so count the source rows the fragments carry.
    expect(new Set(geometry.flatMap(table => table.rows)).size).toBe(61);
    expect(result.pageMap.fragments.filter(fragment => fragment.anchorId.startsWith('tbl:'))).toHaveLength(geometry.length);
  });
}
