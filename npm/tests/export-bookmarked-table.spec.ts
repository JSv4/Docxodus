import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { bookmarkedTableDocx } from './docx-bookmarked-table-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

test('a bookmarked table continues with all rows and one working bookmark', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(bookmarkedTableDocx(), { browser, outputs: ['html', 'pdf'],
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
  const geometry = await page.locator('.page-content table').evaluateAll(tables => tables.map(table => {
    const r = table.getBoundingClientRect(), band = table.closest('.page-content')!.getBoundingClientRect();
    return { overflow: r.bottom - band.bottom, rows: table.querySelectorAll('tr').length };
  }));
  expect(geometry.every(table => table.overflow <= 1)).toBe(true);
  expect(geometry.reduce((sum, table) => sum + table.rows, 0)).toBe(61);
  expect(result.pageMap.fragments.filter(fragment => fragment.anchorId.startsWith('tbl:'))).toHaveLength(geometry.length);
});
