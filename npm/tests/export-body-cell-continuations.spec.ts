import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { longBodyCellDocx } from './docx-long-body-cell-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

// Word splits a row taller than a page whatever its cells' vertical alignment, and aligns a short
// cell within the part of the row on the row's first page (issue #1558).
for (const [shape, vAlign] of [['paragraph', 'top'], ['nested-table', 'top'],
  ['paragraph', 'center'], ['paragraph', 'bottom']] as const) {
  test(`a ${vAlign}-aligned body cell's ${shape} preserves its full inventory across PDF pages`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(longBodyCellDocx(shape, 73, vAlign), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
      commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const pages: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(pages.join(' ').match(/Umber\d+/g)).toEqual(Array.from({ length: 73 }, (_, index) =>
        `Umber${String(index + 1).padStart(3, '0')}`));
      expect(pages.join(' ').match(/Receipt confirmed\./g)).toHaveLength(1);
      expect(pages.join(' ').match(/Inventory complete\./g)).toHaveLength(1);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const overflow = await page.locator('.page-content').evaluateAll(bodies => bodies.flatMap(body =>
      Array.from(body.querySelectorAll('p'), p => p.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom)));
    expect(overflow.every(height => height <= 1)).toBe(true);
    const receipt = await page.locator('.page-box').first().locator('p', { hasText: 'Receipt confirmed.' })
      .evaluate(p => {
        const row = p.closest('tr')!.getBoundingClientRect(), line = p.getBoundingClientRect();
        return { above: line.top - row.top, below: row.bottom - line.bottom };
      });
    const offset = { top: receipt.above, center: (receipt.above - receipt.below) / 2, bottom: receipt.below }[vAlign];
    expect(Math.abs(offset)).toBeLessThanOrEqual(3);
  });
}
