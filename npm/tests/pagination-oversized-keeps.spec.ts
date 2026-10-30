import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
const entries = Array.from({ length: 17 }, (_, index) => `Saffron${String(index + 1).padStart(2, '0')}`);

for (const keep of ['keep-lines', 'widow-control'] as const) {
  test(`an oversized ${keep} paragraph continues without losing a line`, async ({ page }) => {
    await page.setContent(`<style>
      #pagination-staging, #pagination-container { font: 12pt/16pt Arial; }
      p { margin: 0; }
    </style>
    <div id="pagination-staging"><div data-section-index="0"
      data-page-width="260" data-page-height="120" data-content-width="220" data-content-height="80"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      <p id="transfer-register" data-anchor="p:body:register" data-${keep}="true"
        style="break-inside:avoid">${entries.map(entry => `<b>${entry}</b> transferred.`).join('<br>')}</p>
      <p>Closing receipt.</p>
    </div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const container = document.getElementById('pagination-container')!;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const result = new PaginationEngine(document.getElementById('pagination-staging'), container,
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      const parts = Array.from(container.querySelectorAll<HTMLElement>('p')).filter(p => p.querySelector('b'));
      return {
        pages: result.totalPages,
        entries: parts.flatMap(p => Array.from(p.querySelectorAll('b'), b => b.textContent)),
        lines: parts.map(p => p.querySelectorAll('b').length),
        ids: container.querySelectorAll('#transfer-register').length,
        anchors: container.querySelectorAll('[data-anchor="p:body:register"]').length,
        overflow: Array.from(container.querySelectorAll<HTMLElement>('.page-content')).map(body =>
          Math.max(0, ...Array.from(body.children, child => child.getBoundingClientRect().bottom
            - body.getBoundingClientRect().bottom))),
        closing: container.textContent?.match(/Closing receipt\./g)?.length,
      };
    });
    expect(result.pages).toBeGreaterThan(1);
    expect(result.entries).toEqual(entries);
    expect(result.lines.every(count => count >= 2)).toBe(true);
    expect(result.overflow.every(height => height <= 1)).toBe(true);
    expect(result.ids).toBe(1);
    expect(result.anchors).toBe(1);
    expect(result.closing).toBe(1);
  });
}

test('a keep that fits a fresh page moves there as one paragraph', async ({ page }) => {
  await page.setContent(`<style>
    #pagination-staging, #pagination-container { font: 12pt/16pt Arial; }
    p { margin: 0; }
  </style>
  <div id="pagination-staging"><div data-section-index="0"
    data-page-width="260" data-page-height="120" data-content-width="220" data-content-height="80"
    data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
    <p style="height:48pt">Opening receipt.</p>
    <p data-keep-lines="true" data-case="kept">First transfer.<br>Second transfer.<br>Third transfer.</p>
  </div></div><div id="pagination-container"></div>`);
  await page.addScriptTag({ path: bundle });
  const result = await page.evaluate(() => {
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const container = document.getElementById('pagination-container')!;
    const layout = new PaginationEngine(document.getElementById('pagination-staging'), container,
      { fragmentParagraphs: true, showPageNumbers: false }).paginate();
    const paragraphs = Array.from(container.querySelectorAll<HTMLElement>('[data-case="kept"]'));
    return { pages: layout.totalPages, count: paragraphs.length,
      page: paragraphs[0].closest<HTMLElement>('.page-box')?.dataset.pageNumber,
      text: paragraphs.map(p => p.textContent).join('') };
  });
  expect(result).toEqual({ pages: 2, count: 1, page: '2',
    text: 'First transfer.Second transfer.Third transfer.' });
});
