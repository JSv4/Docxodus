import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
const entries = Array.from({ length: 41 }, (_, index) => `Indigo${String(index + 1).padStart(2, '0')}`);

for (const shape of ['paragraph', 'nested-table'] as const) {
  test(`a body cell's long ${shape} continues across pages without repeating its contents`, async ({ page }) => {
    const contents = shape === 'paragraph'
      ? `<p id="inventory" data-anchor="p:body:inventory">${entries.map(entry => `${entry} logged.`).join('<br>')}</p>`
      : `<div><table style="width:100%"><colgroup><col style="width:100%"></colgroup><tbody>
        ${entries.map((entry, index) => `<tr><td><p id="item-${index}" data-anchor="p:body:item-${index}">${entry} logged.</p></td></tr>`).join('')}
      </tbody></table></div>`;
    await page.setContent(`<style>
      #pagination-staging, #pagination-container { font: 12pt/16pt Arial; }
      table { border-collapse: collapse; table-layout: fixed; }
      td { vertical-align: top; padding: 2pt; border: 1pt solid black; }
      p { margin: 0; }
    </style>
    <div id="pagination-staging"><div data-section-index="0"
      data-page-width="340" data-page-height="180" data-content-width="300" data-content-height="140"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      <div><table id="outer-table" data-anchor="tbl:body:inventory" style="width:300pt">
        <colgroup><col style="width:65%"><col style="width:35%"></colgroup><tbody><tr>
        <td id="inventory-cell" data-anchor="tc:body:inventory">${contents}</td>
        <td id="receipt-cell" data-anchor="tc:body:receipt"><p>Receipt confirmed.</p></td>
      </tr></tbody></table></div><p>Inventory complete.</p>
    </div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const container = document.getElementById('pagination-container')!;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const layout = new PaginationEngine(document.getElementById('pagination-staging'), container,
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      const bodies = Array.from(container.querySelectorAll<HTMLElement>('.page-content'));
      const ids = Array.from(container.querySelectorAll('[id]'), element => element.id);
      return { pages: layout.totalPages, text: bodies.map(body => body.textContent).join(' '),
        overflow: bodies.map(body => Math.max(0, ...Array.from(body.querySelectorAll('p'), child =>
          child.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom))),
        uniqueIds: new Set(ids).size === ids.length };
    });
    expect(result.pages).toBeGreaterThan(1);
    expect(result.overflow.every(height => height <= 1)).toBe(true);
    expect(result.text.match(/Indigo\d+/g)).toEqual(entries);
    expect(result.text.match(/Receipt confirmed\./g)).toHaveLength(1);
    expect(result.text.match(/Inventory complete\./g)).toHaveLength(1);
    expect(result.uniqueIds).toBe(true);
  });
}
