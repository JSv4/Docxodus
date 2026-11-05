import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const factor of [1, 1.05, 2]) {
  test(`table columns retain their proportions at source width factor ${factor}`, async ({ page }) => {
    const widths = [70, 120, 90].map(width => width * factor);
    await page.setContent(`<style>body{font:11pt/15pt Arial}p{margin:0}table{border-collapse:collapse;table-layout:fixed}
      td{padding:0;vertical-align:top}</style><div id="staging"><section data-section-index="0"
      data-page-width="320" data-page-height="200" data-content-width="280" data-content-height="160"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20">
      <div><table style="width:${280 * factor}pt"><colgroup>${widths.map(width => `<col style="width:${width}pt">`).join('')}</colgroup>
      <tbody>${Array.from({ length: 16 }, (_, row) => `<tr>${widths.map((width, column) =>
        `<td style="width:${width}pt"><p>Slate${row}${String.fromCharCode(65 + column)}</p></td>`).join('')}</tr>`).join('')}</tbody></table></div>
      </section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#container .page-content')).map(body => ({
        text: body.textContent,
        widths: Array.from(body.querySelectorAll('tr:first-child td'), cell => cell.getBoundingClientRect().width * 0.75),
        overflow: Array.from(body.querySelectorAll('table'), table => table.getBoundingClientRect().right
          - body.closest('.page-box')!.getBoundingClientRect().right),
      }));
    });
    expect(result.length).toBeGreaterThan(1);
    expect(result.map(page => page.text).join(' ').match(/Slate\d+[A-C]/g)).toEqual(Array.from({ length: 16 }, (_, row) =>
      ['A', 'B', 'C'].map(column => `Slate${row}${column}`)).flat());
    for (const body of result) {
      expect(body.overflow.every(overflow => overflow <= 1)).toBe(true);
      const expectedFactor = factor === 2 ? 1 : factor;
      for (let column = 0; column < 3; column++) expect(body.widths[column]).toBeCloseTo([70, 120, 90][column] * expectedFactor, 0);
    }
  });
}
