import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const kind of ['tab', 'unmarked', 'text'] as const) {
  test(`paragraph continuation handles ${kind} inline boxes without cutting their content`, async ({ page }) => {
    const lines = Array.from({ length: 29 }, (_, i) => `Alder${i}<span ${kind === 'unmarked' ? '' : 'data-docx-tab="left"'}
      style="display:inline-block;width:35pt;height:0;border-bottom:1px dotted">${kind === 'text' ? 'atomic' : ''}</span> checked`).join('<br>');
    await page.setContent(`<style>p{margin:0;font:12pt/16pt Arial}</style><div id="staging"><section data-section-index="0"
      data-page-width="300" data-page-height="160" data-content-width="260" data-content-height="120"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20"><p>${lines}</p></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      return { pages: result.totalPages, paragraphs: document.querySelectorAll('#container p').length,
        text: document.getElementById('container')!.textContent,
        boxes: document.querySelectorAll('#container p span').length,
        widths: Array.from(document.querySelectorAll('#container p span'), e => e.getBoundingClientRect().width) };
    });
    expect(result.text?.match(/Alder\d+/g)).toEqual(Array.from({ length: 29 }, (_, i) => `Alder${i}`));
    expect(result.boxes).toBe(29);
    expect(result.widths.every(width => Math.abs(width - 35 * 4 / 3) < 0.1)).toBe(true);
    if (kind === 'tab') expect(result.paragraphs).toBeGreaterThan(1);
    else expect(result.paragraphs).toBe(1);
  });
}
