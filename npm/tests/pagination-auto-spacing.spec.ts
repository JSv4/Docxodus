import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const spacing of [0.9, 1.4]) {
  test(`automatic baseline placement at ${spacing} preserves formatted Unicode lines`, async ({ page }) => {
    const lines = Array.from({ length: 41 }, (_, index) => `<span>Larch${String(index + 1).padStart(2, '0')} A</span><span class="emphasis">̊ verified.</span>`).join('<br>');
    await page.setContent(`<style>body{font:12pt Arial}p{margin:0}p>span{
      position:relative;line-height:calc(1lh * var(--docx-auto-line-spacing));
      top:min(0px,calc(1lh * (1 / var(--docx-auto-line-spacing) - 1) / 2))} .emphasis{font-weight:bold}</style>
      <div id="staging"><section data-section-index="0" data-page-width="300" data-page-height="160"
      data-content-width="260" data-content-height="120" data-margin-top="20" data-margin-bottom="20"
      data-margin-left="20" data-margin-right="20"><p id="inventory" data-anchor="p:body:inventory"
      style="--docx-auto-line-spacing:${spacing}">${lines}</p></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const actual = await page.evaluate(() => {
      const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      const paragraphs = Array.from(document.querySelectorAll<HTMLElement>('#container .page-content p'));
      return { pages: result.totalPages, text: paragraphs.map(p => p.textContent).join(' '),
        overflow: paragraphs.map(p => p.getBoundingClientRect().bottom - p.closest('.page-content')!.getBoundingClientRect().bottom),
        fragments: paragraphs.map(p => p.textContent),
        bold: paragraphs.flatMap(p => Array.from(p.querySelectorAll('.emphasis'), span => getComputedStyle(span).fontWeight)),
        sourceIdentities: document.querySelectorAll('#container [data-anchor="p:body:inventory"]').length };
    });
    expect(actual.pages).toBeGreaterThan(1);
    expect(actual.text.match(/Larch\d+/g)).toEqual(Array.from({ length: 41 }, (_, index) => `Larch${String(index + 1).padStart(2, '0')}`));
    expect(actual.text.match(/Å/g)).toHaveLength(41);
    expect(actual.fragments.every(text => !text?.startsWith('̊'))).toBe(true);
    expect(actual.overflow.every(overflow => overflow <= 1)).toBe(true);
    expect(actual.bold.every(weight => Number(weight) >= 700)).toBe(true);
    expect(actual.sourceIdentities).toBe(1);
  });
}

for (const layout of ['unmarked', 'horizontal'] as const) {
  test(`an unsupported ${layout} relative layout remains indivisible`, async ({ page }) => {
    await page.setContent(`<style>body{font:12pt/16pt Arial}p{margin:0}</style><div id="staging">
      <section data-section-index="0" data-page-width="300" data-page-height="120"
      data-content-width="260" data-content-height="80" data-margin-top="20" data-margin-bottom="20"
      data-margin-left="20" data-margin-right="20"><p ${layout === 'horizontal' ? 'style="--docx-auto-line-spacing:1.4"' : ''}>
      <span style="position:relative;top:-1px;${layout === 'horizontal' ? 'left:4pt' : ''}">
      ${Array.from({ length: 15 }, (_, i) => `Guard${i}`).join('<br>')}</span></p></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    await page.evaluate(() => new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
      { fragmentParagraphs: true, showPageNumbers: false }).paginate());
    await expect(page.locator('#container .page-content p')).toHaveCount(1);
  });
}
