import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const content of [false, true]) {
  test(`table wrapper ${content ? 'with other content remains indivisible' : 'keeps leading bookmarks on its first fragment'}`, async ({ page }) => {
    await page.setContent(`<style>p{margin:0;font:12pt/16pt Arial}td{padding:0}table{border-collapse:collapse}</style>
      <div id="staging"><section data-section-index="0" data-page-width="300" data-page-height="160"
      data-content-width="260" data-content-height="120" data-margin-top="20" data-margin-bottom="20"
      data-margin-left="20" data-margin-right="20"><div><a id="BirchStart">${content ? 'Intro' : ''}</a><a id="BirchAlias"></a>
      <table data-anchor="tbl:body:birch">${Array.from({ length: 29 }, (_, i) => `<tr><td><p>Birch${i}</p></td></tr>`).join('')}</table>
      </div></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      return { text: document.getElementById('container')!.textContent, tables: document.querySelectorAll('#container table').length,
        bookmarks: ['BirchStart', 'BirchAlias'].map(id => document.querySelectorAll(`#container #${id}`).length),
        bookmarkOnFirst: document.querySelector('#container .page-content')!.contains(document.querySelector('#container #BirchStart')),
        canonicalTables: document.querySelectorAll('#container [data-anchor="tbl:body:birch"]').length,
        pages: result.totalPages };
    });
    expect(result.text?.match(/Birch\d+/g)).toEqual(Array.from({ length: 29 }, (_, i) => `Birch${i}`));
    expect(result.bookmarks).toEqual([1, 1]);
    expect(result.bookmarkOnFirst).toBe(true);
    expect(result.canonicalTables).toBe(1);
    if (content) expect(result.tables).toBe(1);
    else expect(result.tables).toBeGreaterThan(1);
  });
}
