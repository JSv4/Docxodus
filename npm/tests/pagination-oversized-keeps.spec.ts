import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import { fittingTilePng } from './docx-page-fitting-fixture.js';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
const entries = Array.from({ length: 17 }, (_, index) => `Saffron${String(index + 1).padStart(2, '0')}`);

for (const form of ['plain', 'bordered', 'image']) for (const widow of [false, true]) {
  test(`three wrapped lines, ${form}, widow ${widow}`, async ({ page }) => {
    const picture = form === 'image' ? `<img width="1" height="1" src="data:image/png;base64,${fittingTilePng().toString('base64')}">` : '';
    const paragraph = `<p data-case="target" data-widow-control="${widow}">aaaa${picture} bbbb cccc dddd eeee</p>`;
    await page.setContent(`<style>
      #pagination-staging, #pagination-container { font: 12pt/16pt monospace; }
      p { margin: 0; }
    </style><div id="pagination-staging"><div data-section-index="0"
      data-page-width="140" data-page-height="120" data-content-width="100" data-content-height="80"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      <p style="height:47pt">Opening.</p>
      ${form === 'bordered' ? `<div style="border-left:1px solid">${paragraph}</div>` : paragraph}
    </div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const parts = await page.evaluate(() => {
      const container = document.getElementById('pagination-container')!;
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('pagination-staging'),
        container, { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      return Array.from(container.querySelectorAll('[data-case="target"]'), p => ({
        page: p.closest<HTMLElement>('.page-box')!.dataset.pageNumber, text: p.textContent,
      }));
    });
    expect(parts.map(p => p.page)).toEqual(widow ? ['2'] : ['1', '2']);
    expect(parts.map(p => p.text).join('').replace(/\s+/g, ' ').trim()).toBe('aaaa bbbb cccc dddd eeee');
    await expect(page.locator('#pagination-container img')).toHaveCount(form === 'image' ? 1 : 0);
  });
}

for (const [name, preludeLines, widow, keep, expected] of [
  ['six and three', 10, true, false, [6, 3]],
  ['avoid an orphan', 15, true, false, [0, 9]],
  ['avoid a widow', 8, true, false, [7, 2]],
  ['widow off', 8, false, false, [8, 1]],
  ['explicit keep', 10, true, true, [0, 9]],
] as const) for (const hardBreaks of [false, true]) {
  test(`widow control: ${name}, ${hardBreaks ? 'hard breaks' : 'wrapped text'}`, async ({ page }) => {
    const labels = entries.slice(0, 9);
    await page.setContent(`<style>
      #pagination-staging, #pagination-container { font: 12pt/16pt monospace; }
      p { margin: 0; }
    </style><div id="pagination-staging"><div data-section-index="0"
      data-page-width="220" data-page-height="296" data-content-width="180" data-content-height="256"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      <p style="height:${preludeLines * 16}pt">Opening receipt.</p>
      <p data-case="target" data-widow-control="${widow}" data-keep-lines="${keep}">
        ${labels.map(label => `<b>${label}</b> received.`).join(hardBreaks ? '<br>' : ' ')}
      </p>
    </div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(() => {
      const container = document.getElementById('pagination-container')!;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      new PaginationEngine(document.getElementById('pagination-staging'), container,
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      return Array.from(container.querySelectorAll('.page-box'), page =>
        Array.from(page.querySelectorAll('[data-case="target"] b'), b => b.textContent));
    });
    expect(pages.map(p => p.length)).toEqual(expected);
    expect(pages.flat()).toEqual(labels);
  });
}

for (const keep of ['keep-lines', 'widow-control', 'keep-with-next'] as const) {
  test(`an oversized ${keep} paragraph continues without losing a line`, async ({ page }) => {
    await page.setContent(`<style>
      #pagination-staging, #pagination-container { font: 12pt/16pt Arial; }
      p { margin: 0; }
    </style>
    <div id="pagination-staging"><div data-section-index="0"
      data-page-width="260" data-page-height="120" data-content-width="220" data-content-height="80"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      <p id="transfer-register" data-anchor="p:body:register" data-${keep}="true"
        style="${keep === 'keep-with-next' ? 'break-after:avoid' : 'break-inside:avoid'}">${entries.map(entry => `<b>${entry}</b> transferred.`).join('<br>')}</p>
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

for (const guard of ['unowned avoid', 'forced break', 'editor', 'unresolved note']) {
  test(`oversized keep-next retains the ${guard} restriction`, async ({ page }) => {
    await page.setContent(`<style>#staging,#container{font:12pt/16pt Arial}p{margin:0}</style>
      <div id="staging"><div data-section-index="0" data-page-width="260" data-page-height="120"
        data-content-width="220" data-content-height="80" data-margin-top="20" data-margin-bottom="20"
        data-margin-left="20" data-margin-right="20">
        <p data-case="target" ${guard === 'unowned avoid' ? '' : 'data-keep-with-next="true"'}
          style="break-after:${guard === 'forced break' ? 'page' : 'avoid'}">
          ${entries.join('<br>')}${guard === 'unresolved note' ? '<sup data-footnote-id="missing">1</sup>' : ''}</p>
        <p>Closing receipt.</p>
      </div></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    await page.evaluate(guard => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('container'), { fragmentParagraphs: guard !== 'editor', showPageNumbers: false }).paginate();
    }, guard);
    await expect(page.locator('#container [data-case="target"]')).toHaveCount(1);
    expect(await page.locator('#container [data-case="target"]').textContent()).toContain(entries.at(-1));
  });
}
