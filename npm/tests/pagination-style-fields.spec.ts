import { test, expect, Page } from '@playwright/test';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const directory = path.dirname(fileURLToPath(import.meta.url));
const heading = (id: string, text: string, pageBreak = false) =>
  `<p data-anchor="${id}" data-paragraph-style="Chapter" data-styleref-text="${text}"` +
  `${pageBreak ? ' data-page-break-before="true"' : ''}>${text}</p>`;
const marker = (last = false) => `<span data-field="STYLEREF" data-field-style="Chapter"` +
  `${last ? ' data-field-last="true"' : ''}>ALDER SECTION</span>`;

function documentHtml(body: string, header = marker(), footer = marker(), height = 240): string {
  return `<style>
    #pagination-staging, #pagination-container { font:12pt/12pt 'Liberation Serif'; }
    p { margin:0; } #pagination-hf-registry { display:none; }
  </style><div id="pagination-staging">
    <div id="pagination-hf-registry">
      <div data-section="0" data-hf-type="header-default"><p>Running ${header}</p></div>
      <div data-section="0" data-hf-type="footer-default"><p>Closing ${footer}</p></div>
    </div>
    <div data-section-index="0" data-page-width="228" data-page-height="${height}"
      data-content-width="180" data-content-height="${height - 48}"
      data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24"
      data-header-height="5" data-footer-height="5">${body}</div>
  </div><div id="pagination-container"></div>`;
}

async function paginate(page: Page, html: string) {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(directory, '../dist/pagination.bundle.js') });
  return page.evaluate(() => {
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const engine = new PaginationEngine(document.getElementById('pagination-staging'),
      document.getElementById('pagination-container'), { showPageNumbers: false, fragmentParagraphs: true });
    engine.paginate();
    return Array.from(document.querySelectorAll<HTMLElement>('#pagination-container .page-box')).map(box => {
      const header = box.querySelector<HTMLElement>('.page-header')!;
      const body = box.querySelector<HTMLElement>('.page-content')!;
      const footer = box.querySelector<HTMLElement>('.page-footer')!;
      return {
        header: header.textContent?.trim(), footer: footer.textContent?.trim(), body: body.textContent?.trim(),
        headerBottom: header.getBoundingClientRect().bottom, bodyTop: body.getBoundingClientRect().top,
        bodyBottom: body.getBoundingClientRect().bottom, footerTop: footer.getBoundingClientRect().top,
        headerHeight: header.getBoundingClientRect().height,
      };
    });
  });
}

test('running styles follow each physical page instead of a stale shared cache', async ({ page }) => {
  const pages = await paginate(page, documentHtml(
    heading('alder', 'ALDER SECTION') + '<p>First page body.</p>' +
    heading('birch', 'BIRCH SECTION', true) + '<p>Second page body.</p>'));
  expect(pages.map(p => p.header)).toEqual(['Running ALDER SECTION', 'Running BIRCH SECTION']);
  expect(pages.map(p => p.footer)).toEqual(['Closing ALDER SECTION', 'Closing BIRCH SECTION']);
});

test('first and last on-page styles are distinct, with nearest preceding fallback', async ({ page }) => {
  const pages = await paginate(page, documentHtml(
    heading('alder', 'ALDER SECTION') + heading('birch', 'BIRCH SECTION') +
    '<p data-page-break-before="true">No new heading.</p>', marker(), marker(true)));
  expect(pages.map(p => p.header)).toEqual(['Running ALDER SECTION', 'Running BIRCH SECTION']);
  expect(pages.map(p => p.footer)).toEqual(['Closing BIRCH SECTION', 'Closing BIRCH SECTION']);
});

test('a page before the first matching heading uses the forward match', async ({ page }) => {
  const pages = await paginate(page, documentHtml('<p>Opening body without a heading.</p>' +
    heading('birch', 'BIRCH SECTION', true)));
  expect(pages.map(p => p.header)).toEqual(['Running BIRCH SECTION', 'Running BIRCH SECTION']);
});

test('a long resolved running heading is measured before body allocation', async ({ page }) => {
  const title = 'Juniper calibration record with several descriptive words for the northern research station';
  const pages = await paginate(page, documentHtml(heading('alder', 'ALDER SECTION') +
    heading('long-heading', title, true) + '<p>Closing body remains visible.</p>'));
  expect(pages).toHaveLength(2);
  expect(pages[1].header).toBe('Running ' + title);
  expect(pages[1].headerHeight).toBeGreaterThan(pages[0].headerHeight + 20);
  for (const p of pages) {
    expect(p.headerBottom).toBeLessThanOrEqual(p.bodyTop + 0.1);
    expect(p.bodyBottom).toBeLessThanOrEqual(p.footerTop + 0.1);
  }
  expect(pages[1].body).toContain('Closing body remains visible.');
});

test('locked or literal running text without a dynamic marker remains unchanged', async ({ page }) => {
  const pages = await paginate(page, documentHtml(heading('alder', 'ALDER SECTION') +
    heading('birch', 'BIRCH SECTION', true), 'LOCKED REGISTER', 'LOCKED REGISTER'));
  expect(pages.map(p => p.header)).toEqual(['Running LOCKED REGISTER', 'Running LOCKED REGISTER']);
});

test('a heading that oscillates across its own running band fails explicitly', async ({ page }) => {
  const title = 'Juniper calibration record with several descriptive words for the northern research station';
  const html = documentHtml(heading('alder', 'ALDER SECTION') + '<p style="height:80pt">Filler.</p>' +
    heading('moving-heading', title).replace('<p ', '<p data-keep-lines="true" '), marker(true), 'Fixed footer', 180);
  const outcome = await paginate(page, html).then(() => 'unexpected success', error => String(error));
  expect(outcome).toMatch(/running.*field.*(?:stabili|cycle)/i);
});

test('first/even/default variants and page fields follow a section numbering restart', async ({ page }) => {
  const number = '<span data-field="PAGE" data-field-format="ROMAN">I</span>';
  const total = '<span data-field="NUMPAGES">1</span>';
  const html = documentHtml(heading('alder', 'ALDER SECTION') + heading('birch', 'BIRCH SECTION', true)
    + heading('cedar', 'CEDAR SECTION', true), `Default ${marker()} ${number}/${total}`)
    .replace('data-section-index="0"', 'data-section-index="0" data-page-num-start="2"')
    .replace('<div data-section="0" data-hf-type="footer-default">',
      `<div data-section="0" data-hf-type="header-first"><p>First ${marker()} ${number}/${total}</p></div>`
      + `<div data-section="0" data-hf-type="header-even"><p>Even ${marker()} ${number}/${total}</p></div>`
      + '<div data-section="0" data-hf-type="footer-default">');
  const pages = await paginate(page, html);
  expect(pages.map(p => p.header)).toEqual([
    'First ALDER SECTION II/3', 'Running Default BIRCH SECTION III/3', 'Even CEDAR SECTION IV/3',
  ]);
});

test('split headings retain their full source text on every page', async ({ page }) => {
  const title = Array.from({ length: 90 }, (_, i) => `Juniper${i}`).join(' ');
  const pages = await paginate(page, documentHtml(heading('long', title),
    `<span style="font-size:3pt;line-height:3pt">${marker()}</span>`, 'Fixed footer'));
  expect(pages.length).toBeGreaterThan(1);
  expect(pages.every(p => p.header === 'Running ' + title)).toBe(true);
  expect(pages.map(p => p.body).join(' ').replace(/\s+/g, ' ').trim()).toBe(title);
});

test('textbox and note headings cannot supply the main-story running style', async ({ page }) => {
  const html = documentHtml('<div data-docx-textbox="true">' + heading('textbox', 'WRONG BOX') + '</div>'
    + heading('birch', 'BIRCH SECTION'))
    .replace('<div id="pagination-hf-registry">',
      '<div id="pagination-footnote-registry" style="display:none">' + heading('note', 'WRONG NOTE')
      + '</div><div id="pagination-hf-registry">');
  const pages = await paginate(page, html);
  expect(pages[0].header).toBe('Running BIRCH SECTION');
});

test('a later tall running heading reflows balanced columns into its actual body band', async ({ page }) => {
  const title = 'Juniper calibration record with several descriptive words for the northern research station';
  let body = '';
  for (let i = 0; i < 4; i++) {
    if (i > 0) body += '<div data-page-break="true"></div>';
    body += heading(`chapter-${i}`, i === 3 ? title : 'ALDER SECTION');
    for (let j = 0; j < (i === 3 ? 23 : 30); j++) body += `<p style="height:12pt">Row${i}-${j}</p>`;
  }
  const html = documentHtml(body, marker(), 'Fixed footer')
    .replace('data-section-index="0"', 'data-section-index="0" data-cols="2" data-col-gap="12"');
  const pages = await paginate(page, html);
  expect(pages.length).toBeGreaterThanOrEqual(4);
  expect(pages[3].header).toBe('Running ' + title);
  expect(pages.every(p => p.body!.length > 0)).toBe(true);
  const rows = pages.map(p => p.body).join('').match(/Row\d+-\d+/g);
  const expected = Array.from({ length: 4 }, (_, i) =>
    Array.from({ length: i === 3 ? 23 : 30 }, (_, j) => `Row${i}-${j}`)).flat();
  expect(rows).toEqual(expected);
});

test('an oversized inline picture fits the actual fresh-page band below a long running heading', async ({ page }) => {
  const title = 'Juniper calibration record with several descriptive words for the northern research station';
  const picture = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="60" height="240"><rect width="60" height="240" fill="teal"/></svg>');
  const body = Array.from({ length: 4 }, (_, i) => heading(`chapter-${i}`, i === 3 ? title : 'ALDER SECTION', i > 0)).join('')
    + `<p><img src="${picture}" style="width:45pt;height:180pt"></p>`;
  const pages = await paginate(page, documentHtml(body));
  expect(pages).toHaveLength(5);
  expect(pages[4].header).toBe('Running ' + title);
  const geometry = await page.evaluate(() => {
    const image = document.querySelector<HTMLElement>('#pagination-container img')!;
    const body = image.closest<HTMLElement>('.page-content')!.getBoundingClientRect();
    const box = image.getBoundingClientRect();
    return { top: box.top, bottom: box.bottom, height: box.height, bodyTop: body.top, bodyBottom: body.bottom };
  });
  expect(geometry.height).toBeLessThan(180 * 4 / 3);
  expect(geometry.top).toBeGreaterThanOrEqual(geometry.bodyTop - 0.1);
  expect(geometry.bottom).toBeLessThanOrEqual(geometry.bodyBottom + 0.1);
});
