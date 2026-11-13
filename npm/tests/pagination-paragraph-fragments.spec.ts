import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const words = [
  'alpha', 'bravo', 'charlie', 'delta', 'echo', 'foxtrot', 'golf', 'hotel',
  'india', 'juliet', 'kilo', 'lima', 'mike', 'november', 'oscar', 'papa',
  'quebec', 'romeo', 'sierra', 'tango', 'uniform', 'victor', 'whiskey', 'xray',
  'yankee', 'zulu',
].join(' ');

function normalize(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

function viewerHtml(blocks: string): string {
  return `
    <style>
      #pagination-staging, #pagination-container { font: 12pt/12pt Arial; }
      #pagination-staging p { box-sizing: border-box; width: 100%; }
    </style>
    <div id="pagination-staging">
      <div data-section-index="0"
           data-page-width="122" data-page-height="56"
           data-content-width="120" data-content-height="54"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        ${blocks}
      </div>
    </div>
    <div id="pagination-container"></div>`;
}

async function addPaginationBundle(page: Page): Promise<void> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
}

test.describe('Pagination paragraph fragments', () => {
  for (const kind of ['owned-math', 'owned-structured-math', 'plain-marker', 'raw-math', 'ruby-math', 'frame-math', 'display-math', 'floating-math']) {
    test(`equation fragmentation preserves the ${kind} boundary`, async ({ page }) => {
      // Scripts, n-ary limits, fences and tables (#1705) are also whole converter-owned equations.
      const formula = kind === 'owned-structured-math'
        ? '<mrow><munderover><mo largeop="true">∑</mo><mi>k</mi><mi>n</mi></munderover><msup><mi>x</mi><mn>2</mn></msup>'
          + '<mrow><mo fence="true" form="prefix">(</mo><mtable><mtr><mtd><mi>a</mi></mtd></mtr></mtable><mo fence="true" form="postfix">)</mo></mrow></mrow>'
        : '<mfrac><mn>17</mn><mn>5</mn></mfrac>';
      const math = `<math xmlns="http://www.w3.org/1998/Math/MathML" display="${kind === 'display-math' ? 'block' : 'inline'}"
        ${kind === 'floating-math' ? 'style="float:left"' : ''}>${formula}</math>`;
      const owned = `<span data-docx-math="true">${math}</span>`;
      const text = Array.from({ length: 4 }, () => words).join(' ');
      const body = kind === 'owned-math' || kind === 'owned-structured-math' ? owned.repeat(100) : kind === 'plain-marker'
        ? `<span data-docx-math="true">${text}</span>` : (kind === 'ruby-math'
          ? `<span data-docx-ruby="true" style="display:inline-block"><ruby>${math}</ruby></span>`
          : kind === 'frame-math' ? `<span data-docx-frame-line>${math}</span>`
            : kind === 'raw-math' ? math : owned) + text;
      await page.setContent(viewerHtml(`<p style="margin:0">${body}</p>`));
      await addPaginationBundle(page);
      const result = await page.evaluate(() => {
        const source = document.querySelector('#pagination-staging p')!;
        const text = source.textContent, mathCount = source.querySelectorAll('math').length;
        const container = document.getElementById('pagination-container')!;
        new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('pagination-staging'), container,
          { showPageNumbers: false, fragmentParagraphs: true }).paginate();
        return { text, mathCount, parts: container.querySelectorAll('.page-content > p').length,
          actualText: Array.from(container.querySelectorAll('.page-content > p'), p => p.textContent).join(''),
          actualMathCount: container.querySelectorAll('math').length };
      });
      expect(result.actualText).toBe(result.text);
      expect(result.actualMathCount).toBe(result.mathCount);
      if (kind === 'owned-math' || kind === 'owned-structured-math' || kind === 'plain-marker') expect(result.parts).toBeGreaterThan(1);
      else expect(result.parts).toBe(1);
    });
  }

  for (const kind of ['plain-host', 'shifted-host', 'leading-drawing', 'later-drawing', 'wrapped-drawing',
    'aligned-drawing', 'empty-aligned-drawing', 'break-before-drawing', 'image-before-drawing']) {
    test(`range fragmentation preserves the ${kind} eligibility boundary`, async ({ page }) => {
      const drawing = `<span data-docx-drawing-anchor="true" data-docx-anchor-h-relative="column"
        data-docx-anchor-v-relative="paragraph" data-docx-anchor-wrap="${kind === 'wrapped-drawing' ? 'wrapSquare' : 'wrapNone'}"
        ${kind.includes('aligned-drawing') ? 'data-docx-anchor-v-align="bottom"' : ''}
        style="position:absolute;width:12pt;height:12pt;left:100pt;top:0;background:blue"><span>${kind === 'empty-aligned-drawing' ? '' : 'Box'}</span></span>`;
      const rows = Array.from({ length: 12 }, (_, i) => `<span>Row${i + 1}</span>`);
      if (kind.includes('drawing')) rows[kind === 'later-drawing' ? 5 : 0] =
        (kind === 'break-before-drawing' ? '<br>' : kind === 'image-before-drawing'
          ? '<img style="width:60pt;height:120pt" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7">' : '')
        + drawing + rows[kind === 'later-drawing' ? 5 : 0];
      await page.setContent(viewerHtml(`<p style="position:relative;margin:0;left:${kind === 'shifted-host' ? 1 : 0}px">${rows.join('<br>')}</p>`));
      await addPaginationBundle(page);
      const count = await page.evaluate(() => {
        const { PaginationEngine } = (window as any).DocxodusPagination;
        const container = document.getElementById('pagination-container')!;
        new PaginationEngine(document.getElementById('pagination-staging'), container,
          { showPageNumbers: false, fragmentParagraphs: true }).paginate();
        return container.querySelectorAll('.page-content > p').length;
      });
      if (['plain-host', 'leading-drawing', 'later-drawing', 'break-before-drawing', 'image-before-drawing'].includes(kind))
        expect(count).toBeGreaterThan(1);
      else expect(count).toBe(1);
    });
  }

  test('read-only paginateHtml fragments simple paragraphs and preserves identity only on the head', async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await addPaginationBundle(page);

    const result = await page.evaluate((html) => {
      const { paginateHtml } = (window as any).DocxodusPagination;
      const pagination = paginateHtml(html, document.getElementById('viewer'), {
        showPageNumbers: false,
      });
      const parts = Array.from(document.querySelectorAll<HTMLElement>(
        '#pagination-container .page-content > p[data-case="simple"]',
      ));

      return {
        totalPages: pagination.totalPages,
        text: parts.map((part) => part.textContent || '').join(' '),
        anchors: parts.map((part) => part.getAttribute('data-anchor')),
        ids: parts.map((part) => part.id),
        styles: parts.map((part) => {
          const style = getComputedStyle(part);
          return {
            marginTop: parseFloat(style.marginTop),
            marginBottom: parseFloat(style.marginBottom),
            textIndent: parseFloat(style.textIndent),
          };
        }),
      };
    }, viewerHtml(`
      <p id="source-paragraph" data-anchor="p-source" data-case="simple"
         style="margin: 9pt 0 11pt; text-indent: 18pt">${words}</p>`));

    expect(result.totalPages).toBeGreaterThan(1);
    expect(normalize(result.text)).toBe(words);
    expect(result.anchors).toEqual(['p-source', ...Array(result.anchors.length - 1).fill(null)]);
    expect(result.ids).toEqual(['source-paragraph', ...Array(result.ids.length - 1).fill('')]);

    // The source's leading spacing/indent stays with the head. Every synthetic
    // continuation starts as a normal line, and only the last part owns the
    // original trailing paragraph spacing.
    expect(result.styles[0].marginTop).toBeGreaterThan(0);
    expect(result.styles.slice(1).every((style) => style.marginTop === 0 && style.textIndent === 0)).toBe(true);
    expect(result.styles.slice(0, -1).every((style) => style.marginBottom === 0)).toBe(true);
    expect(result.styles.at(-1)!.marginBottom).toBeGreaterThan(0);
  });

  test('direct PaginationEngine callers remain opt-in', async ({ page }) => {
    await page.setContent(viewerHtml(`<p data-case="plain">${words}</p>`));
    await addPaginationBundle(page);

    const result = await page.evaluate(() => {
      const staging = document.getElementById('pagination-staging') as HTMLElement;
      const container = document.getElementById('pagination-container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const pagination = new PaginationEngine(staging, container, {
        showPageNumbers: false,
      }).paginate();
      return {
        totalPages: pagination.totalPages,
        paragraphParts: container.querySelectorAll('p[data-case="plain"]').length,
      };
    });

    expect(result.totalPages).toBe(1);
    expect(result.paragraphParts).toBe(1);
  });

  test('hard breaks work without a grapheme segmenter and retain inline identity and atomic fields', async ({ page }) => {
    const lines = Array.from({ length: 12 }, (_, index) =>
      `<span data-field="PAGE">${index + 1}</span><b>日本語</b><i>e\u0301👩‍💻</i>`);
    await page.setContent(viewerHtml(`<p data-case="unicode" style="margin:0"><a
      id="inline-link" data-anchor="inline-anchor" href="https://example.org/lines">${lines.join('<br>')}</a></p>`));
    await addPaginationBundle(page);
    const result = await page.evaluate(() => {
      Object.defineProperty(Intl, 'Segmenter', { value: undefined, configurable: true });
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const container = document.getElementById('pagination-container')!;
      const pagination = new PaginationEngine(document.getElementById('pagination-staging'), container,
        { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      const parts = Array.from(container.querySelectorAll<HTMLElement>('p[data-case="unicode"]'));
      return {
        pages: pagination.totalPages,
        text: parts.map(p => p.textContent).join(''),
        fields: parts.flatMap(p => Array.from(p.querySelectorAll('[data-field]'), f => f.textContent)),
        links: parts.map(p => p.querySelector('a')?.getAttribute('href')),
        ids: container.querySelectorAll('#inline-link').length,
        anchors: container.querySelectorAll('[data-anchor="inline-anchor"]').length,
        heights: parts.map(p => p.getBoundingClientRect().height),
      };
    });
    expect(result.pages).toBeGreaterThanOrEqual(3);
    expect(result.text).toBe(Array.from({ length: 12 }, (_, i) => `${i + 1}日本語e\u0301👩‍💻`).join(''));
    expect(result.fields).toEqual(Array.from({ length: 12 }, (_, i) => String(i + 1)));
    expect(result.links.every(link => link === 'https://example.org/lines')).toBe(true);
    expect(result.ids).toBe(1);
    expect(result.anchors).toBe(1);
    expect(result.heights.filter(height => height > 72)).toEqual([]);
  });

  test('ordinary prose continuations keep fragmenting after the second page', async ({ page }) => {
    const text = Array(6).fill(words).join(' ');
    await page.setContent(viewerHtml(`<p data-case="prose" style="margin:0">${text}</p>`));
    await addPaginationBundle(page);
    const result = await page.evaluate(() => {
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const container = document.getElementById('pagination-container')!;
      const pagination = new PaginationEngine(document.getElementById('pagination-staging'), container,
        { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      const parts = Array.from(container.querySelectorAll<HTMLElement>('p[data-case="prose"]'));
      return { pages: pagination.totalPages, text: parts.map(p => p.textContent).join(''),
        heights: parts.map(p => p.getBoundingClientRect().height) };
    });
    expect(result.pages).toBeGreaterThan(2);
    expect(normalize(result.text)).toBe(text);
    expect(result.heights.every(height => height <= 72)).toBe(true);
  });

  test('relaxes impossible line keeps while guarding pagination-sensitive descendants', async ({ page }) => {
    await page.setContent(viewerHtml(`
      <p data-case="keep-lines" data-keep-lines="true">${words}</p>
      <p data-case="unresolved-note" data-keep-lines="true">${words}<sup data-footnote-id="missing">9</sup></p>
      <p data-case="keep-next" data-keep-with-next="true">${words}</p>
      <p data-case="clearing-break">alpha<br style="clear:left">bravo<br>charlie<br>delta<br>echo<br>foxtrot</p>
      <p data-case="page-break">alpha<br style="break-before:page">bravo<br>charlie<br>delta<br>echo<br>foxtrot</p>
      <p data-case="line-break" data-keep-lines="true">alpha<br>bravo<br>charlie<br>delta<br>echo<br>foxtrot</p>`));
    await addPaginationBundle(page);

    const result = await page.evaluate(() => {
      const staging = document.getElementById('pagination-staging') as HTMLElement;
      const container = document.getElementById('pagination-container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      new PaginationEngine(staging, container, {
        showPageNumbers: false,
        fragmentParagraphs: true,
      }).paginate();

      const count = (caseName: string) =>
        container.querySelectorAll(`p[data-case="${caseName}"]`).length;
      return {
        keepLines: count('keep-lines'),
        unresolvedNote: count('unresolved-note'),
        keepNext: count('keep-next'),
        lineBreak: count('line-break'),
        clearingBreak: count('clearing-break'),
        pageBreak: count('page-break'),
      };
    });

    expect(result).toEqual({
      keepLines: 3,
      unresolvedNote: 1,
      keepNext: 2,
      lineBreak: 2,
      clearingBreak: 1,
      pageBreak: 1,
    });
  });
});
