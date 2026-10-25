import { test, expect, type Page, type Browser, type TestInfo } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { generateNoteTableDocx } from './docx-note-table-fixture.js';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { inspectPdf } from './visual-parity/pdf.js';

/**
 * Issue #1178: a footnote holding a table. Word reference PDFs split the table between pages at row
 * boundaries, as a body table splits, and keep the citing line visible above the note:
 * - line 1 cites an 80-row table: page 1 holds the citing line and rows 1-51, page 2 rows 52-80
 *   and the body that follows;
 * - line 30 cites a 30-row table: page 1 holds 30 body lines and rows 1-18, page 2 rows 19-30.
 * The table was indivisible: it was clipped beside its citation in a note band that, at the full
 * band, also covered the body above it.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

interface Page_ {
  /** Body lines on the page, by number. */
  body: number[];
  /** The text of the page's first note table row. */
  firstRow: string;
  /** Note table rows on the page, by number. */
  rows: number[];
  /** Whether each body line on the page is inside the page's visible body area. */
  bodyVisible: boolean;
  notesVisible: boolean;
  sideRows: number[];
  uniqueAnchors: boolean;
}

async function paginate(page: Page, bytes: Uint8Array, fragmentParagraphs: boolean): Promise<Page_[]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ input, fragmentParagraphs }) => {
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      /* footnotes */ true, false, false, false, false, false, null, true,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    new (window as any).DocxodusPagination.PaginationEngine(
      host.querySelector('#pagination-staging'), host.querySelector('#pagination-container'),
      { showPageNumbers: false, fragmentParagraphs },
    ).paginate();
    const anchors = Array.from(host.querySelectorAll('.page-footnotes [data-anchor]'), el => el.getAttribute('data-anchor'));
    const numbers = (text: string, pattern: RegExp) => Array.from(text.matchAll(pattern), (match) => Number(match[1]));
    return Array.from(host.querySelectorAll<HTMLElement>('.page-box')).map((box) => {
      const content = box.querySelector<HTMLElement>('.page-content')!;
      const area = content.getBoundingClientRect();
      const lines = Array.from(content.querySelectorAll<HTMLElement>('p'));
      const noteArea = box.querySelector<HTMLElement>('.page-footnotes')?.getBoundingClientRect();
      return {
        body: numbers(content.textContent ?? '', /Body line (\d+)/g),
        rows: numbers(box.querySelector('.page-footnotes')?.textContent ?? '', /Note table row (\d+)/g),
        sideRows: numbers(box.querySelector('.page-footnotes')?.textContent ?? '', /Side table row (\d+)/g),
        uniqueAnchors: anchors.length === new Set(anchors).size,
        firstRow: (box.querySelector('.page-footnotes tr')?.textContent ?? '').trim(),
        bodyVisible: lines.every((line) => {
          const rect = line.getBoundingClientRect();
          return rect.top >= area.top - 0.5 && rect.bottom <= area.bottom + 0.5;
        }),
        notesVisible: Array.from(box.querySelectorAll<HTMLElement>('.page-footnotes p')).every(p => {
          const rect = p.getBoundingClientRect();
          return !p.textContent?.trim() || !!noteArea && rect.top >= noteArea.top - 0.5 && rect.bottom <= noteArea.bottom + 0.5;
        }),
      };
    });
  }, { input: Array.from(bytes), fragmentParagraphs });
}

const range = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, index) => from + index);

async function verifyPdf(browser: Browser, info: TestInfo, bytes: Uint8Array, rows: number, name: string, side = false) {
  const result = await renderDocxArtifacts(bytes,
    { browser, outputs: ['pdf'], timeoutMs: 120_000, reviewProfile: 'final', commentProfile: 'hidden' });
  const text = (await inspectPdf(result.pdf!)).pages.map(p => p.text).join(' ');
  for (const label of side ? ['Note', 'Side'] : ['Note']) {
    expect(Array.from(text.matchAll(new RegExp(`${label} table row\\s+(\\d+)`, 'g')), m => Number(m[1]))).toEqual(range(1, rows));
  }
  expect(Array.from(text.matchAll(/Body line\s+(\d+)/g), m => Number(m[1]))).toEqual(range(1, 61));
  const notes = result.pageMap.fragments.filter(f => f.story === 'footnote');
  expect(notes.length).toBeGreaterThan(0);
  if (rows === 80) expect(new Set(notes.map(f => f.pageNumber)).size).toBeGreaterThan(1);
  await info.attach(`${name}.pdf`, { body: Buffer.from(result.pdf!), contentType: 'application/pdf' });
}

for (const shape of ['rows', 'nested', 'tall-row'] as const) {
  for (const rows of [3, 80]) {
    test(`note overflow ${shape} with ${rows} lines remains visible in both pagination modes`, async ({ page }) => {
      for (const fragmentParagraphs of [true, false]) {
        const pages = await paginate(page, generateNoteTableDocx(61, 1, rows, false, shape), fragmentParagraphs);
        expect(pages.flatMap(p => p.rows)).toEqual(range(1, rows));
        expect(pages.flatMap(p => p.body)).toEqual(range(1, 61));
        expect(pages.every(p => p.bodyVisible && p.notesVisible && p.uniqueAnchors)).toBe(true);
      }
    });

    test(`note overflow ${shape} with ${rows} lines survives PDF and PageMap verification`, async ({ browser }, testInfo) => {
      test.setTimeout(120_000);
      await verifyPdf(browser, testInfo, generateNoteTableDocx(61, 1, rows, false, shape), rows, `${shape}-${rows}`);
    });
  }
}

for (const fragmentParagraphs of [true, false]) {
  test.describe(`A note's table (paragraph fragmentation ${fragmentParagraphs ? 'on' : 'off'})`, () => {
    test('splits between rows below the line that cites it, which stays visible', async ({ page }) => {
      const pages = await paginate(page, generateNoteTableDocx(61, 1, 80), fragmentParagraphs);

      expect(pages[0].body).toEqual([1]);
      expect(pages[0].bodyVisible).toBe(true);
      // Word: rows 1-51 on page 1. Allow a row either way for font metrics.
      const firstPageRows = pages[0].rows;
      expect(firstPageRows[0]).toBe(1);
      expect(firstPageRows[firstPageRows.length - 1]).toBeGreaterThanOrEqual(50);
      expect(firstPageRows[firstPageRows.length - 1]).toBeLessThanOrEqual(52);
      // Every row once, in order; page 2 resumes the table and its body follows.
      expect(pages.flatMap((p) => p.rows)).toEqual(range(1, 80));
      expect(pages[1].rows[pages[1].rows.length - 1]).toBe(80);
      expect(pages[1].body.length).toBeGreaterThan(0);
      expect(pages.every((p) => p.bodyVisible)).toBe(true);
    });

    test('splits between rows when cited lower on the page', async ({ page }) => {
      const pages = await paginate(page, generateNoteTableDocx(90, 30, 30), fragmentParagraphs);

      expect(pages[0].body).toEqual(range(1, 30));
      // Word: rows 1-18 on page 1, 19-30 on page 2.
      const firstPageRows = pages[0].rows;
      expect(firstPageRows[firstPageRows.length - 1]).toBeGreaterThanOrEqual(17);
      expect(firstPageRows[firstPageRows.length - 1]).toBeLessThanOrEqual(19);
      expect(pages.flatMap((p) => p.rows)).toEqual(range(1, 30));
      expect(pages.every((p) => p.bodyVisible)).toBe(true);
    });
  });
}

for (const shape of ['nested', 'tall-row'] as const) {
  for (const singleParagraph of [false, true]) {
    test(`two-cell ${shape} note preserves headers and ${singleParagraph ? 'wrapped' : 'separate'} paragraphs`, async ({ page }) => {
      const pages = await paginate(page, generateNoteTableDocx(61, 1, 80, true, shape, { columns: 2, singleParagraph }), true);
      expect(pages.flatMap(p => p.rows)).toEqual(range(1, 80));
      expect(pages.flatMap(p => p.sideRows)).toEqual(range(1, 80));
      expect(pages.flatMap(p => p.body)).toEqual(range(1, 61));
      expect(pages.every(p => p.bodyVisible && p.notesVisible && p.uniqueAnchors)).toBe(true);
      expect(pages.filter(p => p.rows.length > 0).every(p => p.firstRow.includes('Note table heading.'))).toBe(true);
    });
  }
  test(`two-cell ${shape} wrapped note survives PDF and PageMap verification`, async ({ browser }, info) => {
    test.setTimeout(120_000);
    await verifyPdf(browser, info,
      generateNoteTableDocx(61, 1, 80, true, shape, { columns: 2, singleParagraph: true }), 80, `two-cell-${shape}`, true);
  });
}

test('a footnote row that forbids splitting moves intact to a fresh band', async ({ page }) => {
  const pages = await paginate(page, generateNoteTableDocx(61, 42, 10, false, 'tall-row', { cantSplit: true }), true);
  expect(pages.filter(p => p.rows.length > 0).map(p => p.rows)).toEqual([range(1, 10)]);
  expect(pages.flatMap(p => p.body)).toEqual(range(1, 61));
  expect(pages.every(p => p.bodyVisible && p.notesVisible)).toBe(true);
});

test('a note table repeats its header row on the page it continues on', async ({ page }) => {
  const pages = await paginate(page, generateNoteTableDocx(61, 1, 80, true), true);

  expect(pages[0].firstRow).toMatch(/Note table heading\.$/);
  expect(pages[1].rows.length).toBeGreaterThan(0);
  expect(pages[1].firstRow).toMatch(/Note table heading\.$/);
  expect(pages.flatMap((p) => p.rows)).toEqual(range(1, 80));
});
test('a note cell with unsupported bare text keeps its complete source content', async ({ page }) => {
  const rows = Array.from({ length: 30 }, (_, i) => `<p>RawLine${i + 1}</p>`).join('');
  await page.setContent(`<style>body{font:12pt/15pt serif}p{margin:0}td{vertical-align:top;padding:0}</style>
    <div id="staging"><div id="pagination-footnote-registry"><div class="footnote-item" data-footnote-id="1">
      <span class="footnote-number">1</span><div class="footnote-content"><table><tbody><tr><td>RawPrefix${rows}</td></tr></tbody></table></div>
      </div></div><section data-section-index="0" data-page-width="400" data-page-height="320"
      data-content-width="380" data-content-height="300" data-margin-top="10" data-margin-bottom="10"
      data-margin-left="10" data-margin-right="10"><p>Cite<sup data-footnote-id="1">1</sup></p><p>BodyTail</p></section></div><div id="container"></div>`);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  const actual = await page.evaluate(() => {
    const api = (window as any).DocxodusPagination;
    const result = new api.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
      { showPageNumbers: false, fragmentParagraphs: true }).paginate();
    return {
      pages: result.totalPages,
      notes: Array.from(document.querySelectorAll('#container .page-footnotes'), n => n.textContent).join(' '),
      body: document.querySelector('#container')!.textContent,
    };
  });
  expect(actual.notes.match(/RawPrefix/g) ?? []).toHaveLength(1);
  expect(actual.notes.match(/RawLine\d+/g)).toEqual(Array.from({ length: 30 }, (_, i) => `RawLine${i + 1}`));
  expect(actual.body).toContain('BodyTail');
  expect(actual.pages).toBeLessThan(5);
});
