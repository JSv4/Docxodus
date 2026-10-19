import { test, expect, type Page } from '@playwright/test';
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
}

async function paginate(page: Page, bytes: Uint8Array, fragmentParagraphs: boolean): Promise<Page_[]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ input, fragmentParagraphs }) => {
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      /* footnotes */ true, false, false, false, false, false, null, false,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    new (window as any).DocxodusPagination.PaginationEngine(
      host.querySelector('#pagination-staging'), host.querySelector('#pagination-container'),
      { showPageNumbers: false, fragmentParagraphs },
    ).paginate();
    const numbers = (text: string, pattern: RegExp) => Array.from(text.matchAll(pattern), (match) => Number(match[1]));
    return Array.from(host.querySelectorAll<HTMLElement>('.page-box')).map((box) => {
      const content = box.querySelector<HTMLElement>('.page-content')!;
      const area = content.getBoundingClientRect();
      const lines = Array.from(content.querySelectorAll<HTMLElement>('p'));
      const noteArea = box.querySelector<HTMLElement>('.page-footnotes')?.getBoundingClientRect();
      return {
        body: numbers(content.textContent ?? '', /Body line (\d+)/g),
        rows: numbers(box.querySelector('.page-footnotes')?.textContent ?? '', /Note table row (\d+)/g),
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

for (const shape of ['rows', 'nested'] as const) {
  const rows = 3;
  test(`fitting note ${shape} with ${rows} lines remains visible in both pagination modes`, async ({ page }) => {
    for (const fragmentParagraphs of [true, false]) {
      const pages = await paginate(page, generateNoteTableDocx(61, 1, rows, false, shape === 'nested'), fragmentParagraphs);
      expect(pages.flatMap(p => p.rows)).toEqual(range(1, rows));
      expect(pages.flatMap(p => p.body)).toEqual(range(1, 61));
      expect(pages.every(p => p.bodyVisible && p.notesVisible)).toBe(true);
    }
  });

  test(`fitting note ${shape} with ${rows} lines survives PDF and PageMap verification`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(generateNoteTableDocx(61, 1, rows, false, shape === 'nested'),
      { browser, outputs: ['pdf'], timeoutMs: 120_000, reviewProfile: 'final', commentProfile: 'hidden' });
    const text = (await inspectPdf(result.pdf!)).pages.map(p => p.text).join(' ');
    expect(Array.from(text.matchAll(/Note table row\s+(\d+)/g), m => Number(m[1]))).toEqual(range(1, rows));
    expect(Array.from(text.matchAll(/Body line\s+(\d+)/g), m => Number(m[1]))).toEqual(range(1, 61));
    expect(result.pageMap.fragments.some(f => f.story === 'footnote')).toBe(true);
    await testInfo.attach(`${shape}-${rows}.pdf`, { body: Buffer.from(result.pdf!), contentType: 'application/pdf' });
  });
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

test('a note table repeats its header row on the page it continues on', async ({ page }) => {
  const pages = await paginate(page, generateNoteTableDocx(61, 1, 80, true), true);

  expect(pages[0].firstRow).toMatch(/Note table heading\.$/);
  expect(pages[1].rows.length).toBeGreaterThan(0);
  expect(pages[1].firstRow).toMatch(/Note table heading\.$/);
  expect(pages.flatMap((p) => p.rows)).toEqual(range(1, 80));
});

