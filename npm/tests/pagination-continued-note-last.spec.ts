import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #1179: a continued note stays last on its page, and notes keep their citation order. Once a
 * note runs on to the next page, or a note cited on a page has been deferred to the next one, no
 * later note starts on that page; later notes start on the next page, below the continuation and
 * after the deferred notes. Word reference PDFs and LibreOffice both keep this order
 * (docs/ooxml_corner_cases.md).
 *
 * Every page here is 300pt of body at 15pt lines.
 */

const LINE_PT = 15;

function lines(prefix: string, count: number): string {
  return Array.from({ length: count }, (_, index) => `${prefix} ${index + 1}`).join('<br>');
}

/** A note of `lineCount` lines, one paragraph that may split between lines. */
function note(id: string, lineCount: number): string {
  return `
    <div class="footnote-item" data-footnote-id="${id}">
      <span class="footnote-number">${id}</span>
      <div class="footnote-content"><p>${lines(`note ${id} line`, lineCount)}</p></div>
    </div>`;
}

/** A note of `paragraphs` paragraphs of `linesEach` lines that each keep their lines together. */
function keptNote(id: string, paragraphs: number, linesEach: number): string {
  return `
    <div class="footnote-item" data-footnote-id="${id}">
      <span class="footnote-number">${id}</span>
      <div class="footnote-content">${Array.from({ length: paragraphs }, (_, index) =>
        `<p data-keep-lines="true" style="break-inside:avoid">${lines(`note ${id} part ${index + 1} line`, linesEach)}</p>`).join('')}</div>
    </div>`;
}

function reference(id: string): string {
  return `<a class="footnote-ref" data-footnote-id="${id}"><sup>${id}</sup></a>`;
}

function staging(blocks: string, notes: string): string {
  return `
    <style>
      body { font: 12pt/${LINE_PT}pt serif; }
      p { margin: 0; padding: 0; font: inherit; }
      .filler10 { height: ${LINE_PT * 10}pt; overflow: hidden; }
      table { border-collapse: collapse; } td { padding: 0; font: inherit; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">${notes}</div>
      <div data-section-index="0"
           data-page-width="400" data-page-height="320"
           data-content-width="380" data-content-height="300"
           data-margin-top="10" data-margin-right="10"
           data-margin-bottom="10" data-margin-left="10">
        ${blocks}
      </div>
    </div>
    <div id="container"></div>`;
}

interface Layout {
  pages: number;
  /** Body text of each page, one entry per block on it. */
  body: string[][];
  /** The page a note first appears on, by note id. */
  notePage: Record<string, number>;
  /** Every page's notes, in order: `<id>` for a note's first part, `<id>c` for a continuation. */
  notes: string[][];
}

async function paginate(page: Page, html: string, fragmentParagraphs = true): Promise<Layout> {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((fragmentParagraphs) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const result = new PaginationEngine(staging, container, {
      scale: 1, showPageNumbers: false, fragmentParagraphs,
    }).paginate();
    const boxes = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    const notePage: Record<string, number> = {};
    const notes = boxes.map((box, index) => Array.from(
      box.querySelectorAll<HTMLElement>('.page-footnotes [data-footnote-id]'))
      .filter((item) => !item.parentElement?.closest('[data-footnote-id]'))
      .map((item) => {
        const id = item.dataset.footnoteId!;
        if (!(id in notePage)) notePage[id] = index + 1;
        return `${id}${item.classList.contains('footnote-continuation') ? 'c' : ''}`;
      }));
    return {
      pages: result.totalPages,
      body: boxes.map((box) => Array.from(box.querySelector('.page-content')!.children)
        .map((block) => (block.textContent || '').replace(/\s+/g, ' ').trim())),
      notePage,
      notes,
    };
  }, fragmentParagraphs);
}

/** No note follows one that continues on the next page, and notes start in citation order. */
function expectContinuedNotesLast(layout: Layout): void {
  layout.notes.forEach((items, index) => {
    const last = items[items.length - 1];
    for (const item of items.slice(0, -1)) {
      const id = item.replace(/c$/, '');
      const continues = (layout.notes[index + 1] ?? []).includes(`${id}c`);
      expect(continues, `page ${index + 1}: note ${id} continues, yet ${last} follows it`).toBe(false);
    }
  });
  const starts = layout.notes.flat().filter((item) => !item.endsWith('c'));
  expect(starts).toEqual([...starts].sort((a, b) => Number(a) - Number(b)));
}

test.describe('A continued note stays last on its page', () => {
  test('a note cited after a split note starts on the next page, below the continuation', async ({ page }) => {
    // Note 1 splits on page 1; the paragraph after cites a one-line note that would fit below it.
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><p>cites B${reference('2')}</p>`, note('1', 30) + note('2', 1)));

    expectContinuedNotesLast(layout);
    expect(layout.notes[0]).toEqual(['1']);
    expect(layout.notePage['2']).toBeGreaterThan(1);
  });

  test('a note cited beside the end of a carried note starts where that note ends', async ({ page }) => {
    // Note 1 is seven paragraphs that each keep their lines together. It fills page 1 below
    // `cites A` and ends on page 2, where `cites B` opens the body and note 2 follows note 1's
    // end, beside its citation, as in the Word reference PDF of this shape at Letter size.
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><p>cites B${reference('2')}</p><p>${lines('tail', 30)}</p>`,
      keptNote('1', 7, 4) + note('2', 1)));

    expectContinuedNotesLast(layout);
    expect(layout.notes[0]).toEqual(['1']);
    expect(layout.notes[layout.notePage['2'] - 1]).toEqual(['1c', '2']);
    expect(layout.body[layout.notePage['2'] - 1][0]).toBe('cites B2');
  });

  test('notes cited beside a carried continuation keep their order behind it', async ({ page }) => {
    // Note 1 carries over three pages. Page 2 cannot hold the rest of it, so, as in the Word
    // reference PDF of this shape, it holds the note alone, and `cites B` and `cites C` open page 3,
    // where note 1 ends and notes 2 and 3 follow it in order (issue #1178).
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><div data-page-break="true"></div>`
        + `<p>cites B${reference('2')}</p><p>cites C${reference('3')}</p>`,
      note('1', 40) + note('2', 1) + note('3', 1)));

    expectContinuedNotesLast(layout);
    expect(layout.body.map((blocks) => blocks.join(' '))).toEqual(['cites A1', '', 'cites B2 cites C3']);
    expect(layout.notes[layout.notes.length - 1]).toEqual(['1c', '2', '3']);
  });

  for (const [shape, fragmentParagraphs] of [['a paragraph', false], ['a table', false], ['a table', true]] as const) {
    test(`a note cited after a deferred one does not overtake it (${shape}, fragmentation ${fragmentParagraphs ? 'on' : 'off'})`, async ({ page }) => {
      // Note 1, one 13-line paragraph that keeps its lines together, is too tall for the room the
      // filler leaves, so it is deferred whole; note 2 is a single line that would still fit.
      const citing = shape === 'a table'
        ? `<table><tr><td>cites A${reference('1')}</td></tr></table>`
        : `<p data-keep-lines="true">cites A${reference('1')}</p>`;
      const layout = await paginate(page, staging(
        `<p class="filler10">filler</p>${citing}<p>cites B${reference('2')}</p>`,
        keptNote('1', 1, 13) + note('2', 1)), fragmentParagraphs);

      expectContinuedNotesLast(layout);
      expect(layout.notePage['1']).toBeLessThanOrEqual(layout.notePage['2']);
    });
  }

  test('a note deferred behind a carried note follows it on the page where it ends', async ({ page }) => {
    // Note 2 is cited by the table on page 1, where note 1 splits, so it waits for note 1 to end.
    // Page 2 holds the rest of note 1 beside the remaining body, and note 2 follows it there
    // rather than opening a page of its own.
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><table><tr><td>cites B${reference('2')}</td></tr></table>`
        + `<p>${lines('tail', 7)}</p>`,
      note('1', 22) + note('2', 1)));

    expectContinuedNotesLast(layout);
    expect(layout.pages).toBe(2);
    expect(layout.notes[1]).toEqual(['1c', '2']);
  });
});
