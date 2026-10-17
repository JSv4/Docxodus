import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #1179: Word keeps a continued note last on its page. Once a note runs on to the next page,
 * no later note starts on the page it continues from; later notes start on the next page, below the
 * continuation, and notes keep their citation order.
 *
 * Every page here is 300pt of body at 15pt lines; the note area may take 60% of a page while body
 * text remains.
 */

const LINE_PT = 15;

interface Options {
  fragmentParagraphs?: boolean;
}

function lines(prefix: string, count: number): string {
  return Array.from({ length: count }, (_, index) => `${prefix} ${index + 1}`).join('<br>');
}

/**
 * A note of `lineCount` lines. `tabLed` opens its paragraph the way the converter writes the
 * FootnoteText style's tab after the number: an empty inline-flex box of block children.
 */
function note(id: string, lineCount: number, tabLed = false, numberInside = false): string {
  const tab = tabLed
    ? '<span style="display:inline-flex"><span style="display:block;white-space:nowrap"></span>'
      + '<span data-docx-tab="left" style="display:block;width:18pt"></span></span>'
    : '';
  // A note whose number is a run of its first paragraph, as Word stores it.
  const number = `<span class="footnote-number">${id}</span>`;
  return `
    <div class="footnote-item" data-footnote-id="${id}">
      ${numberInside ? '' : number}
      <div class="footnote-content"><p>${numberInside ? number : ''}${tab}${lines(`note ${id} line`, lineCount)}</p></div>
    </div>`;
}

function reference(id: string): string {
  return `<a class="footnote-ref" data-footnote-id="${id}"><sup>${id}</sup></a>`;
}

function staging(blocks: string, notes: string): string {
  return `
    <style>
      body { font: 12pt/${LINE_PT}pt serif; }
      p, h4 { margin: 0; padding: 0; font: inherit; }
      .filler { height: ${LINE_PT * 15}pt; overflow: hidden; }
      .filler10 { height: ${LINE_PT * 10}pt; overflow: hidden; }
      td { padding: 0; font: inherit; }
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
  /** Note text of each page. */
  notes: string[];
  /** The page a citation or a note first appears on, by note id. */
  referencePage: Record<string, number>;
  notePage: Record<string, number>;
  /** How many body elements carry each citation, across all pages. */
  referenceCount: Record<string, number>;
  /** Whether the first body block of each page starts with a list marker. */
  startsWithMarker: boolean[];
  order: string[][];
}

async function paginate(page: Page, html: string, options: Options = {}): Promise<Layout> {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((options) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const result = new PaginationEngine(staging, container, {
      scale: 1, showPageNumbers: false, fragmentParagraphs: options.fragmentParagraphs ?? true,
    }).paginate();
    const boxes = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    const referencePage: Record<string, number> = {};
    const notePage: Record<string, number> = {};
    const referenceCount: Record<string, number> = {};
    boxes.forEach((box, index) => {
      box.querySelectorAll<HTMLElement>('.page-content [data-footnote-id]').forEach((ref) => {
        const id = ref.dataset.footnoteId!;
        referenceCount[id] = (referenceCount[id] ?? 0) + 1;
        if (!(id in referencePage)) referencePage[id] = index + 1;
      });
      box.querySelectorAll<HTMLElement>('.page-footnotes [data-footnote-id]').forEach((item) => {
        const id = item.dataset.footnoteId!;
        if (!(id in notePage)) notePage[id] = index + 1;
      });
    });
    return {
      pages: result.totalPages,
      body: boxes.map((box) => Array.from(box.querySelector('.page-content')!.children)
        .map((block) => (block.textContent || '').replace(/\s+/g, ' ').trim())),
      notes: boxes.map((box) => (box.querySelector('.page-footnotes')?.textContent || '').replace(/\s+/g, ' ')),
      referencePage,
      notePage,
      referenceCount,
      order: boxes.map((box) => Array.from(box.querySelectorAll<HTMLElement>('.page-footnotes [data-footnote-id]'))
        .filter((el) => !el.parentElement?.closest('[data-footnote-id]'))
        .map((el) => `${el.dataset.footnoteId}${el.classList.contains('footnote-continuation') ? 'c' : ''}:${(el.textContent||'').replace(/\s+/g,' ').trim().slice(-14)}`)),
      startsWithMarker: boxes.map((box) =>
        box.querySelector('.page-content')?.firstElementChild?.querySelector('[data-list-marker]') != null),
    };
  }, options);
}


/** Every page's notes, in order: `<id>` for a note's first part, `<id>c` for a continuation. */
function noteItems(layout: Layout): string[][] {
  return layout.order.map((items) => items.map((item) => item.split(':')[0]));
}

/** No note follows one that continues on the next page, and notes start in citation order. */
function expectContinuedNotesLast(layout: Layout): void {
  const items = noteItems(layout);
  items.forEach((page, index) => {
    const last = page[page.length - 1];
    for (const item of page.slice(0, -1)) {
      const id = item.replace(/c$/, '');
      const continues = (items[index + 1] ?? []).includes(`${id}c`);
      expect(continues, `page ${index + 1}: note ${id} continues, yet ${last} follows it`).toBe(false);
    }
  });
  const starts = items.flat().filter((item) => !item.endsWith('c'));
  expect(starts).toEqual([...starts].sort((a, b) => Number(a) - Number(b)));
}

test.describe('A continued note stays last on its page', () => {
  test('a note cited after a split note starts on the next page, below the continuation', async ({ page }) => {
    // Note 1 splits on page 1; the paragraph after cites a one-line note that would fit below it.
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><p>cites B${reference('2')}</p>`, note('1', 30) + note('2', 1)));

    expectContinuedNotesLast(layout);
    expect(noteItems(layout)[0]).toEqual(['1']);
    expect(layout.notePage['2']).toBeGreaterThan(1);
  });

  test('a note deferred to a page whose continuation runs on moves on with it', async ({ page }) => {
    // Note 1 is eight paragraphs that each keep their lines together, so every page it fills has
    // room to spare that a one-line note would otherwise take.
    const unsplittable = `<div class="footnote-item" data-footnote-id="1"><span class="footnote-number">1</span>`
      + `<div class="footnote-content">${Array.from({ length: 8 }, (_, i) =>
        `<p data-keep-lines="true" style="break-inside:avoid">${lines(`A${i}`, 4)}</p>`).join('')}</div></div>`;
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><p>cites B${reference('2')}</p><p>${lines('tail', 30)}</p>`,
      unsplittable + note('2', 1)));

    expectContinuedNotesLast(layout);
    // Note 2 starts on the page where note 1 ends.
    const items = noteItems(layout);
    expect(items[layout.notePage['2'] - 1]).toEqual(['1c', '2']);
  });

  test('notes cited beside a carried continuation keep their order behind it', async ({ page }) => {
    // Note 1 carries over several pages. Page 2 cites notes 2 and 3; neither may start before
    // note 1 ends, and note 3 may not overtake note 2.
    const layout = await paginate(page, staging(
      `<p>cites A${reference('1')}</p><div data-page-break="true"></div>`
        + `<p>cites B${reference('2')}</p><p>cites C${reference('3')}</p>`,
      note('1', 40) + note('2', 1) + note('3', 1)));

    expectContinuedNotesLast(layout);
    const items = noteItems(layout);
    expect(items[items.length - 1]).toEqual(['1c', '2', '3']);
  });
});
