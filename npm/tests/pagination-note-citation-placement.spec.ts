import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #1151: where the paginator puts a footnote relative to the line that cites it.
 *
 * LibreOffice begins 93 of the 94 notes of the NVCA model charter on the page that cites them; the
 * paginator used to defer a note it could not start to the next page while leaving its citation
 * behind, and the deferred notes piled up into pages that held nothing but notes. A citing
 * paragraph now breaks before the citing line, which moves to the page where the note begins.
 *
 * Every page here is 300pt of body at 15pt lines, so the geometry is exact.
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
function note(id: string, lineCount: number, tabLed = false): string {
  const tab = tabLed
    ? '<span style="display:inline-flex"><span style="display:block;white-space:nowrap"></span>'
      + '<span data-docx-tab="left" style="display:block;width:18pt"></span></span>'
    : '';
  return `
    <div class="footnote-item" data-footnote-id="${id}">
      <span class="footnote-number">${id}</span>
      <div class="footnote-content"><p>${tab}${lines(`note ${id} line`, lineCount)}</p></div>
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
      startsWithMarker: boxes.map((box) =>
        box.querySelector('.page-content')?.firstElementChild?.querySelector('[data-list-marker]') != null),
    };
  }, options);
}

/** 15 lines of filler, then a four-line paragraph whose last line cites a six-line note: the
 * paragraph fits, leaving one line of space, too little to start the note in. */
function citingLastLine(tag = 'p', attributes = ''): string {
  return staging(
    `<p class="filler">filler</p>
     <${tag} class="citing" ${attributes}>${lines('citing', 3)}<br>citing 4${reference('1')}</${tag}>`,
    note('1', 6));
}

test.describe('Paginated note placement beside its citation', () => {
  test('the citing line moves to the page where its note begins', async ({ page }) => {
    // The paragraph's four lines fit beneath the filler, but no line of the note does.
    const layout = await paginate(page, citingLastLine());

    expect(layout.referencePage['1']).toBe(layout.notePage['1']);
    // Only the citing line moves: the three before it stay on the first page.
    expect(layout.body[0].join(' ')).toContain('citing 3');
    expect(layout.body[0].join(' ')).not.toContain('citing 4');
    expect(layout.body[1].join(' ')).toContain('citing 4');
    expect(layout.referenceCount['1']).toBe(1);
  });

  test('a citing paragraph that cannot be split moves whole with its citation', async ({ page }) => {
    // Word moves a paragraph that keeps its lines together to the page where its note begins.
    const layout = await paginate(page, citingLastLine('p', 'data-keep-lines="true"'));

    expect(layout.referencePage['1']).toBe(layout.notePage['1']);
    expect(layout.body[0].join(' ')).not.toContain('citing 1');
  });

  test('without paragraph fragmentation the citing paragraph stays and its note follows', async ({ page }) => {
    // The paginated editor splits no paragraph. Moving every such paragraph whole would leave far
    // more of each page empty than Word's split does, so its note is deferred as before.
    const layout = await paginate(page, citingLastLine(), { fragmentParagraphs: false });

    expect(layout.body[0].join(' ')).toContain('citing 4');
    expect(layout.notePage['1']).toBe(2);
  });

  test('a heading-styled citing paragraph breaks like any other paragraph', async ({ page }) => {
    // A Word paragraph with an outline level is written as h1-h6.
    const layout = await paginate(page, citingLastLine('h4'));

    expect(layout.referencePage['1']).toBe(layout.notePage['1']);
    expect(layout.body[0].join(' ')).toContain('citing 3');
    expect(layout.body[1].join(' ')).toContain('citing 4');
  });

  test('a numbered paragraph breaks across pages with its marker kept whole', async ({ page }) => {
    // The converter writes a list number and its tab as one inline-flex box of block children.
    const marker = '<span data-list-marker="true" style="display:inline-flex">'
      + '<span style="display:block;white-space:nowrap"><span data-list-marker="true">1.</span></span>'
      + '<span data-docx-tab="left" style="display:block;width:24pt"></span></span>';
    const layout = await paginate(page, staging(
      `<p class="filler">filler</p><p>${marker}${lines('item', 8)}</p>`, ''));

    expect(layout.pages).toBe(2);
    expect(layout.body[0].join(' ')).toContain('1.item 1');
    expect(layout.startsWithMarker).toEqual([false, false]);
    expect(layout.body[1].join(' ')).toMatch(/^item \d/);
    expect(layout.body[1].join(' ')).not.toContain('1.');
  });

  test('a citation moves when the note area is full though as large as any page\'s', async ({ page }) => {
    // Note 1 continues from page 1 and fills page 2's note area, which is as large as any page's.
    // The paragraph opening page 2 cites note 2 on its second line: that line fits, but no room
    // is left for the note, while page 3 will have some. So the first line stays on page 2 and
    // the citing line moves to page 3 with its note.
    const layout = await paginate(page, staging(
      `<p>cites one${reference('1')}</p><div data-page-break="true"></div>`
        + `<p>${lines('two', 1)}<br>two 2${reference('2')}</p><p>after</p>`,
      note('1', 26) + note('2', 2)));

    expect(layout.referencePage['2']).toBe(layout.notePage['2']);
    expect(layout.body[1].join(' ')).toContain('two 1');
  });

  test('a page whose note area a continuation fills still takes body text that cites nothing', async ({ page }) => {
    // The continuation of note 1 fills page 2's capped note area exactly. The paragraph that
    // follows cites no note, so the cap has nothing to judge; it used to be refused all the same,
    // leaving page 2 with notes and no body.
    const layout = await paginate(page, staging(
      `<p>cites one${reference('1')}</p><p>${lines('filler', 7)}</p><p>after</p>`,
      note('1', 24)));

    const lastWithBody = layout.body.map((blocks) => blocks.length > 0).lastIndexOf(true);
    expect(layout.body.slice(0, lastWithBody + 1).every((blocks) => blocks.length > 0)).toBe(true);
  });

  test('a note keeps the citing paragraph whole when the note is what does not fit', async ({ page }) => {
    // The paragraph's own lines fit; its 30-line note must split. Only the note splits:
    // the paragraph is not cut down word by word to make room for a note too tall for any page.
    const layout = await paginate(page, staging(
      `<p class="citing">alpha${reference('1')}<br>beta<br>gamma</p>${'<p>after</p>'.repeat(3)}`,
      note('1', 30)));

    expect(layout.body[0][0]).toBe('alpha1betagamma');
    expect(layout.referenceCount['1']).toBe(1);
    expect(layout.notePage['1']).toBe(1);
  });

  test('a note that opens with a tab after its number continues on the next page', async ({ page }) => {
    // 30 lines of note cannot fit one page; the note splits instead of moving or clipping whole.
    const layout = await paginate(page, staging(
      `<p>cites${reference('1')}</p><p>${lines('body', 10)}</p>`,
      note('1', 30, true)));

    expect(layout.notePage['1']).toBe(1);
    expect(layout.notes[0]).toContain('note 1 line 1');
    expect(layout.notes[0]).not.toContain('note 1 line 30');
    expect(layout.notes.slice(1).some((text) => text.includes('note 1 line 30'))).toBe(true);
  });
});
