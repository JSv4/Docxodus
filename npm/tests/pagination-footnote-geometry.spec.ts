import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { FOOTNOTE_PAGE, generateFootnoteDocx, twipsToPt } from './docx-footnote-fixture.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #378 — where the paginated footnote area sits vertically on the page.
 *
 * Word anchors the footnote area to the bottom of the text column: the LAST note line ends on
 * the bottom margin line, the notes stack upward with no spacing of their own (FootnoteText is
 * single-spaced with zero spacing-after), and the separator rule is drawn about one line above
 * the first note. The renderer used to add web chrome inside the bottom-anchored container —
 * 1.4 line-height, 4pt inter-note margins, a trailing back-reference link — each point of which
 * lifted the visible ink off the margin (≈13px on the tracked benchmark case).
 *
 * These assertions pin the note block and the separator SEPARATELY, so a regression in one
 * cannot hide behind correctness of the other.
 */

const PT_TO_PX = 4 / 3;
/** Sub-pixel tolerance: offsets are set in `pt` and read back from layout in `px`. */
const TOL = 0.75;

const MARGIN_PX = twipsToPt(FOOTNOTE_PAGE.marginTwips) * PT_TO_PX;

interface NoteGeometry {
  boxTop: number;
  boxBottom: number;
  notesTop: number;
  notesBottom: number;
  separator: { top: number; bottom: number; width: number } | null;
  items: { top: number; bottom: number; number: string; text: string }[];
  backrefs: number;
  lineHeight: string;
}

const MEASURE = () => {
  const container = document.getElementById('pagination-container') as HTMLElement;
  return Array.from(container.querySelectorAll('.page-box')).map((box) => {
    const r = box.getBoundingClientRect();
    const notes = box.querySelector('.page-footnotes') as HTMLElement | null;
    if (!notes) {
      return { boxTop: r.top, boxBottom: r.bottom, notesTop: 0, notesBottom: 0,
        separator: null, items: [], backrefs: 0, lineHeight: '' };
    }
    const nr = notes.getBoundingClientRect();
    const hr = notes.querySelector('hr');
    const hrRect = hr ? hr.getBoundingClientRect() : null;
    return {
      boxTop: r.top,
      boxBottom: r.bottom,
      notesTop: nr.top,
      notesBottom: nr.bottom,
      separator: hrRect
        ? { top: hrRect.top, bottom: hrRect.bottom, width: hrRect.width }
        : null,
      items: Array.from(notes.querySelectorAll('.footnote-item')).map((item) => {
        const ir = item.getBoundingClientRect();
        return {
          top: ir.top,
          bottom: ir.bottom,
          number: (item.querySelector('.footnote-number')?.textContent ?? ''),
          text: (item.textContent || '').replace(/\s+/g, ' ').trim(),
        };
      }),
      backrefs: notes.querySelectorAll('.footnote-backref').length,
      // The author-level value, not the used px: 'normal' is the single-spacing contract.
      lineHeight: getComputedStyle(notes).lineHeight,
    };
  });
};

async function paginateGeneratedDocument(page: Page, docx: Uint8Array): Promise<NoteGeometry[]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });

  const html = await page.evaluate((bytes) => {
    const converter = (window as any).Docxodus.DocumentConverter;
    return converter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
      /* paginationMode */ 1, /* paginationScale */ 1, 'page-',
      /* renderAnnotations */ false, 0, 'annot-',
      /* footnotes */ true, /* headersAndFooters */ true,
      /* trackedChanges */ false, false, false,
    ) as string;
  }, Array.from(docx));

  expect(html.startsWith('{'), `conversion failed: ${html.slice(0, 300)}`).toBe(false);

  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ measure }: { measure: string }) => {
    const staging = document.getElementById('pagination-staging') as HTMLElement;
    const container = document.getElementById('pagination-container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { scale: 1, showPageNumbers: false }).paginate();
    // eslint-disable-next-line no-eval
    return (0, eval)(`(${measure})`)();
  }, { measure: MEASURE.toString() }) as Promise<NoteGeometry[]>;
}

test.describe('Paginated footnote geometry', () => {
  let onePage: NoteGeometry;
  let twoNotes: NoteGeometry;

  test.beforeAll(async ({ browser }) => {
    const page = await browser.newPage();
    try {
      [onePage] = await paginateGeneratedDocument(page, generateFootnoteDocx(1));
      [twoNotes] = await paginateGeneratedDocument(page, generateFootnoteDocx(2));
    } finally {
      await page.close();
    }
  });

  test('the note area ends on the bottom margin line', async () => {
    expect(onePage.items.length).toBe(1);
    expect(
      onePage.boxBottom - onePage.notesBottom,
      'note area bottom vs bottom margin',
    ).toBeCloseTo(MARGIN_PX, 0);
  });

  test('the last note line sits flush at the note area bottom, with no trailing spacing', async () => {
    const last = onePage.items[onePage.items.length - 1];
    // Word puts the last FootnoteText line box ON the margin line. Anything beyond sub-pixel
    // rounding here is renderer-invented trailing space (the old 4pt item margin).
    expect(onePage.notesBottom - last.bottom, 'gap under the last note').toBeLessThanOrEqual(TOL);
  });

  test('a single note occupies one single-spaced 10pt line, not a web-spaced one', async () => {
    const item = onePage.items[0];
    // A 10pt single-spaced line is ~15.5px, and the raised superscript number legitimately
    // expands the CSS line box to ~19.5px; the old 1.4 line-height pushed the same line past
    // 21px. Bound the box so the web spacing cannot return without pinning one font's metrics.
    expect(item.bottom - item.top, 'note line box height').toBeLessThan(20.5);
    expect(item.bottom - item.top, 'note line box height').toBeGreaterThan(12);
    // The direct pin: FootnoteText is single-spaced, so the note container must not impose
    // a multiplied line-height of its own.
    expect(onePage.lineHeight, 'note area line-height').toBe('normal');
  });

  test('the separator keeps its own contract: 2in wide, one line above the first note', async () => {
    expect(onePage.separator, 'paginated note area must render a separator').not.toBeNull();
    // Word's built-in separator is two inches (192px at 96dpi), independent of the note block.
    expect(onePage.separator!.width, 'separator width').toBeCloseTo(192, 0);
    // The rule is drawn on the baseline of one empty FootnoteText line above the first note:
    // rule-to-note-ink is about a descent plus an ascent (~4-10px), not the old 6pt+slack ~11px.
    const gap = onePage.items[0].top - onePage.separator!.bottom;
    expect(gap, 'separator to first note gap').toBeGreaterThanOrEqual(2);
    expect(gap, 'separator to first note gap').toBeLessThanOrEqual(10);
  });

  test('notes stack with no renderer-invented spacing between them', async () => {
    expect(twoNotes.items.length).toBe(2);
    const [first, second] = twoNotes.items;
    expect(second.top - first.bottom, 'inter-note gap').toBeLessThanOrEqual(TOL);
    // Both notes' text must still be present and distinct.
    expect(first.text).toContain('Footnote 1 text.');
    expect(second.text).toContain('Footnote 2 text.');
    // The area still ends on the margin line with two notes stacked.
    expect(
      twoNotes.boxBottom - twoNotes.notesBottom,
      'two-note area bottom vs bottom margin',
    ).toBeCloseTo(MARGIN_PX, 0);
  });

  test('the paginated note is print-shaped: bare number, no back-reference link', async () => {
    // Word and LibreOffice print the bare superscript number and no navigation chrome; the
    // web-view <ol> section keeps its "N." labels and backrefs.
    expect(onePage.items[0].number.trim()).toBe('1');
    expect(onePage.backrefs, 'backref links inside the paginated note area').toBe(0);
    expect(twoNotes.items.map((i) => i.number.trim())).toEqual(['1', '2']);
  });
});
