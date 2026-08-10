import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import {
  MARGIN_PT,
  PAGE_HEIGHT_PT,
  SEPARATOR_WIDTH_IN,
  generateFootnoteDocx,
} from './docx-footnote-fixture.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #378 — where a page's footnote block sits, and what it is made of.
 *
 * The separator's WIDTH was already correct; its POSITION was not, and the two were being
 * confused because one number (the block's height) moved both. The note area is bottom-aligned
 * to the body text band, so everything above that edge is the sum of what the block contains:
 *
 *   [separator band: one line of the note font, with the rule on its baseline]
 *   [note, note, … with spacing BETWEEN them]
 *   ─────────────────────────────────── body band bottom
 *
 * Two defects moved the whole block up by 14 px on a Word-default page: a `margin-bottom` on
 * the last note left dead space below content that is anchored to that edge, and the separator
 * was a bare 1 px rule with a hard-coded 6 pt gap under it and nothing above, rather than a
 * paragraph line. These assertions therefore pin the separator and the note block SEPARATELY,
 * so a future change to one cannot be masked by a compensating change to the other.
 */

const PT_TO_PX = 4 / 3;
const TOL = 1;

interface NoteGeometry {
  page: number;
  boxTop: number;
  boxBottom: number;
  bodyBottom: number;
  notesTop: number;
  notesBottom: number;
  separatorTop: number;
  separatorBottom: number;
  separatorIsContinuation: boolean;
  ruleTop: number;
  ruleBottom: number;
  ruleWidth: number;
  noteLineHeight: number;
  items: Array<{ top: number; bottom: number; marginTop: number }>;
}

const MEASURE = () => {
  const container = document.getElementById('pagination-container') as HTMLElement;
  return Array.from(container.querySelectorAll('.page-box')).map((box, i) => {
    const b = box.getBoundingClientRect();
    const notes = box.querySelector('.page-footnotes') as HTMLElement | null;
    const body = box.querySelector('.page-content') as HTMLElement;
    const sep = notes?.querySelector('.footnote-separator') as HTMLElement | null;
    const rule = sep?.querySelector('hr') as HTMLElement | null;
    const rect = (el: Element | null) => (el ? el.getBoundingClientRect() : null);
    // One line of the note area's own font, measured rather than parsed: the area's resolved
    // `line-height` is `normal`, which is not a number the document can state for it.
    let noteLineHeight = NaN;
    if (notes) {
      const probe = document.createElement('div');
      probe.textContent = ' ';
      notes.appendChild(probe);
      noteLineHeight = probe.getBoundingClientRect().height;
      notes.removeChild(probe);
    }
    return {
      page: i + 1,
      boxTop: b.top,
      boxBottom: b.bottom,
      bodyBottom: rect(body)!.bottom,
      notesTop: rect(notes)?.top ?? NaN,
      notesBottom: rect(notes)?.bottom ?? NaN,
      separatorTop: rect(sep)?.top ?? NaN,
      separatorBottom: rect(sep)?.bottom ?? NaN,
      separatorIsContinuation: !!sep?.classList.contains('footnote-separator-continuation'),
      ruleTop: rect(rule)?.top ?? NaN,
      ruleBottom: rect(rule)?.bottom ?? NaN,
      ruleWidth: rect(rule)?.width ?? NaN,
      noteLineHeight,
      items: Array.from(notes?.querySelectorAll('.footnote-item') ?? []).map((el) => ({
        top: el.getBoundingClientRect().top,
        bottom: el.getBoundingClientRect().bottom,
        marginTop: parseFloat(getComputedStyle(el as HTMLElement).marginTop),
      })),
    };
  });
};

async function paginate(page: Page, docx: Uint8Array): Promise<NoteGeometry[]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });

  const html = await page.evaluate((bytes) => {
    const c = (window as any).Docxodus.DocumentConverter;
    return c.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
      /* paginationMode */ 1, /* paginationScale */ 1, 'page-',
      false, 0, 'annot-',
      /* footnotes */ true, /* headersAndFooters */ false,
      false, false, false,
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

test.describe('Paginated footnote block geometry', () => {
  test('the note block is bottom-aligned to the body band, with nothing below the last note',
    async ({ page }) => {
      const pages = await paginate(page, generateFootnoteDocx({ paragraphs: [1] }));
      expect(pages).toHaveLength(1);
      const p = pages[0];

      // The anchor: the block's bottom edge IS the bottom of the body text band.
      const bodyBandBottom = p.boxTop + (PAGE_HEIGHT_PT - MARGIN_PT) * PT_TO_PX;
      expect(p.notesBottom, 'note block bottom vs body band bottom')
        .toBeCloseTo(bodyBandBottom, 0);

      // And the last note ends ON that edge — a trailing margin there lifts the entire block,
      // separator included, off the edge it is anchored to.
      expect(p.items).toHaveLength(1);
      expect(p.items[0].bottom, 'last note bottom vs note block bottom')
        .toBeCloseTo(p.notesBottom, 0);
    });

  test('the separator is one line of the note font with the rule on its baseline',
    async ({ page }) => {
      const pages = await paginate(page, generateFootnoteDocx({ paragraphs: [1] }));
      const p = pages[0];

      // `w:separator` is a run in a paragraph, so the band is a LINE, not a bare rule.
      expect(p.separatorBottom - p.separatorTop, 'separator band height vs one note line')
        .toBeCloseTo(p.noteLineHeight, 0);

      // The rule sits on that line's baseline: real space above it, real space below it, both
      // coming from the font's own metrics rather than a tuned margin.
      expect(p.ruleTop - p.separatorTop, 'space above the rule').toBeGreaterThan(2);
      expect(p.separatorBottom - p.ruleBottom, 'space below the rule').toBeGreaterThan(2);
      // Baseline, not centre: a text baseline sits well below the middle of its line box.
      expect(p.ruleBottom - p.separatorTop)
        .toBeGreaterThan((p.separatorBottom - p.separatorTop) / 2);

      // The band is the top of the block, and the notes follow it.
      expect(p.separatorTop).toBeCloseTo(p.notesTop, 0);
      expect(p.items[0].top).toBeCloseTo(p.separatorBottom, 0);
    });

  test('the separator keeps Word\'s two-inch default width', async ({ page }) => {
    const pages = await paginate(page, generateFootnoteDocx({ paragraphs: [1] }));
    expect(pages[0].ruleWidth).toBeCloseTo(SEPARATOR_WIDTH_IN * 96, 0);
    expect(pages[0].separatorIsContinuation).toBe(false);
  });

  test('note spacing falls between notes, never after the last one', async ({ page }) => {
    const pages = await paginate(page, generateFootnoteDocx({ paragraphs: [3] }));
    const p = pages[0];
    expect(p.items.length).toBe(3);

    expect(p.items[0].marginTop, 'first note has no leading gap').toBeCloseTo(0, 0);
    for (let i = 1; i < p.items.length; i++) {
      expect(p.items[i].marginTop, `gap before note ${i + 1}`).toBeGreaterThan(2);
    }
    expect(p.items[p.items.length - 1].bottom, 'last note bottom vs block bottom')
      .toBeCloseTo(p.notesBottom, 0);
  });

  test('the body band never runs into the note block', async ({ page }) => {
    const pages = await paginate(page, generateFootnoteDocx({ paragraphs: [1, 1, 1] }));
    for (const p of pages) {
      if (Number.isNaN(p.notesTop)) continue;
      expect(p.bodyBottom, `page ${p.page} body vs notes`).toBeLessThanOrEqual(p.notesTop + TOL);
    }
  });
});

/**
 * `w:continuationSeparator` — the separator a page shows when it opens with note text carried
 * over from the previous page. Word draws it across the whole text column instead of the
 * two-inch default, and the engine used to draw the ordinary separator on every page.
 */
const CONTINUATION_STAGING = `
  <style>
    #pagination-staging { font: 12px/12px Arial; }
    .body { height: 20pt; margin: 0; }
    .footnote-item { margin: 0; }
    .footnote-content > p { height: 10pt; margin: 0; }
  </style>
  <div id="pagination-staging">
    <div id="pagination-footnote-registry">
      <div class="footnote-item" data-footnote-id="f1">
        <span class="footnote-number">1</span>
        <span class="footnote-content">${
          Array.from({ length: 30 }, (_, i) => `<p>note line ${i}</p>`).join('')
        }</span>
      </div>
    </div>
    <div data-section-index="0"
         data-page-width="312" data-page-height="312"
         data-content-width="300" data-content-height="300"
         data-margin-top="6" data-margin-right="6"
         data-margin-bottom="6" data-margin-left="6">
      <p class="body">body 0 <sup data-footnote-id="f1">1</sup></p>
      ${Array.from({ length: 10 }, (_, i) => `<p class="body">body ${i + 1}</p>`).join('')}
    </div>
  </div>
  <div id="pagination-container"></div>`;

test.describe('Continuation separator', () => {
  test('a page that opens with carried-over note text shows the full-width separator',
    async ({ page }) => {
      await page.setContent(CONTINUATION_STAGING);
      await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
      const seps = await page.evaluate(() => {
        const staging = document.getElementById('pagination-staging') as HTMLElement;
        const container = document.getElementById('pagination-container') as HTMLElement;
        const { PaginationEngine } = (window as any).DocxodusPagination;
        new PaginationEngine(staging, container, {
          scale: 1, showPageNumbers: false, fragmentParagraphs: true,
        }).paginate();
        return Array.from(container.querySelectorAll('.page-box')).map((box) => {
          const notes = box.querySelector('.page-footnotes') as HTMLElement | null;
          const sep = notes?.querySelector('.footnote-separator') as HTMLElement | null;
          return {
            hasNotes: !!notes,
            hasContinuation: !!notes?.querySelector('.footnote-continuation'),
            isContinuationSeparator: !!sep?.classList.contains('footnote-separator-continuation'),
            ruleWidth: sep ? (sep.querySelector('hr') as HTMLElement).getBoundingClientRect().width : NaN,
            notesWidth: notes ? notes.getBoundingClientRect().width : NaN,
          };
        });
      });

      const carried = seps.filter((s) => s.hasContinuation);
      expect(carried.length, 'the fixture must actually split a note across pages')
        .toBeGreaterThan(0);
      for (const s of carried) {
        expect(s.isContinuationSeparator).toBe(true);
        expect(s.ruleWidth).toBeCloseTo(s.notesWidth, 0);
      }
      // The page the note STARTS on keeps the ordinary two-inch separator.
      const started = seps.filter((s) => s.hasNotes && !s.hasContinuation);
      expect(started.length).toBeGreaterThan(0);
      expect(started.every((s) => !s.isContinuationSeparator)).toBe(true);
    });
});
