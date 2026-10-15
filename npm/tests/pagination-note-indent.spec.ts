import { expect, Page, test } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { FOOTNOTE_PAGE, generateFootnoteDocx, twipsToPt } from './docx-footnote-fixture.js';
import { generateLongEndnoteDocx } from './docx-page-map-fixture.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #1130: a paginated note paragraph keeps its own indentation on every line. In Word the
 * note number is a run of the first note paragraph, so a hanging indent (`w:left=360
 * w:hanging=360`) starts the first line, number included, at the band edge and every wrapped
 * line 18pt in. The first paragraph used to be forced `display: inline` to share a line with a
 * number rendered beside it, and an inline paragraph ignores its own margins and text-indent:
 * its wrapped lines started at the band edge.
 */

const HANGING = '<w:ind w:left="360" w:hanging="360"/>';
const CONTENT_WIDTH_PT = twipsToPt(FOOTNOTE_PAGE.widthTwips - 2 * FOOTNOTE_PAGE.marginTwips);
// generateLongEndnoteDocx's page: 9000 twips wide with 720 + 1800 twips side margins.
const ENDNOTE_CONTENT_WIDTH_PT = twipsToPt(9000 - 720 - 1800);

interface ParagraphLines {
  /** Each line's start, in points from the band's left edge. */
  starts: number[];
  /** How many note numbers start the paragraph's first line. */
  numbers: number;
  /** Whether the number is on the paragraph's first line. */
  numberOnFirstLine: boolean | null;
  text: string;
}

interface NoteBand {
  pageNumber: number;
  numbers: number;
  paragraphs: ParagraphLines[];
}

async function readyPage(page: Page): Promise<void> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
}

/**
 * Convert with the real converter, paginate, and read each note paragraph's line starts from
 * its text's line boxes (Range client rects), relative to the edge of the band holding it: the
 * note band for footnotes, the page's content area for endnotes (which flow with the body).
 */
async function measureNoteLines(
  page: Page,
  docx: Uint8Array,
  endnotes: boolean,
  contentWidthPt = CONTENT_WIDTH_PT,
): Promise<NoteBand[]> {
  return page.evaluate(({ bytes, endnotes, contentWidthPt }) => {
    const D = (window as any).Docxodus;
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', false, '', -1, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      /* renderFootnotesAndEndnotes */ true, false, false, false, false, false, null,
      /* stampAnchors */ false,
    );
    if (html.startsWith('{')) throw new Error(`conversion failed: ${html.slice(0, 300)}`);
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
      showPageNumbers: false,
      fragmentParagraphs: true,
    }).paginate();

    const numberSelector = '.footnote-number, [data-endnote-number]';
    const bandSelector = endnotes ? '.page-content' : '.page-footnotes';
    return Array.from(container.querySelectorAll<HTMLElement>(bandSelector)).map((band) => {
      const bandRect = band.getBoundingClientRect();
      // The page box may be zoomed; measure in the band's own points.
      const ptPerPx = contentWidthPt / bandRect.width;
      const lineRects = (node: Node): DOMRect[] => {
        const range = document.createRange();
        range.selectNodeContents(node);
        return Array.from(range.getClientRects()).filter((rect) => rect.width > 0.5);
      };
      const paragraphs = Array.from(band.querySelectorAll<HTMLElement>('p'))
        // Endnotes share the page with the citing body paragraph; keep only note text.
        .filter((p) => !endnotes || /endnote\d/.test(p.textContent ?? ''))
        .filter((p) => (p.textContent ?? '').trim().length > 0)
        .map((p) => {
          const lines: Array<{ top: number; left: number }> = [];
          // A number drawn beside the content, not inside it, still starts the first line.
          const item = p.closest<HTMLElement>('.footnote-item, .footnote-continuation');
          const outsideNumber = item?.querySelector<HTMLElement>(':scope > .footnote-number');
          const number = p.querySelector<HTMLElement>(numberSelector)
            ?? (outsideNumber && item!.querySelector('p') === p ? outsideNumber : null);
          for (const rect of [...lineRects(p), ...(number ? lineRects(number) : [])]) {
            const line = lines.find((l) => Math.abs(l.top - rect.top) < 3);
            if (line) line.left = Math.min(line.left, rect.left);
            else lines.push({ top: rect.top, left: rect.left });
          }
          lines.sort((a, b) => a.top - b.top);
          const numberTop = number ? lineRects(number)[0]?.top : undefined;
          const clone = p.cloneNode(true) as HTMLElement;
          clone.querySelectorAll(numberSelector).forEach((n) => n.remove());
          return {
            starts: lines.map((l) => Math.round((l.left - bandRect.left) * ptPerPx * 10) / 10),
            numbers: number ? 1 : 0,
            numberOnFirstLine: numberTop === undefined || lines.length === 0
              ? null
              : Math.abs(numberTop - lines[0].top) < 3,
            text: (clone.textContent ?? '').replace(/\s+/g, ' ').trim(),
          };
        });
      return {
        pageNumber: Number(band.closest<HTMLElement>('.page-box')?.dataset.pageNumber ?? '0'),
        numbers: band.querySelectorAll(numberSelector).length,
        paragraphs,
      };
    }).filter((band) => band.paragraphs.length > 0);
  }, { bytes: Array.from(docx), endnotes, contentWidthPt });
}

const near = (value: number, expected: number) => Math.abs(value - expected) <= 1;

test.describe('paginated note paragraph indentation (issue #1130)', () => {
  test.beforeEach(async ({ page }) => {
    await readyPage(page);
  });

  test('a hanging-indented first footnote paragraph wraps its lines to its indent', async ({ page }) => {
    const bands = await measureNoteLines(page, generateFootnoteDocx(1, 3, 1, 70, {}, HANGING), false);
    expect(bands).toHaveLength(1);
    const [note] = bands[0].paragraphs;
    expect(note.numbers).toBe(1);
    expect(note.numberOnFirstLine).toBe(true);
    expect(note.starts.length).toBeGreaterThan(2);
    // First line, number included, at the band edge; every wrapped line 18pt in.
    expect(near(note.starts[0], 0), `line starts ${note.starts}`).toBe(true);
    for (const start of note.starts.slice(1)) {
      expect(near(start, 18), `line starts ${note.starts}`).toBe(true);
    }
  });

  test('every paragraph of a multi-paragraph footnote keeps its hanging indent', async ({ page }) => {
    const bands = await measureNoteLines(page, generateFootnoteDocx(1, 3, 2, 50, {}, HANGING), false);
    expect(bands).toHaveLength(1);
    const paragraphs = bands[0].paragraphs;
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs.map((p) => p.numbers)).toEqual([1, 0]);
    for (const paragraph of paragraphs) {
      expect(paragraph.starts.length).toBeGreaterThan(2);
      expect(near(paragraph.starts[0], 0), `line starts ${paragraph.starts}`).toBe(true);
      for (const start of paragraph.starts.slice(1)) {
        expect(near(start, 18), `line starts ${paragraph.starts}`).toBe(true);
      }
    }
  });

  test('an unindented footnote keeps its number and text on one line at the band edge', async ({ page }) => {
    const bands = await measureNoteLines(page, generateFootnoteDocx(1, 3, 1, 70), false);
    const [note] = bands[0].paragraphs;
    expect(note.numbers).toBe(1);
    expect(note.numberOnFirstLine).toBe(true);
    expect(note.starts.length).toBeGreaterThan(2);
    expect(note.starts.every((start) => near(start, 0)), `line starts ${note.starts}`).toBe(true);
  });

  test('a hanging-indented footnote split across pages keeps its indent and one number', async ({ page }) => {
    const words = 900;
    const bands = await measureNoteLines(page, generateFootnoteDocx(1, 3, 1, words, {}, HANGING), false);
    expect(bands.length).toBeGreaterThan(1);
    // The number is drawn once, on the citing page, and the continuation does not repeat it.
    expect(bands.map((band) => band.numbers)).toEqual([1, ...bands.slice(1).map(() => 0)]);
    const [first, ...rest] = bands.map((band) => band.paragraphs);
    expect(first[0].numberOnFirstLine).toBe(true);
    expect(near(first[0].starts[0], 0), `line starts ${first[0].starts}`).toBe(true);
    for (const start of first[0].starts.slice(1)) {
      expect(near(start, 18), `line starts ${first[0].starts}`).toBe(true);
    }
    // A continuation fragment is the rest of the hanging paragraph: every line sits at 18pt.
    for (const paragraph of rest.flat()) {
      expect(paragraph.starts.every((start) => near(start, 18)), `line starts ${paragraph.starts}`).toBe(true);
    }
    // No word is lost or repeated across the split.
    const text = bands.flatMap((band) => band.paragraphs.map((p) => p.text)).join(' ').split(/\s+/);
    expect(text).toEqual(Array.from({ length: words }, (_, i) => `footnote-1-1-${i + 1}`));
  });

  test('a hanging-indented endnote paragraph wraps its lines to its indent', async ({ page }) => {
    const bands = await measureNoteLines(page, generateLongEndnoteDocx(80, HANGING), true,
      ENDNOTE_CONTENT_WIDTH_PT);
    const notes = bands.flatMap((band) => band.paragraphs);
    expect(notes).toHaveLength(1);
    const [note] = notes;
    expect(note.numbers).toBe(1);
    expect(note.numberOnFirstLine).toBe(true);
    expect(note.starts.length).toBeGreaterThan(2);
    expect(near(note.starts[0], 0), `line starts ${note.starts}`).toBe(true);
    for (const start of note.starts.slice(1)) {
      expect(near(start, 18), `line starts ${note.starts}`).toBe(true);
    }
  });
  test('an exact-spaced note number shares its first line\'s exact baseline', async ({ page }) => {
    // Issue #1120's exact baseline: the number, now inside its paragraph, moves with that line.
    const docx = generateFootnoteDocx(1, 3, 1, 70, {}, '<w:spacing w:line="360" w:lineRule="exact"/>');
    const tops = await page.evaluate(({ bytes }) => {
      const D = (window as any).Docxodus;
      const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Document', 'docx-', false, '', -1, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', true, false, false, false, false, false, null, false,
      );
      const host = document.createElement('div');
      host.innerHTML = html;
      document.body.appendChild(host);
      const container = host.querySelector<HTMLElement>('#pagination-container')!;
      new (window as any).DocxodusPagination.PaginationEngine(
        host.querySelector<HTMLElement>('#pagination-staging')!, container, { showPageNumbers: false },
      ).paginate();
      const runs = Array.from(container.querySelectorAll<HTMLElement>('.page-footnotes [data-docx-exact-run]'));
      const isNumber = (run: HTMLElement) => !!run.closest('.footnote-number');
      return {
        number: runs.filter(isNumber).map((run) => run.style.top),
        text: runs.filter((run) => !isNumber(run)).map((run) => run.style.top),
      };
    }, { bytes: Array.from(docx) });
    expect(tops.number).toHaveLength(1);
    expect(tops.text.length).toBeGreaterThan(0);
    expect(Number.parseFloat(tops.number[0])).not.toBe(0);
    expect(new Set([...tops.number, ...tops.text]).size).toBe(1);
  });
});

test.describe('paginated note spacing (issue #1130)', () => {
  test.beforeEach(async ({ page }) => readyPage(page));

  test('spacing after separates notes but is not reserved below the last one', async ({ page }) => {
    // FootnoteText's common 6pt after. LibreOffice sets the gap between notes and puts the last
    // note line on the bottom margin; a trailing 6pt lifted every page's notes off it.
    const docx = generateFootnoteDocx(2, 3, 1, 0, {}, '<w:spacing w:after="120"/>');
    const band = await page.evaluate((bytes: number[]) => {
      const D = (window as any).Docxodus;
      const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Document', 'docx-', false, '', -1, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', true, false, false, false, false, false, null, false,
      );
      const host = document.createElement('div');
      host.innerHTML = html;
      document.body.appendChild(host);
      const container = host.querySelector<HTMLElement>('#pagination-container')!;
      new (window as any).DocxodusPagination.PaginationEngine(
        host.querySelector<HTMLElement>('#pagination-staging')!, container, { showPageNumbers: false },
      ).paginate();
      const footnotes = container.querySelector<HTMLElement>('.page-footnotes')!;
      const textBox = (element: Element) => {
        const range = document.createRange();
        range.selectNodeContents(element);
        const rects = Array.from(range.getClientRects()).filter((rect) => rect.height > 0);
        return { top: Math.min(...rects.map((r) => r.top)), bottom: Math.max(...rects.map((r) => r.bottom)) };
      };
      const notes = Array.from(footnotes.querySelectorAll('.footnote-item p')).map(textBox);
      const px = 4 / 3;
      return {
        count: notes.length,
        gapPt: (notes[1].top - notes[0].bottom) / px,
        slackPt: (footnotes.getBoundingClientRect().bottom - notes[notes.length - 1].bottom) / px,
      };
    }, Array.from(docx));
    expect(band.count).toBe(2);
    expect(band.gapPt).toBeGreaterThan(5);
    expect(band.slackPt).toBeLessThan(2);
  });
});

test.describe('paginated editor note number (issue #1130)', () => {
  test('the number inside a note paragraph is chrome, never committed as text', async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
    const out = await page.evaluate(async (bytes: number[]) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, { paginated: true });
      const note = container.querySelector<HTMLElement>(
        '.page-footnotes [data-anchor][contenteditable="true"]');
      if (!note) return { err: 'paginated note paragraph is not an editable block' };
      const number = note.querySelector<HTMLElement>('.footnote-number');
      const baseline = note.dataset.committedText;
      // Retype the note's text, the human gesture.
      const walker = document.createTreeWalker(note, NodeFilter.SHOW_TEXT);
      let target: Text | null = null;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if ((n as Text).data.includes('Footnote 1')) { target = n as Text; break; }
      }
      if (!target) return { err: 'note text node not found' };
      note.focus();
      const range = document.createRange();
      range.setStart(target, 0);
      range.setEnd(target, target.data.length);
      const selection = window.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      document.execCommand('insertText', false, 'Edited note.');
      note.dispatchEvent(new Event('blur'));
      await new Promise((resolve) => setTimeout(resolve, 300));
      // The repainted note still opens with its number, once.
      const numbersAfterEdit = Array.from(container.querySelectorAll<HTMLElement>(
        '.page-footnotes .footnote-number')).map((n) => n.textContent);
      const noteAfterEdit = container.querySelector<HTMLElement>(
        '.page-footnotes [data-anchor][contenteditable="true"]');
      const numberOpensNote = noteAfterEdit?.querySelector('.footnote-number')
        === noteAfterEdit?.firstElementChild;
      const projection = JSON.parse(D.DocxSessionBridge.Project(editor.sessionHandle));
      // Undo and redo re-render the note through the editor's other repaint paths.
      const numbersAt = () => {
        const block = container.querySelector<HTMLElement>(
          '.page-footnotes [data-anchor][contenteditable="true"]');
        return {
          numbers: Array.from(container.querySelectorAll<HTMLElement>('.page-footnotes .footnote-number'))
            .map((n) => n.textContent),
          opens: !!block && block.querySelector('.footnote-number') === block.firstElementChild,
          text: block?.textContent ?? '',
        };
      };
      editor.undo();
      await new Promise((resolve) => setTimeout(resolve, 300));
      const afterUndo = numbersAt();
      editor.redo();
      await new Promise((resolve) => setTimeout(resolve, 300));
      const afterRedo = numbersAt();
      const notes = Object.entries(projection.anchorIndex as Record<string, any>)
        .filter(([id]) => id.startsWith('p:fn:'))
        .map(([, target]) => target.textPreview as string);
      editor.close();
      container.remove();
      return {
        numberInside: !!number,
        numberEditable: number?.getAttribute('contenteditable') ?? null,
        baseline,
        notes,
        numbersAfterEdit,
        numberOpensNote,
        afterUndo,
        afterRedo,
      };
    }, Array.from(generateFootnoteDocx(1, 3, 1)));
    expect(out.err, out.err).toBeUndefined();
    expect(out.numberInside).toBe(true);
    expect(out.numberEditable).toBe('false');
    expect(out.baseline?.trim()).toBe('Footnote 1 paragraph 1 text.');
    expect(out.notes).toEqual(['Edited note.']);
    expect(out.numbersAfterEdit).toEqual(['1']);
    expect(out.numberOpensNote).toBe(true);
    expect(out.afterUndo).toMatchObject({ numbers: ['1'], opens: true });
    expect(out.afterUndo!.text).toContain('Footnote 1 paragraph 1 text.');
    expect(out.afterRedo).toMatchObject({ numbers: ['1'], opens: true });
    expect(out.afterRedo!.text).toContain('Edited note.');
  });
});
