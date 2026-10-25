import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { generateFootnoteDocx, generateNormalKeptNoteDocx } from './docx-footnote-fixture.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #1178: how much of a page paginated footnotes may take. Word reference PDFs hold back no
 * room for body text: a note takes everything below the line that cites it, and a carried note
 * that does not end on a page takes that page alone, with body text resuming on the page where the
 * note ends (docs/ooxml_corner_cases.md). The paginator used to cap the note area at 60% of the
 * page while body text remained.
 *
 * Every page here is 300pt of body at 15pt lines.
 */

const LINE_PT = 15;

function lines(prefix: string, count: number): string {
  return Array.from({ length: count }, (_, index) => `${prefix} ${index + 1}`).join('<br>');
}

function bodyLines(from: number, count: number): string {
  return Array.from({ length: count }, (_, index) => `<p>body ${from + index}</p>`).join('');
}

function staging(blocks: string, notes: string): string {
  return `
    <style>
      body { font: 12pt/${LINE_PT}pt serif; }
      p { margin: 0; padding: 0; font: inherit; }
      .tall { display: inline-block; width: 10pt; height: 400pt; }
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

function note(id: string, paragraphs: string[]): string {
  return `
    <div class="footnote-item" data-footnote-id="${id}">
      <span class="footnote-number">${id}</span>
      <div class="footnote-content">${paragraphs.map((text) => `<p>${text}</p>`).join('')}</div>
    </div>`;
}

const reference = (id: string) => `<a class="footnote-ref" data-footnote-id="${id}"><sup>${id}</sup></a>`;

interface Layout {
  /** Body paragraphs on each page. */
  body: number[];
  /** Whether every body paragraph on each page lies inside that page's visible body area. */
  bodyVisible: boolean[];
  /** Height of each page's note area, in points (0 without one). */
  band: number[];
  /** The note text each page shows. */
  notes: string[];
  /** Every nonempty text range lies inside its painted footnote band. */
  notesVisible: boolean[];
  citationPage: number;
}

async function paginate(page: Page, html: string | Uint8Array, fragmentParagraphs = true, embedHtml = false): Promise<Layout> {
  if (typeof html !== 'string') {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    html = await page.evaluate((input): string => (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-', 1, 1, 'page-', false, 0, 'annot-',
      true, false, false, false, false, false, null, false), Array.from(html));
  }
  await page.setContent(embedHtml ? '<div id="viewer"></div>' : html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ fragmentParagraphs, embeddedHtml }) => {
    const { PaginationEngine, paginateHtml } = (window as any).DocxodusPagination;
    const options = { scale: 1, showPageNumbers: false, fragmentParagraphs };
    if (embeddedHtml) paginateHtml(embeddedHtml, document.querySelector('#viewer'), options);
    else new PaginationEngine(document.querySelector('#staging, #pagination-staging'),
      document.querySelector('#container, #pagination-container'), options).paginate();
    const boxes = Array.from(document.querySelectorAll<HTMLElement>('#container .page-box, #pagination-container .page-box'));
    return {
      citationPage: boxes.findIndex(box => box.querySelector('.page-content [data-footnote-id]')) + 1,
      body: boxes.map((box) => box.querySelector('.page-content')!.children.length),
      bodyVisible: boxes.map((box) => {
        const content = box.querySelector<HTMLElement>('.page-content')!;
        const area = content.getBoundingClientRect();
        return Array.from(content.children).every((block) => {
          const rect = block.getBoundingClientRect();
          return rect.top >= area.top - 0.5 && rect.bottom <= area.bottom + 0.5;
        });
      }),
      band: boxes.map((box) => {
        const notes = box.querySelector<HTMLElement>('.page-footnotes');
        return notes ? notes.getBoundingClientRect().height * 0.75 : 0;
      }),
      notes: boxes.map((box) => (box.querySelector('.page-footnotes')?.textContent || '').replace(/\s+/g, ' ').trim()),
      notesVisible: boxes.map((box) => {
        const band = box.querySelector<HTMLElement>('.page-footnotes');
        if (!band) return true;
        const area = band.getBoundingClientRect();
        const walker = document.createTreeWalker(band, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (!node.textContent?.trim()) continue;
          const range = document.createRange(); range.selectNodeContents(node);
          if (Array.from(range.getClientRects()).some(rect => rect.width > 0 && rect.height > 0
            && (rect.top < area.top - 0.5 || rect.bottom > area.bottom + 0.5))) return false;
        }
        return true;
      }),
    };
  }, { fragmentParagraphs, embeddedHtml: embedHtml ? html : null });
}

test.describe('Paginated note area', () => {
  test('a note takes everything below the line that cites it', async ({ page }) => {
    // A 30-line note cited by the first line, then 30 lines of body: Word keeps only the citing
    // line on page 1 and gives the note the rest of the page.
    const layout = await paginate(page, staging(
      `<p>cites${reference('1')}</p>${bodyLines(1, 30)}`, note('1', [lines('note line', 30)])));

    expect(layout.body[0]).toBe(1);
    // All but the citing line and the separator: far beyond the 60% the paginator used to allow.
    expect(layout.band[0]).toBeGreaterThan(300 * 0.85);
  });

  test('a carried note that does not end on a page holds that page alone', async ({ page }) => {
    // A 50-line note: page 2 cannot hold its remaining 30-odd lines beside anything, so it holds
    // only the note, and the body resumes on page 3, where the note ends.
    const layout = await paginate(page, staging(
      `<p>cites${reference('1')}</p>${bodyLines(1, 30)}`, note('1', [lines('note line', 50)])));

    expect(layout.body[0]).toBe(1);
    expect(layout.body[1]).toBe(0);
    expect(layout.notes[1]).toContain('note line 30');
    expect(layout.body[2]).toBeGreaterThan(0);
    expect(layout.notes[2]).toContain('note line 50');
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(31);
  });

  test('a carried part taller than any page still advances page by page', async ({ page }) => {
    // A liveness guard: the note's middle paragraph is one indivisible 400pt box. Its page clips
    // it, the note's last paragraph follows on the next page, and pagination ends.
    const layout = await paginate(page, staging(
      `<p>cites${reference('1')}</p>${bodyLines(1, 5)}`,
      note('1', ['first part', '<span class="tall"></span> tall part', 'last part'])));

    const text = layout.notes.join(' ');
    for (const part of ['first part', 'tall', 'last part']) expect(text).toContain(part);
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(6);
    expect(layout.body.length).toBeLessThan(6);
  });

  for (const fragmentParagraphs of [true, false]) {
    test(`a note part taller than any page never covers the body placed above it (fragmentation ${fragmentParagraphs ? 'on' : 'off'})`, async ({ page }) => {
      // Five body lines, then a line citing a note whose only part is an indivisible 400pt box.
      // The band is clipped to what the body leaves; every body line placed on the page shows.
      const layout = await paginate(page, staging(
        `${bodyLines(1, 5)}<p>cites${reference('1')}</p>${bodyLines(6, 5)}`,
        note('1', ['<span class="tall"></span>'])), fragmentParagraphs);

      expect(layout.bodyVisible.every(Boolean)).toBe(true);
      expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(11);
    });
  }

  test('a carried note of kept paragraphs that runs past the next page holds those pages alone', async ({ page }) => {
    // Six-line paragraphs that each keep their lines together, as in the Word reference PDF of
    // this shape: pages whose carried note does not end on them hold no body (issue #1178).
    const parts = Array.from({ length: 10 }, (_, index) =>
      `<span style="display:inline-block">${lines(`part ${index + 1} line`, 6)}</span>`);
    const layout = await paginate(page, staging(
      `<p>cites${reference('1')}</p>${bodyLines(1, 10)}`, note('1', parts)));

    // Every page after the first whose note runs on to the next page holds no body.
    const runsOn = layout.notes.map((text, index) => index > 0 && text !== '' && (layout.notes[index + 1] ?? '') !== '');
    expect(runsOn.filter(Boolean).length).toBeGreaterThanOrEqual(2);
    layout.body.forEach((count, index) => { if (runsOn[index]) expect(count, `page ${index + 1}`).toBe(0); });
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(11);
  });
});

for (const lineCount of [10, 30]) {
  test(`editor mode ${lineCount}-line kept note ${lineCount > 20 ? 'defers oversized paragraph' : 'fits beside citation'}`, async ({ page }) => {
    const layout = await paginate(page, staging(
      `${bodyLines(1, 5)}<p>cites${reference('1')}</p>${bodyLines(6, 5)}`,
      note('1', [lines('kept note', lineCount)]).replace('<p>', '<p data-keep-lines="true">')), false);
    expect(layout.bodyVisible.every(Boolean)).toBe(true);
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(11);
    if (lineCount > 20) {
      expect(layout.notes[0]).toBe('');
      expect(layout.notes[1]).toContain('kept note 1');
      expect(layout.body[1]).toBe(0);
    } else {
      expect(layout.notes[0]).toContain('kept note 1');
    }
  });
}

for (const precedingNote of [false, true]) {
  test(`editor oversized note advances alone ${precedingNote ? 'after an earlier note' : 'when cited first'}`, async ({ page }) => {
    const layout = await paginate(page, staging(
      `${precedingNote ? `<p>earlier${reference('1')}</p>` : ''}<p>cites${reference('2')}</p>${bodyLines(1, 5)}`,
      (precedingNote ? note('1', ['earlier definition']) : '')
        + note('2', [lines('oversized line', 30)]).replace('<p>', '<p data-keep-lines="true">')), false);
    const ownPage = layout.notes.findIndex(text => text.includes('oversized line 1'));
    expect(ownPage).toBeGreaterThan(0);
    expect(layout.body[ownPage]).toBe(0);
    expect(layout.bodyVisible.every(Boolean)).toBe(true);
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(precedingNote ? 7 : 6);
    expect(layout.notes.join(' ').match(/oversized line \d+/g)).toEqual(
      Array.from({ length: 30 }, (_, index) => `oversized line ${index + 1}`));
    if (precedingNote) expect(layout.notes[0]).toContain('earlier definition');
    expect(layout.body.length).toBeLessThan(5);
  });
}


for (const [name, words, firstCitation] of [
  ['mid-page oversized', 1200, 16], ['first-line oversized', 1200, 1], ['mid-page fitting', 60, 16],
] as const) {
  test(`editor DOCX ${name} preserves note text and deferral placement`, async ({ page }) => {
    const layout = await paginate(page,
      generateFootnoteDocx(1, 30, 1, words, {}, '<w:keepLines/>', firstCitation), false);
    const notePage = layout.notes.findIndex(text => text.includes('footnote-1-1-1')) + 1;
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(30);
    expect(layout.bodyVisible.every(Boolean)).toBe(true);
    expect(layout.notes.join(' ').match(/footnote-1-1-\d+/g)).toEqual(
      Array.from({ length: words }, (_, i) => `footnote-1-1-${i + 1}`));
    if (words > 60) {
      expect(notePage).toBeGreaterThan(layout.citationPage);
      expect(layout.body[notePage - 1]).toBe(0);
    } else expect(notePage).toBe(layout.citationPage);
    expect(layout.body.length).toBeLessThan(5);
  });
}

for (const firstCitation of [1, 16]) {
  test(`read-only DOCX oversized kept note: citation line ${firstCitation}`, async ({ page }) => {
    const layout = await paginate(page,
      generateFootnoteDocx(1, 30, 1, 1200, {}, '<w:keepLines/>', firstCitation,
        { compatibilityMode: 15, fontFamily: 'Times New Roman' }));
    const firstNotePage = layout.notes.findIndex(text => text.includes('footnote-1-1-1')) + 1;
    // Native Word retries a mid-page citation once, then leaves the whole kept note
    // until a note-only page. Following body stays with the citation, ahead of the note.
    expect(layout.citationPage).toBe(firstCitation === 1 ? 1 : 2);
    expect(firstNotePage).toBe(layout.citationPage + 1);
    expect(layout.body.slice(0, firstNotePage - 1)).toEqual(firstCitation === 1 ? [30] : [15, 15]);
    expect(layout.body.slice(firstNotePage - 1).every(count => count === 0)).toBe(true);
    expect(layout.bodyVisible.every(Boolean)).toBe(true);
    expect(layout.notesVisible.every(Boolean)).toBe(true);
    expect(layout.notes.join(' ').match(/footnote-1-1-\d+/g)).toEqual(
      Array.from({ length: 1200 }, (_, i) => `footnote-1-1-${i + 1}`));
    expect(layout.body.length).toBeLessThan(9);
  });
}

test('read-only DOCX fitting kept note stays beside its mid-page citation', async ({ page }) => {
  const layout = await paginate(page,
    generateFootnoteDocx(1, 30, 1, 60, {}, '<w:keepLines/>', 16,
      { compatibilityMode: 15, fontFamily: 'Times New Roman' }));
  expect(layout.citationPage).toBe(1);
  expect(layout.body).toEqual([30]);
  expect(layout.notes[0]).toContain('footnote-1-1-60');
  expect(layout.notesVisible).toEqual([true]);
});

for (const [count, widow, firstPageLines] of [
  [18, undefined, 18], [19, undefined, 17], [19, false, 18], [19, true, 17],
  [20, undefined, 18], [20, false, 18], [20, true, 18], [21, undefined, 18],
] as const) {
  test(`Word normal kept note ${count} lines, widow ${widow ?? 'default'} starts beside citation`, async ({ page }) => {
    const layout = await paginate(page, generateNormalKeptNoteDocx(count, widow));
    expect(layout.citationPage).toBe(1);
    expect(layout.body).toEqual([1, 1]);
    expect(layout.notesVisible.every(Boolean)).toBe(true);
    expect(layout.notes.map(text => text.match(/SpruceNoteLine\d{2}/g) ?? [])).toEqual([
      Array.from({ length: firstPageLines }, (_, i) => `SpruceNoteLine${String(i + 1).padStart(2, '0')}`),
      Array.from({ length: count - firstPageLines }, (_, i) => `SpruceNoteLine${String(i + firstPageLines + 1).padStart(2, '0')}`),
    ]);
  });
}

for (const mode of [14, undefined, 15]) {
  test(`Word compatibility ${mode ?? 'absent'} controls oversized kept-note deferral`, async ({ page }) => {
    const layout = await paginate(page, generateNormalKeptNoteDocx(21, undefined, mode), true, true);
    expect(layout.citationPage).toBe(1);
    expect(layout.body).toEqual(mode === 15 ? [2, 0, 0] : [1, 1]);
    expect(layout.notesVisible.every(Boolean)).toBe(true);
    const markers = layout.notes.map(text => text.match(/SpruceNoteLine\d{2}/g) ?? []);
    expect(markers.map(m => m.length)).toEqual(mode === 15 ? [0, 19, 2] : [18, 3]);
    expect(markers.flat()).toEqual(Array.from({ length: 21 }, (_, i) => `SpruceNoteLine${String(i + 1).padStart(2, '0')}`));
  });
}

for (const compatibilityMode of [14, undefined]) {
  test(`Word original long kept note compatibility ${compatibilityMode ?? 'absent'} begins beside citation`, async ({ page }) => {
    const layout = await paginate(page, generateFootnoteDocx(1, 30, 1, 1200, {}, '<w:keepLines/>', 1,
      { compatibilityMode, fontFamily: 'Times New Roman' }));
    expect(layout.citationPage).toBe(1);
    expect(layout.notes[0]).toContain('footnote-1-1-1');
    expect(layout.body[0]).toBe(1);
    expect(layout.notesVisible.every(Boolean)).toBe(true);
    expect(layout.bodyVisible.every(Boolean)).toBe(true);
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(30);
    expect(layout.notes.join(' ').match(/footnote-1-1-\d+/g)).toEqual(
      Array.from({ length: 1200 }, (_, i) => `footnote-1-1-${i + 1}`));
  });
}
