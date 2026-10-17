import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

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
  /** Height of each page's note area, in points (0 without one). */
  band: number[];
  /** The note text each page shows. */
  notes: string[];
}

async function paginate(page: Page, html: string): Promise<Layout> {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(() => {
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(document.getElementById('staging'), document.getElementById('container'), {
      scale: 1, showPageNumbers: false,
    }).paginate();
    const boxes = Array.from(document.querySelectorAll<HTMLElement>('#container .page-box'));
    return {
      body: boxes.map((box) => box.querySelector('.page-content')!.children.length),
      band: boxes.map((box) => {
        const notes = box.querySelector<HTMLElement>('.page-footnotes');
        return notes ? notes.getBoundingClientRect().height * 0.75 : 0;
      }),
      notes: boxes.map((box) => (box.querySelector('.page-footnotes')?.textContent || '').replace(/\s+/g, ' ').trim()),
    };
  });
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
    // The note's middle paragraph is one indivisible 400pt box. Its page clips it, the note's last
    // paragraph follows on the next page, and pagination ends.
    const layout = await paginate(page, staging(
      `<p>cites${reference('1')}</p>${bodyLines(1, 5)}`,
      note('1', ['first part', '<span class="tall"></span> tall part', 'last part'])));

    const text = layout.notes.join(' ');
    for (const part of ['first part', 'tall', 'last part']) expect(text).toContain(part);
    expect(layout.body.reduce((sum, count) => sum + count, 0)).toBe(6);
    expect(layout.body.length).toBeLessThan(6);
  });
});
