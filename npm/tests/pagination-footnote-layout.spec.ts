import { test, expect } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Layout contract for the paginated footnote area. The engine bottom-anchors the note block and
 * lets it grow upward into body space, so two things have to hold that nothing pinned before:
 * the body must never be drawn underneath the notes, and reserving note space must not throw the
 * rest of the page away. Both were broken on a real 94-footnote document — one page rendered body
 * and note glyphs on top of each other, and the median page ran 57% full (93% with notes off).
 */

/** A page of `bodyBlocks` paragraphs, the first `citing` of which each cite a distinct footnote. */
function stagingHtml(bodyBlocks: number, citing: number, noteLines: number): string {
  const notes = Array.from({ length: citing }, (_, i) => `
    <div class="footnote-item" data-footnote-id="f${i}">
      <span class="footnote-number">${i + 1}. </span>
      <span class="footnote-content">${
        Array.from({ length: noteLines }, () => `<p>note ${i} line</p>`).join('')
      }</span>
    </div>`).join('');

  const body = Array.from({ length: bodyBlocks }, (_, i) =>
    i < citing
      ? `<p class="body">body ${i} <sup data-footnote-id="f${i}">${i + 1}</sup></p>`
      : `<p class="body">body ${i}</p>`).join('');

  return `
    <style>
      #staging { font: 16px/16px Arial; }
      .body { height: 20pt; margin: 0; }
      .footnote-item { margin: 0; }
      .footnote-content > p { height: 10pt; margin: 0; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">${notes}</div>
      <div data-section-index="0"
           data-page-width="312" data-page-height="312"
           data-content-width="300" data-content-height="300"
           data-margin-top="6" data-margin-right="6"
           data-margin-bottom="6" data-margin-left="6">
        ${body}
      </div>
    </div>
    <div id="container"></div>`;
}

/** The document's final body block cites several whole notes, forcing at least one to defer. */
function finalBlockDeferralHtml(): string {
  const notes = Array.from({ length: 4 }, (_, i) => `
    <div class="footnote-item" data-footnote-id="last${i}">
      <span class="footnote-number">${i + 1}. </span>
      <span class="footnote-content">${
        Array.from({ length: 10 }, () => `<p>last note ${i} line</p>`).join('')
      }</span>
    </div>`).join('');
  const citations = Array.from({ length: 4 }, (_, i) =>
    `<sup data-footnote-id="last${i}">${i + 1}</sup>`).join('');
  return `
    <style>
      #staging { font: 16px/16px Arial; }
      .body { height: 100pt; margin: 0; }
      .footnote-item { margin: 0; }
      .footnote-content > p { height: 10pt; margin: 0; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">${notes}</div>
      <div data-section-index="0"
           data-page-width="312" data-page-height="312"
           data-content-width="300" data-content-height="300"
           data-margin-top="6" data-margin-right="6"
           data-margin-bottom="6" data-margin-left="6">
        <p class="body">only and final body block ${citations}</p>
      </div>
    </div>
    <div id="container"></div>`;
}

/**
 * Geometry of every rendered page. `usedPct` is BODY + NOTES against the content box: notes
 * legitimately consume page height, so body extent alone understates how full a page is — the
 * measure that matters is how much of the page is doing any work at all.
 */
const MEASURE = () => {
  const container = document.getElementById('container') as HTMLElement;
  return Array.from(container.querySelectorAll('.page-box')).map((box, i) => {
    const content = box.querySelector('.page-content') as HTMLElement | null;
    const notes = box.querySelector('.page-footnotes') as HTMLElement | null;
    const kids = content ? Array.from(content.children) : [];
    const contentTop = content ? content.getBoundingClientRect().top : 0;
    const bodyBottom = kids.length
      ? kids[kids.length - 1].getBoundingClientRect().bottom
      : contentTop;
    const avail = content ? content.getBoundingClientRect().height : 0;
    const bodyPx = bodyBottom - contentTop;
    const notesPx = notes ? notes.getBoundingClientRect().height : 0;
    return {
      page: i + 1,
      overlapPx: notes ? Math.round(bodyBottom - notes.getBoundingClientRect().top) : 0,
      usedPct: avail ? Math.round(((bodyPx + notesPx) / avail) * 100) : 0,
      bodyPx: Math.round(bodyPx),
      notesPx: Math.round(notesPx),
      hasNotes: !!notes,
    };
  });
};

/** Every note id cited on a page must be rendered somewhere in the document. */
const MEASURE_NOTES = () => {
  const container = document.getElementById('container') as HTMLElement;
  const renderedItems = Array.from(container.querySelectorAll<HTMLElement>('.footnote-item'));
  const rendered = new Set(renderedItems.map((i) => i.getAttribute('data-footnote-id')));
  const cited = Array.from(
    new Set(Array.from(container.querySelectorAll('[data-footnote-id]'))
      .filter((e) => e.tagName === 'SUP' || e.tagName === 'A')
      .map((e) => e.getAttribute('data-footnote-id'))),
  );
  return {
    cited: cited.length,
    rendered: rendered.size,
    lost: cited.filter((id) => !rendered.has(id)),
    nested: container.querySelectorAll('.footnote-item .footnote-item').length,
    clipped: renderedItems.filter((item) => {
      const band = item.closest('.page-footnotes');
      if (!band) return true;
      const itemRect = item.getBoundingClientRect();
      const bandRect = band.getBoundingClientRect();
      return itemRect.top < bandRect.top - 1 || itemRect.bottom > bandRect.bottom + 1;
    }).map((item) => {
      const band = item.closest('.page-footnotes')!;
      const itemRect = item.getBoundingClientRect();
      const bandRect = band.getBoundingClientRect();
      return {
        id: item.dataset.footnoteId,
        itemTop: Math.round(itemRect.top),
        itemBottom: Math.round(itemRect.bottom),
        bandTop: Math.round(bandRect.top),
        bandBottom: Math.round(bandRect.bottom),
      };
    }),
  };
};

async function paginate(page: any, html: string) {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(
    ({ measure }: { measure: string }) => {
      const staging = document.getElementById('staging') as HTMLElement;
      const container = document.getElementById('container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
      // eslint-disable-next-line no-eval
      return (0, eval)(`(${measure})`)();
    },
    { measure: MEASURE.toString() },
  );
}

test.describe('Paginated footnote layout', () => {
  test('body text is never drawn underneath the footnote area', async ({ page }) => {
    // Many notes on one page is the shape that produced overlapping glyphs on the real document.
    const pages = await paginate(page, stagingHtml(14, 8, 3));

    const overlapping = pages.filter((p: any) => p.overlapPx > 1);
    expect(
      overlapping,
      `pages where body overlaps notes: ${JSON.stringify(overlapping)}`,
    ).toEqual([]);
  });

  test('reserving footnote space does not leave the rest of the page empty', async ({ page }) => {
    // One early citation reserves note space; the remaining body must still fill the page.
    const pages = await paginate(page, stagingHtml(40, 3, 4));

    // Every page but the last should be reasonably full, counting body AND notes.
    const starved = pages.slice(0, -1).filter((p: any) => p.usedPct < 80);
    expect(
      starved,
      `under-filled pages: ${JSON.stringify(starved)} of ${pages.length}`,
    ).toEqual([]);
  });

  test('no cited note is dropped when several will not fit on one page', async ({ page }) => {
    // The engine carried an unfitted note in a SINGLE continuation slot, so when two notes on the
    // same page both failed to fit, the second overwrote the first and that note vanished from the
    // document — four notes disappeared from a real 94-footnote file. Deferred notes now queue.
    await page.setContent(stagingHtml(12, 10, 10));
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const notes = await page.evaluate(
      ({ measure }: { measure: string }) => {
        const staging = document.getElementById('staging') as HTMLElement;
        const container = document.getElementById('container') as HTMLElement;
        const { PaginationEngine } = (window as any).DocxodusPagination;
        new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
        // eslint-disable-next-line no-eval
        return (0, eval)(`(${measure})`)();
      },
      { measure: MEASURE_NOTES.toString() },
    );

    expect(notes.cited).toBeGreaterThan(0);
    expect(notes.lost, `notes cited but never rendered: ${JSON.stringify(notes.lost)}`).toEqual([]);
    // A note must never be wrapped inside another note (breaks the number/text line).
    expect(notes.nested).toBe(0);
  });

  test('a page whose notes are capped still fills its remaining body space', async ({ page }) => {
    // Notes long enough to hit the 60% cap: the other 40% must still take body text.
    const pages = await paginate(page, stagingHtml(30, 6, 8));

    const withNotes = pages.filter((p: any) => p.hasNotes);
    expect(withNotes.length).toBeGreaterThan(0);
    // No page may sit mostly idle while body blocks are still pending.
    const wasteful = pages.slice(0, -1).filter((p: any) => p.usedPct < 80);
    expect(
      wasteful,
      `pages left nearly empty: ${JSON.stringify(wasteful)} of ${pages.length}`,
    ).toEqual([]);
  });

  test('whole notes deferred from the final body block drain onto note-only pages', async ({ page }) => {
    await page.setContent(finalBlockDeferralHtml());
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const result = await page.evaluate(
      ({ measure }: { measure: string }) => {
        const staging = document.getElementById('staging') as HTMLElement;
        const container = document.getElementById('container') as HTMLElement;
        const { PaginationEngine } = (window as any).DocxodusPagination;
        new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
        // eslint-disable-next-line no-eval
        const notes = (0, eval)(`(${measure})`)();
        return {
          ...notes,
          noteOnlyPages: Array.from(container.querySelectorAll('.page-box')).filter((box) =>
            box.querySelector('.page-footnotes') && !box.querySelector('.page-content')?.children.length,
          ).length,
        };
      },
      { measure: MEASURE_NOTES.toString() },
    );
    expect(result.cited).toBe(4);
    expect(result.lost, `notes cited but never rendered: ${JSON.stringify(result.lost)}`).toEqual([]);
    expect(result.clipped, `notes rendered outside their visible band: ${JSON.stringify(result.clipped)}`)
      .toEqual([]);
    expect(result.noteOnlyPages).toBeGreaterThan(0);
  });
});
