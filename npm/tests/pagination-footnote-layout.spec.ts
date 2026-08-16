import { test, expect, type Page, type TestInfo } from '@playwright/test';
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

const LONG_NOTE_WORDS = Array.from({ length: 120 }, (_, index) => `word-${index}`).join(' ');

/**
 * One ordinary long paragraph is deliberately wider than several 60%-height note bands. The
 * trailing paragraphs make the former C2 failure visible: an oversized leader must not evacuate
 * the citing page or strand otherwise splittable siblings inside one clipped continuation.
 */
function longParagraphFootnoteHtml(
  trailingParagraphs = '',
  leadingContent = LONG_NOTE_WORDS,
  extraCss = '',
): string {
  return `
    <style>
      #staging { font: 12pt/12pt Arial; }
      .body { height: 20pt; margin: 0; }
      .page-footnotes { font-size: 10pt; line-height: 10pt; }
      .page-footnotes hr { height: 1px; margin: 0 0 3pt; border: 0; }
      .footnote-item { margin: 0; }
      .footnote-number { display: inline; vertical-align: super; }
      .footnote-content { display: inline; }
      .footnote-content p:first-of-type { display: inline; }
      .footnote-content p:not(:first-of-type) { display: block; margin: 0; }
      .footnote-continuation p { margin: 0; }
      ${extraCss}
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f-long"
             data-source-anchor-id="fn:fn:f-long">
          <span class="footnote-number">1</span>
          <span class="footnote-content">
            <p id="source-long" data-anchor="long-unid"
               data-source-anchor-id="p:fn:long">${leadingContent}</p>
            ${trailingParagraphs}
          </span>
        </div>
      </div>
      <div data-section-index="0"
           data-page-width="122" data-page-height="122"
           data-content-width="120" data-content-height="120"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p class="body" data-source-anchor-id="p:body:citation">
          citing body <sup data-footnote-id="f-long">1</sup>
        </p>
      </div>
    </div>
    <div id="container"></div>`;
}

async function paginateLongFootnote(
  page: Page,
  trailingParagraphs = '',
  leadingContent = LONG_NOTE_WORDS,
  extraCss = '',
) {
  const expectedText = `${leadingContent.replace(/<[^>]+>/g, ' ')} ${
    trailingParagraphs.replace(/<[^>]+>/g, ' ')
  }`.replace(/\s+/g, ' ').trim();
  await page.setContent(longParagraphFootnoteHtml(
    trailingParagraphs,
    leadingContent,
    extraCss,
  ));
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ expectedText }: { expectedText: string }) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const pagination = new PaginationEngine(staging, container, {
      showPageNumbers: false,
      layoutToken: { documentVersion: 1, rendererFingerprint: 'issue-489-evidence' },
    }).paginate();
    const boxes = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    const noteBands = boxes.map((box) => box.querySelector<HTMLElement>('.page-footnotes'));
    const paragraphFragments = Array.from(
      container.querySelectorAll<HTMLElement>('[data-source-anchor-id="p:fn:long"]'),
    );
    const visibleNoteText = boxes.map((box) => {
      const clone = box.querySelector<HTMLElement>('.page-footnotes')?.cloneNode(true) as
        HTMLElement | undefined;
      clone?.querySelectorAll('.footnote-number').forEach((number) => number.remove());
      clone?.querySelectorAll('p').forEach((paragraph) => paragraph.after(' '));
      return clone?.textContent ?? '';
    }).join(' ').replace(/\s+/g, ' ').trim();
    const clipped = noteBands.flatMap((band, pageIndex) => {
      if (!band) return [];
      const bandRect = band.getBoundingClientRect();
      return Array.from(band.querySelectorAll<HTMLElement>('p'))
        .filter((paragraph) => paragraph.getBoundingClientRect().bottom > bandRect.bottom + 1)
        .map((paragraph) => ({
          page: pageIndex + 1,
          anchor: paragraph.dataset.sourceAnchorId,
          paragraphBottom: paragraph.getBoundingClientRect().bottom,
          bandBottom: bandRect.bottom,
        }));
    });
    const pageMapFragments = pagination.pageMap.fragments
      .filter((fragment: any) => fragment.anchorId === 'p:fn:long');
    return {
      totalPages: pagination.totalPages,
      notePages: noteBands.filter(Boolean).length,
      paragraphFragments: paragraphFragments.length,
      pageMapFragments: pageMapFragments.map((fragment: any) => ({
        fragmentId: fragment.fragmentId,
        fragmentIndex: fragment.fragmentIndex,
        pageNumber: fragment.pageNumber,
        story: fragment.story,
        geometry: fragment.geometry,
      })),
      bareIds: paragraphFragments.filter((paragraph) => paragraph.id === 'source-long').length,
      bareAnchors: paragraphFragments.filter((paragraph) =>
        paragraph.dataset.anchor === 'long-unid').length,
      visibleNoteText,
      expectedText,
      clipped,
      firstPageText: noteBands[0]?.textContent?.replace(/\s+/g, ' ').trim() ?? '',
      pagesText: noteBands.map((band) => band?.textContent?.replace(/\s+/g, ' ').trim() ?? ''),
    };
  }, { expectedText });
}

type LongFootnoteResult = Awaited<ReturnType<typeof paginateLongFootnote>>;

function describeFailure(error: unknown): { name?: string; message: string; stack?: string } {
  if (error instanceof Error) {
    return { name: error.name, message: error.message, stack: error.stack };
  }
  return { message: String(error) };
}

async function attachLongFootnoteEvidence(
  page: Page,
  testInfo: TestInfo,
  scenario: string,
  result: LongFootnoteResult | undefined,
  failure: ReturnType<typeof describeFailure> | undefined,
): Promise<void> {
  const captureWarnings: string[] = [];
  let html: Buffer | undefined;
  let screenshot: Buffer | undefined;
  try {
    html = Buffer.from(await page.content());
  } catch (error) {
    captureWarnings.push(`HTML capture failed: ${describeFailure(error).message}`);
  }
  try {
    const container = page.locator('#container');
    if (await container.count()) screenshot = await container.screenshot();
  } catch (error) {
    captureWarnings.push(`screenshot capture failed: ${describeFailure(error).message}`);
  }

  await testInfo.attach(`${scenario}.json`, {
    body: Buffer.from(JSON.stringify({ result, failure, captureWarnings }, null, 2)),
    contentType: 'application/json',
  });
  if (html) {
    await testInfo.attach(`${scenario}.html`, { body: html, contentType: 'text/html' });
  }
  if (screenshot) {
    await testInfo.attach(`${scenario}.png`, { body: screenshot, contentType: 'image/png' });
  }
}

async function withLongFootnoteEvidence(
  page: Page,
  testInfo: TestInfo,
  scenario: string,
  paginate: () => Promise<LongFootnoteResult>,
  verify: (result: LongFootnoteResult) => void | Promise<void>,
): Promise<void> {
  let result: LongFootnoteResult | undefined;
  let failure: ReturnType<typeof describeFailure> | undefined;
  try {
    result = await paginate();
    await verify(result);
  } catch (error) {
    failure = describeFailure(error);
    throw error;
  } finally {
    await attachLongFootnoteEvidence(page, testInfo, scenario, result, failure);
  }
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

async function paginate(page: Page, html: string) {
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

  test('splits one long footnote paragraph across as many note bands as needed', async ({
    page,
  }, testInfo) => {
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-long-footnote',
      () => paginateLongFootnote(page),
      (result) => {
        expect(result.totalPages).toBeGreaterThan(2);
        expect(result.notePages).toBe(result.totalPages);
        expect(result.paragraphFragments).toBeGreaterThan(2);
        expect(result.pageMapFragments.length).toBe(result.paragraphFragments);
        expect(new Set(result.pageMapFragments.map((fragment: any) => fragment.pageNumber)).size)
          .toBeGreaterThan(2);
        expect(result.pageMapFragments.map((fragment: any) => fragment.fragmentIndex))
          .toEqual(result.pageMapFragments.map((_: any, index: number) => index));
        expect(new Set(result.pageMapFragments.map((fragment: any) => fragment.fragmentId)).size)
          .toBe(result.pageMapFragments.length);
        expect(result.pageMapFragments.every((fragment: any) =>
          fragment.story === 'footnote'
          && fragment.geometry.x >= 0
          && fragment.geometry.y >= 0
          && fragment.geometry.width > 0
          && fragment.geometry.height > 0)).toBe(true);
        expect(result.bareIds).toBe(1);
        expect(result.bareAnchors).toBe(1);
        expect(result.visibleNoteText).toBe(LONG_NOTE_WORDS);
        expect(result.clipped).toEqual([]);
        expect(result.firstPageText).toContain('word-0');
      },
    );
  });

  test('keeps a long leading paragraph with its citation while later note paragraphs continue', async ({
    page,
  }, testInfo) => {
    const trailing = `
      <p data-source-anchor-id="p:fn:after-one">after-one</p>
      <p data-source-anchor-id="p:fn:after-two">after-two</p>`;
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-leading-paragraph-with-siblings',
      () => paginateLongFootnote(page, trailing),
      (result) => {
        expect(result.firstPageText).toContain('word-0');
        expect(result.pagesText.some((text: string) => text.includes('after-one'))).toBe(true);
        expect(result.pagesText.some((text: string) => text.includes('after-two'))).toBe(true);
        expect(result.visibleNoteText).toBe(`${LONG_NOTE_WORDS} after-one after-two`);
        expect(result.clipped).toEqual([]);
      },
    );
  });

  test('an indivisible oversized note paragraph does not swallow later siblings', async ({
    page,
  }, testInfo) => {
    const trailing = `
      <p data-source-anchor-id="p:fn:after-one">after-one</p>
      <p data-source-anchor-id="p:fn:after-two">after-two</p>`;
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-indivisible-leading-paragraph',
      () => paginateLongFootnote(
        page,
        trailing,
        '<span class="unsafe-note-box">oversized-unsafe</span>',
        '.unsafe-note-box { display: inline-block; height: 150pt; width: 10pt; }',
      ),
      (result) => {
        expect(result.totalPages).toBeGreaterThan(1);
        expect(result.firstPageText).toContain('oversized-unsafe');
        expect(result.firstPageText).not.toContain('after-one');
        expect(result.pagesText.slice(1).some((text: string) =>
          text.includes('after-one'))).toBe(true);
        expect(result.pagesText.slice(1).some((text: string) =>
          text.includes('after-two'))).toBe(true);
        expect(result.pageMapFragments).toHaveLength(1);
        expect(result.expectedText).toBe('oversized-unsafe after-one after-two');
      },
    );
  });
});
