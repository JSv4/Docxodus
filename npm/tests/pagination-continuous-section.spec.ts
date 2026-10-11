import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

/**
 * Continuous section breaks and w:cols column geometry (issue #413).
 *
 * A `w:type="continuous"` section keeps filling the page its predecessor started;
 * the paginator used to open a fresh page for EVERY section wrapper, promoting the
 * continuous break to a page break. A `w:cols` section uses multiple columns;
 * a final continuous section fills their available height before advancing.
 */

const PAGE_GEOMETRY = `
  data-page-width="102" data-page-height="102"
  data-content-width="100" data-content-height="100"
  data-margin-top="1" data-margin-right="1"
  data-margin-bottom="1" data-margin-left="1"`;

const SHARED_PAGE_GEOMETRY = `
  data-page-width="102" data-page-height="302"
  data-content-width="100" data-content-height="300"
  data-margin-top="1" data-margin-right="1"
  data-margin-bottom="1" data-margin-left="1"
  data-header-height="1" data-footer-height="1"`;

const SPILL_SECTION_GEOMETRY = `
  data-page-width="102" data-page-height="302"
  data-content-width="100" data-content-height="300"
  data-margin-top="1" data-margin-right="1"
  data-margin-bottom="1" data-margin-left="1"
  data-header-height="25" data-footer-height="17"`;

function staging(sections: string, stagingAttributes = ''): string {
  return `
    <style>
      #staging { font: 12px/12px Arial; }
      #staging p { box-sizing: border-box; margin: 0; padding: 0; width: 100%; }
    </style>
    <div id="staging" ${stagingAttributes}>${sections}</div>
    <div id="container"></div>`;
}

interface PaginatedShape {
  /** Per page, the text of each top-level block. */
  content: string[][];
  /** Per page, each top-level block's inline column-count ('' when none). */
  columnCounts: string[][];
  /** Section that owns each physical page's running stories and numbering. */
  sectionIndices: number[];
  /** One-based page number within the owning section. */
  pagesInSection: number[];
  headers: string[];
  footers: string[];
  footnotes: string[];
  headerTops: number[];
  footerBottoms: number[];
  displayedPageNumbers: number[];
  sectionFillers: boolean[];
  pageSizes: Array<{ width: number; height: number }>;
  contentBoxes: Array<{ top: number; left: number; width: number; height: number }>;
  checkpoints: number;
}

async function paginate(page: Page, checkpointLimit?: number): Promise<PaginatedShape> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate((limit) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    let checkpoints = 0;
    new PaginationEngine(staging, container, {
      showPageNumbers: false,
      checkCancellation: () => {
        checkpoints++;
        if (limit !== undefined && checkpoints > limit) {
          throw new Error(`pagination exceeded ${limit} checkpoints`);
        }
      },
    }).paginate();

    const pages = Array.from(container.querySelectorAll<HTMLElement>('.page-content'));
    const boxes = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    return {
      content: pages.map(content =>
        Array.from(content.children).map(block => (block.textContent || '').trim())
      ),
      columnCounts: pages.map(content =>
        Array.from(content.children).map(
          block => (block as HTMLElement).style.columnCount || ''
        )
      ),
      sectionIndices: boxes.map(box => Number(box.dataset.sectionIndex)),
      pagesInSection: boxes.map(box => Number(box.dataset.pageInSection)),
      headers: boxes.map(box =>
        (box.querySelector<HTMLElement>('.page-header')?.innerText || '').trim()),
      footers: boxes.map(box =>
        (box.querySelector<HTMLElement>('.page-footer')?.innerText || '').trim()),
      footnotes: boxes.map(box =>
        (box.querySelector<HTMLElement>('.page-footnotes')?.innerText || '')
          .replace(/\s+/g, ' ')
          .trim()),
      headerTops: boxes.map(box =>
        parseFloat(box.querySelector<HTMLElement>('.page-header')?.style.top || 'NaN')),
      footerBottoms: boxes.map(box =>
        parseFloat(box.querySelector<HTMLElement>('.page-footer')?.style.bottom || 'NaN')),
      displayedPageNumbers: boxes.map(box => Number(box.dataset.displayedPageNumber)),
      sectionFillers: boxes.map(box => box.dataset.sectionFiller === 'true'),
      pageSizes: boxes.map(box => ({
        width: parseFloat(box.style.width),
        height: parseFloat(box.style.height),
      })),
      contentBoxes: pages.map(content => ({
        top: parseFloat(content.style.top),
        left: parseFloat(content.style.left),
        width: parseFloat(content.style.width),
        height: parseFloat(content.style.height),
      })),
      checkpoints,
    };
  }, checkpointLimit);
}

test.describe('Continuous section breaks', () => {
  test('preserves zero margins and distances instead of replacing them with defaults', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0"
           data-page-width="100" data-page-height="100"
           data-content-width="100" data-content-height="100"
           data-margin-top="0" data-margin-right="0"
           data-margin-bottom="0" data-margin-left="0"
           data-header-height="0" data-footer-height="0">
        <p style="height: 20pt">edge to edge</p>
      </div>`));

    const result = await paginate(page);

    expect(result.pageSizes).toEqual([{ width: 100, height: 100 }]);
    expect(result.contentBoxes).toEqual([{ top: 0, left: 0, width: 100, height: 100 }]);
  });

  test('falls back deterministically from partial and non-finite page geometry', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0"
           data-page-width="612evil" data-page-height="Infinity"
           data-content-width="NaN" data-content-height="-1"
           data-margin-left="0">
        <p style="height: 20pt">bounded fallback</p>
      </div>`));

    const result = await paginate(page);

    expect(result.pageSizes).toEqual([{ width: 612, height: 792 }]);
    // Missing margins retain their defaults, explicit zero remains zero, and invalid content
    // dimensions are derived from the effective page/margin box.
    expect(result.contentBoxes).toEqual([{ top: 72, left: 0, width: 540, height: 648 }]);
  });

  test('a continuous section continues on the current page', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 30pt">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous" ${PAGE_GEOMETRY}>
        <p style="height: 30pt">body</p>
      </div>`));

    const { content } = await paginate(page);

    expect(content).toEqual([['title', 'body']]);
  });

  test('a default (nextPage) section still starts a fresh page', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 30pt">title</p>
      </div>
      <div data-section-index="1" ${PAGE_GEOMETRY}>
        <p style="height: 30pt">body</p>
      </div>`));

    const { content } = await paginate(page);

    expect(content).toEqual([['title'], ['body']]);
  });

  test('a continuous section with a different page box starts a fresh page, as Word does', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 30pt">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-page-width="204" data-page-height="204"
           data-content-width="200" data-content-height="200"
           data-margin-top="2" data-margin-right="2"
           data-margin-bottom="2" data-margin-left="2">
        <p style="height: 30pt">body</p>
      </div>`));

    const { content } = await paginate(page);

    expect(content).toEqual([['title'], ['body']]);
  });

  test('an overfilled shared page still turns normally', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous" ${PAGE_GEOMETRY}>
        <p style="height: 30pt">first</p>
        <p style="height: 30pt">second</p>
      </div>`));

    const { content } = await paginate(page);

    expect(content).toEqual([['title', 'first'], ['second']]);
  });

  test('an overflow page belongs to the continuous section that supplies its content', async ({ page }) => {
    await page.setContent(staging(`
      <div id="pagination-hf-registry">
        <div data-section="0" data-hf-type="header-default"><p>old header</p></div>
        <div data-section="0" data-hf-type="footer-default"><p>old footer</p></div>
        <div data-section="1" data-hf-type="header-first"><p>new first header</p></div>
        <div data-section="1" data-hf-type="header-even"><p>new even header</p></div>
        <div data-section="1" data-hf-type="header-default"><p>new default header</p></div>
        <div data-section="1" data-hf-type="footer-first"><p>new first footer</p></div>
        <div data-section="1" data-hf-type="footer-even"><p>new even footer</p></div>
        <div data-section="1" data-hf-type="footer-default"><p>new default footer</p></div>
      </div>
      <div data-section-index="0" ${SHARED_PAGE_GEOMETRY}>
        <p style="height: 80pt">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous" ${SPILL_SECTION_GEOMETRY}>
        <p style="height: 60pt">first</p>
        <p style="height: 180pt">second</p>
      </div>`));

    const result = await paginate(page);

    expect(result.content).toEqual([['title', 'first'], ['second']]);
    expect(result.sectionIndices).toEqual([0, 1]);
    // The shared page is already page 1 of section 1 even though section 0 owns its stories.
    expect(result.pagesInSection).toEqual([1, 2]);
    expect(result.headers).toEqual(['old header', 'new even header']);
    expect(result.footers).toEqual(['old footer', 'new even footer']);
    expect(result.headerTops).toEqual([1, 25]);
    expect(result.footerBottoms).toEqual([1, 17]);
  });

  test('a pre-break footnote promotes a continuous section to a fresh page', async ({ page }) => {
    await page.setContent(staging(`
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f1">
          <span class="footnote-number">1</span>
          <span class="footnote-content">note before section break</span>
        </div>
      </div>
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">citing text <sup data-footnote-id="f1">1</sup></p>
      </div>
      <div data-section-index="1" data-section-type="continuous" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">later section</p>
      </div>`));

    const result = await paginate(page);

    expect(result.content).toEqual([['citing text 1'], ['later section']]);
    expect(result.sectionIndices).toEqual([0, 1]);
    expect(result.pagesInSection).toEqual([1, 1]);
    expect(result.footnotes[0]).toContain('note before section break');
    expect(result.footnotes[1]).toBe('');
  });

  test('long pre-break note tails advance section pages, stories, and restarted numbering', async ({
    page,
  }) => {
    const words = Array.from({ length: 80 }, (_, index) => `tail-${index}`).join(' ');
    await page.setContent(staging(`
      <style>
        .page-footnotes { font: 10pt/10pt Arial; }
        .page-footnotes hr { height: 1px; margin: 0 0 3pt; border: 0; }
        .footnote-item, .footnote-content p, .footnote-continuation p { margin: 0; }
      </style>
      <div id="pagination-hf-registry">
        <div data-section="1" data-hf-type="header-first"><p>section-one-first</p></div>
        <div data-section="1" data-hf-type="header-even"><p>section-one-even</p></div>
        <div data-section="1" data-hf-type="header-default"><p>section-one-default</p></div>
      </div>
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f-long">
          <span class="footnote-number">1</span>
          <span class="footnote-content"><p>${words}</p></span>
        </div>
      </div>
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height:20pt">before <sup data-footnote-id="f-long">1</sup></p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-page-num-start="10" ${PAGE_GEOMETRY}>
        <p style="height:20pt">after continuation</p>
      </div>`));

    const result = await paginate(page, 20_000);
    const ownedBySecondSection = result.sectionIndices
      .map((owner, index) => ({ owner, index }))
      .filter(({ owner }) => owner === 1)
      .map(({ index }) => index);

    expect(ownedBySecondSection.length).toBeGreaterThan(2);
    expect(ownedBySecondSection.map((index) => result.pagesInSection[index]))
      .toEqual(ownedBySecondSection.map((_, index) => index + 1));
    expect(ownedBySecondSection.map((index) => result.displayedPageNumbers[index]))
      .toEqual(ownedBySecondSection.map((_, index) => index + 10));
    // Even/odd stories alternate with the DISPLAYED number (10, 11, 12, …), not the
    // one-based position in section — see the issue #536 parity tests below.
    expect(ownedBySecondSection.map((index) => result.headers[index]))
      .toEqual(ownedBySecondSection.map((_, index) =>
        index === 0 ? 'section-one-first'
          : (index + 10) % 2 === 0 ? 'section-one-even' : 'section-one-default'));
    expect(result.content.flat()).toEqual(['before 1', 'after continuation']);
    const renderedWords = result.footnotes.join(' ');
    for (const word of words.split(' ')) {
      expect(renderedWords.match(new RegExp(`\\b${word}\\b`, 'g'))).toHaveLength(1);
    }
  });

  test('a tall first header advances one empty physical page without retrying forever', async ({
    page,
  }) => {
    await page.setContent(staging(`
      <div id="pagination-hf-registry">
        <div data-section="0" data-hf-type="header-first">
          <p style="height:80pt">tall first header</p>
        </div>
      </div>
      <div data-section-index="0"
           data-page-width="122" data-page-height="122"
           data-content-width="120" data-content-height="120"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1"
           data-header-height="1" data-footer-height="1">
        <p style="height:60pt">fits the default-header page</p>
      </div>`));

    const result = await paginate(page, 2_000);
    expect(result.content).toEqual([[], ['fits the default-header page']]);
    expect(result.pagesInSection).toEqual([1, 2]);
    expect(result.headers).toEqual(['tall first header', '']);
    expect(result.checkpoints).toBeLessThan(2_000);
  });

  test('footnoteLayoutLikeWW8 shares only post-break paragraphs without references', async ({
    page,
  }) => {
    await page.setContent(staging(`
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f1">
          <span class="footnote-content">note before section break</span>
        </div>
        <div class="footnote-item" data-footnote-id="f2">
          <span class="footnote-content">note after section break</span>
        </div>
      </div>
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">before <sup data-footnote-id="f1">1</sup></p>
      </div>
      <div data-section-index="1" data-section-type="continuous" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">compatible paragraph</p>
        <p style="height: 20pt">referenced paragraph <sup data-footnote-id="f2">2</sup></p>
      </div>`, 'data-footnote-layout-like-word8="true"'));

    const result = await paginate(page);

    expect(result.content).toEqual([
      ['before 1', 'compatible paragraph'],
      ['referenced paragraph 2'],
    ]);
    expect(result.sectionIndices).toEqual([0, 1]);
    expect(result.pagesInSection).toEqual([1, 2]);
    expect(result.footnotes[0]).toContain('note before section break');
    expect(result.footnotes[1]).toContain('note after section break');
  });

  test('odd/even section breaks insert blank parity pages with predecessor geometry', async ({
    page,
  }) => {
    await page.setContent(staging(`
      <div data-section-index="0" data-page-num-start="10" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">section zero</p>
      </div>
      <div data-section-index="1" data-section-type="oddPage"
           data-page-width="202" data-page-height="302"
           data-content-width="200" data-content-height="300"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p style="height: 20pt">section one</p>
      </div>
      <div data-section-index="2" data-section-type="evenPage"
           data-page-width="302" data-page-height="202"
           data-content-width="300" data-content-height="200"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p style="height: 20pt">section two</p>
      </div>`));

    const result = await paginate(page);

    expect(result.content).toEqual([
      ['section zero'],
      [],
      ['section one'],
      ['section two'],
    ]);
    expect(result.sectionIndices).toEqual([0, 0, 1, 2]);
    expect(result.pagesInSection).toEqual([1, 2, 1, 1]);
    expect(result.displayedPageNumbers).toEqual([10, 11, 12, 13]);
    expect(result.sectionFillers).toEqual([false, true, false, false]);
    expect(result.pageSizes).toEqual([
      { width: 102, height: 102 },
      // The filler is the back side of section 0, not the new section's paper stock.
      { width: 102, height: 102 },
      { width: 202, height: 302 },
      { width: 302, height: 202 },
    ]);
  });

  test('an even-page section inserts a filler when the next physical page is odd', async ({
    page,
  }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">first page</p>
        <p style="height: 60pt">second page</p>
      </div>
      <div data-section-index="1" data-section-type="evenPage" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">even section</p>
      </div>`));

    const result = await paginate(page);

    expect(result.content).toEqual([
      ['first page'],
      ['second page'],
      [],
      ['even section'],
    ]);
    expect(result.sectionFillers).toEqual([false, false, true, false]);
    expect(result.sectionIndices).toEqual([0, 0, 0, 1]);
    expect(result.pagesInSection).toEqual([1, 2, 3, 1]);
  });

  test('even/odd stories follow the page number, not the position in section', async ({
    page,
  }) => {
    // Issue #536: ECMA-376 §17.10.5 hangs the even story on "even numbered pages" — the
    // page NUMBER, which keeps counting across a section boundary unless w:pgNumType
    // restarts it. Word and LibreOffice both render DB001-Sections this way; selecting by
    // position-in-section flips every story in a section that begins on an even page.
    await page.setContent(staging(`
      <div id="pagination-hf-registry">
        <div data-section="0" data-hf-type="header-default"><p>s0 odd</p></div>
        <div data-section="1" data-hf-type="header-default"><p>s1 odd</p></div>
        <div data-section="1" data-hf-type="header-even"><p>s1 even</p></div>
      </div>
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">alpha</p>
      </div>
      <div data-section-index="1" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">beta</p>
        <p style="height: 60pt">gamma</p>
        <p style="height: 60pt">delta</p>
      </div>`));

    const result = await paginate(page);

    // Section 1 begins on global page 2: its stories alternate with the DOCUMENT page
    // number (even, odd, even), not with its own one-based position (which would give
    // the exact opposite sequence).
    expect(result.sectionIndices).toEqual([0, 1, 1, 1]);
    expect(result.pagesInSection).toEqual([1, 1, 2, 3]);
    expect(result.displayedPageNumbers).toEqual([1, 2, 3, 4]);
    expect(result.headers).toEqual(['s0 odd', 's1 even', 's1 odd', 's1 even']);
  });

  test('a numbering restart moves story parity with the displayed number', async ({
    page,
  }) => {
    // The page that RESTARTS at 1 displays an odd number, so it takes the odd story even
    // though its physical ordinal is even — exactly how Word treats a front-matter restart.
    await page.setContent(staging(`
      <div id="pagination-hf-registry">
        <div data-section="0" data-hf-type="header-default"><p>s0 odd</p></div>
        <div data-section="1" data-hf-type="header-default"><p>s1 odd</p></div>
        <div data-section="1" data-hf-type="header-even"><p>s1 even</p></div>
      </div>
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">alpha</p>
      </div>
      <div data-section-index="1" data-page-num-start="1" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">beta</p>
        <p style="height: 60pt">gamma</p>
      </div>`));

    const result = await paginate(page);

    expect(result.displayedPageNumbers).toEqual([1, 1, 2]);
    expect(result.headers).toEqual(['s0 odd', 's1 odd', 's1 even']);
  });

  test('checks the physical-page limit before allocating an over-limit page', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">first</p>
        <p style="height: 60pt">second</p>
        <p style="height: 60pt">third</p>
      </div>`));
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

    const result = await page.evaluate(() => {
      const admitted: number[] = [];
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const engine = new PaginationEngine('staging', 'container', {
        showPageNumbers: false,
        checkPageCount: (prospective: number) => {
          admitted.push(prospective);
          if (prospective > 2) throw new Error('finalPages admission failed');
        },
      });
      let message = '';
      try {
        engine.paginate();
      } catch (error) {
        message = error instanceof Error ? error.message : String(error);
      }
      return {
        admitted,
        allocated: document.querySelectorAll('#container .page-box').length,
        message,
      };
    });

    expect(result).toEqual({
      admitted: [1, 2, 3],
      allocated: 2,
      message: 'finalPages admission failed',
    });
  });
});

test.describe('w:cols column geometry', () => {
  test('a final columned continuous section fills the available column before advancing', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p style="height: 20pt">alpha</p>
        <p style="height: 20pt">beta</p>
      </div>`));

    const { content, columnCounts } = await paginate(page);

    // Both paragraphs fit below the title in the left column. A balanced box would move
    // beta to the right even though most of the page body remains empty.
    expect(content).toEqual([['title', 'alphabeta']]);
    expect(columnCounts).toEqual([['', '2']]);
    const positions = await page.evaluate(() => {
      const paragraphs = Array.from(document.querySelectorAll<HTMLElement>('#container p'));
      return paragraphs.map(p => ({ text: p.textContent, x: p.getBoundingClientRect().x,
        y: p.getBoundingClientRect().y }));
    });
    expect(positions[2].x).toBeCloseTo(positions[1].x, 1);
    expect(positions[2].y).toBeGreaterThan(positions[1].y);
  });

  test('a final continuous column section reflows around the preceding title without extra columns', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">alpha</p>
        <p style="height: 60pt">beta</p>
        <p style="height: 60pt">gamma</p>
      </div>`));
    const result = await paginate(page);
    expect(result.content).toEqual([['title', 'alpha'], ['betagamma']]);
    const overflow = await page.evaluate(() => Array.from(
      document.querySelectorAll<HTMLElement>('#container .page-content > div'),
    ).filter(e => e.style.columnCount === '2').map(e => e.scrollWidth - e.clientWidth));
    expect(overflow.every(width => width <= 1)).toBe(true);
  });

  test('a column section followed by a continuous section still balances', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p style="height: 20pt; margin: 0">alpha</p>
        <p style="height: 20pt; margin: 0">beta</p>
      </div>
      <div data-section-index="1" data-section-type="continuous" ${PAGE_GEOMETRY}>
        <p style="height: 20pt; margin: 0">after columns</p>
      </div>`));
    const result = await paginate(page);
    expect(result.content).toEqual([['alphabeta', 'after columns']]);
    const positions = await page.evaluate(() => Array.from(
      document.querySelectorAll<HTMLElement>('#container p'),
    ).map(p => ({ x: p.getBoundingClientRect().x, y: p.getBoundingClientRect().y })));
    expect(positions[1].x).toBeGreaterThan(positions[0].x);
    expect(positions[1].y).toBeCloseTo(positions[0].y, 1);
    expect(positions[2].y).toBeGreaterThan(positions[0].y);
  });

  test('a column overflow keeps a feasible heading and following paragraph together', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 80pt; margin: 0">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p data-keep-with-next="true" style="height: 20pt; margin: 0">heading</p>
        <p style="height: 30pt; margin: 0">body</p>
      </div>`));
    expect((await paginate(page, 1000)).content).toEqual([['title'], ['headingbody']]);
  });

  test('a column continuation beginning with a table drops the preceding heading space-before', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 60pt; margin: 0">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p style="height: 30pt; margin: 10pt 0 0">heading</p>
        <table style="height: 60pt; margin: 0; border-collapse: collapse">
          <tbody><tr><td>table</td></tr></tbody>
        </table>
      </div>`));
    const result = await paginate(page, 1000);
    expect(result.content).toEqual([['title', 'heading'], ['table']]);
    const continuation = await page.evaluate(() => {
      const content = document.querySelectorAll<HTMLElement>('#container .page-content')[1];
      const columns = content.firstElementChild as HTMLElement;
      return {
        marginTop: parseFloat(getComputedStyle(columns).marginTop),
        overflowX: columns.scrollWidth - columns.clientWidth,
        overflowY: columns.scrollHeight - columns.clientHeight,
      };
    });
    expect(continuation.marginTop).toBe(0);
    expect(continuation.overflowX).toBeLessThanOrEqual(1);
    expect(continuation.overflowY).toBeLessThanOrEqual(1);
  });

  test('an explicit page break ends a final continuous column container', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt; margin: 0">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p data-source-anchor-id="alpha" style="height: 20pt; margin: 0">alpha</p>
        <div data-page-break="true"></div>
        <p data-source-anchor-id="beta" style="height: 20pt; margin: 0">beta</p>
      </div>`));
    expect((await paginate(page, 1000)).content).toEqual([['title', 'alpha'], ['beta']]);
    const anchors = await page.locator('#container [data-source-anchor-id]').evaluateAll(
      elements => elements.map(element => element.getAttribute('data-source-anchor-id')),
    );
    expect(anchors).toEqual(['alpha', 'beta']);
  });

  test('a final continuous column section reserves its footnote band before flowing', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" ${PAGE_GEOMETRY}>
        <p style="height: 20pt; margin: 0">title</p>
      </div>
      <div data-section-index="1" data-section-type="continuous"
           data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p style="height: 30pt; margin: 0">alpha<sup data-footnote-id="f1">1</sup></p>
        <p style="height: 30pt; margin: 0">beta</p>
        <p style="height: 30pt; margin: 0">gamma</p>
      </div>
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f1">
          <span class="footnote-number">1</span>
          <div class="footnote-content"><p style="height: 20pt; margin: 0">note</p></div>
        </div>
      </div>`));
    const result = await paginate(page, 1000);
    // The note and its separator shorten both columns, so gamma belongs to the next page.
    expect(result.content).toEqual([['title', 'alpha1beta'], ['gamma']]);
    expect(result.footnotes[0]).toContain('note');
    expect(result.footnotes[1]).toBe('');
    const geometry = await page.evaluate(() => {
      const columns = document.querySelector<HTMLElement>('#container .page-content > div')!;
      const notes = document.querySelector<HTMLElement>('#container .page-footnotes')!;
      return { overflow: columns.scrollWidth - columns.clientWidth,
        bodyBottom: columns.getBoundingClientRect().bottom,
        notesTop: notes.getBoundingClientRect().top };
    });
    expect(geometry.overflow).toBeLessThanOrEqual(1);
    expect(geometry.bodyBottom).toBeLessThanOrEqual(geometry.notesTop);
  });

  test('a long columned section splits across pages at block boundaries', async ({ page }) => {
    await page.setContent(staging(`
      <div data-section-index="0" data-cols="2" data-col-gap="10" ${PAGE_GEOMETRY}>
        <p style="height: 60pt">alpha</p>
        <p style="height: 60pt">beta</p>
        <p style="height: 60pt">gamma</p>
        <p style="height: 60pt">delta</p>
      </div>`));

    const { content, columnCounts } = await paginate(page);

    // Balancing may split a paragraph across columns, as Word does: three 60pt
    // blocks balance to ~90pt and fit the 100pt body, while a fourth would need
    // 120pt — so the section fragments at that block boundary.
    expect(content).toEqual([['alphabetagamma'], ['delta']]);
    expect(columnCounts).toEqual([['2'], ['2']]);
  });
});

test.describe('VP003 end-to-end (issue #413)', () => {
  test('title and two-column body share one page, laid out in columns', async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });

    const bytes = new Uint8Array(
      fs.readFileSync(path.join(TEST_FILES_DIR, 'VP/VP003-Two-Column-Section.docx')));
    const conversion = await page.evaluate(
      input => (window as any).DocxodusTests.convertToHtmlWithPagination(
        new Uint8Array(input), 1, 1),
      Array.from(bytes)
    );
    expect(conversion.error).toBeUndefined();

    await page.setContent(conversion.html!);
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

    const result = await page.evaluate(() => {
      const staging = document.getElementById('pagination-staging') as HTMLElement;
      const container = document.getElementById('pagination-container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const { totalPages } = new PaginationEngine(staging, container, {
        showPageNumbers: false,
      }).paginate();

      const columnBlocks = Array.from(
        container.querySelectorAll<HTMLElement>('.page-content > div')
      ).filter(block => block.style.columnCount === '2');
      return { totalPages, columnBlockCount: columnBlocks.length,
        columnOverflow: columnBlocks.map(block => block.scrollWidth - block.clientWidth),
        firstParagraphInsets: columnBlocks.map(block =>
          block.firstElementChild!.getBoundingClientRect().top - block.getBoundingClientRect().top),
      };
    });

    // LibreOffice and Word lay this fixture out on ONE page: the continuous
    // section starts on the title's page and its body flows into two columns.
    expect(result.totalPages).toBe(1);
    expect(result.columnBlockCount).toBeGreaterThan(0);
    expect(result.columnOverflow.every(width => width <= 1)).toBe(true);
    // Heading space-before collapses with the title's space-after in the outer budget.
    expect(result.firstParagraphInsets).toEqual([0]);
  });
});
