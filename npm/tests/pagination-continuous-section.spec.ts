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
 * continuous break to a page break. A `w:cols` section lays out as balanced CSS
 * columns instead of one full-width column.
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

function staging(sections: string): string {
  return `
    <style>
      #staging { font: 12px/12px Arial; }
      #staging p { box-sizing: border-box; margin: 0; padding: 0; width: 100%; }
    </style>
    <div id="staging">${sections}</div>
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
}

async function paginate(page: Page): Promise<PaginatedShape> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();

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
    };
  });
}

test.describe('Continuous section breaks', () => {
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
});

test.describe('w:cols column geometry', () => {
  test('a columned continuous section shares the page as one balanced multicol block', async ({ page }) => {
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

    // One page: the 20pt title plus a balanced two-column container (~20pt tall
    // instead of 40pt stacked). Each source block lands in the container exactly once.
    expect(content).toEqual([['title', 'alphabeta']]);
    expect(columnCounts).toEqual([['', '2']]);
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
      return { totalPages, columnBlockCount: columnBlocks.length };
    });

    // LibreOffice and Word lay this fixture out on ONE page: the continuous
    // section starts on the title's page and its body balances into two columns.
    expect(result.totalPages).toBe(1);
    expect(result.columnBlockCount).toBeGreaterThan(0);
  });
});
