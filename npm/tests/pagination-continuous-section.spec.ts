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
}

async function paginate(page: Page): Promise<PaginatedShape> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();

    const pages = Array.from(container.querySelectorAll<HTMLElement>('.page-content'));
    return {
      content: pages.map(content =>
        Array.from(content.children).map(block => (block.textContent || '').trim())
      ),
      columnCounts: pages.map(content =>
        Array.from(content.children).map(
          block => (block as HTMLElement).style.columnCount || ''
        )
      ),
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
