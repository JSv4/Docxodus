import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { bodyFillDocx } from './docx-body-fill-fixture.js';

/**
 * The converter wraps a table in an unstyled `div`, and the table's own top margin collapses
 * through that wrapper. The paginator measured each block's margins with getComputedStyle() on the
 * wrapper (zero) and its height with getBoundingClientRect() (which excludes a collapsed-through
 * margin), so every table cost the page a few unbudgeted pixels and the last block of a full page
 * ran past the band — the dominant cause of the export's "body content is clipped"
 * pagination_failure (issue #848).
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function clippedPages(page: Page, bytes: Uint8Array): Promise<string[]> {
  return page.evaluate((input) => {
    const D = (window as any).Docxodus;
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, false, false, false, false, null, false,
    );
    document.body.innerHTML = '';
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const engine = new (window as any).DocxodusPagination.PaginationEngine(
      host.querySelector('#pagination-staging'), host.querySelector('#pagination-container'),
      { scale: 1, showPageNumbers: false, pageGap: 0, fragmentParagraphs: true },
    );
    engine.paginate();
    // A block (not an inline glyph box) running past the page's body band.
    return Array.from(host.querySelectorAll<HTMLElement>('.page-content')).flatMap((content) => {
      const band = content.getBoundingClientRect().bottom;
      return Array.from(content.children as HTMLCollectionOf<HTMLElement>)
        .filter((block) => block.getBoundingClientRect().bottom > band + 0.5)
        .map((block) => `page ${content.closest<HTMLElement>('.page-box')?.dataset.pageNumber}: `
          + `${block.localName} +${(block.getBoundingClientRect().bottom - band).toFixed(1)}px`);
    });
  }, Array.from(bytes));
}

test('a table\'s collapsed-through top margin is budgeted, so no page clips its last line', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  // 10pt text at exact 12pt lines: no glyph box pokes past its line, so any overflow is a block
  // the budget placed wrongly. With every line 12pt the space left at the foot of page 1 is set by
  // everything above the lines modulo 12pt, so sweeping the first paragraph's space-before in 1pt
  // steps walks it through every residue — including those under the table's 7.5pt unbudgeted
  // margin, where the page's last line used to run past the band.
  const failures: string[] = [];
  for (let beforePt = 0; beforePt < 12; beforePt++) {
    const clipped = await clippedPages(page, bodyFillDocx({
      paragraphs: 60, fontHalfPoints: 20, lineTwips: 240,
      tableAfterFirstParagraph: true, firstSpaceBeforeTwips: beforePt * 20,
    }));
    failures.push(...clipped.map((entry) => `space before ${beforePt}pt, ${entry}`));
  }
  expect(failures).toEqual([]);
});
