import { test, expect } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

test.describe('Pagination footnote continuations', () => {
  test('does not create a blank page for an empty footnote continuation at a page break', async ({ page }) => {
    await page.setContent(`
      <style>
        #staging { font: 16px/16px Arial; }
        .body { height: 20pt; margin: 0; }
        .footnote-item { padding-top: 70pt; }
        .footnote-content > p { height: 20pt; margin: 0; }
      </style>
      <div id="staging">
        <div id="pagination-footnote-registry">
          <div class="footnote-item" data-footnote-id="f1">
            <span class="footnote-number">1</span>
            <div class="footnote-content"><p>note</p></div>
          </div>
        </div>
        <div data-section-index="0"
             data-page-width="122" data-page-height="122"
             data-content-width="100" data-content-height="120"
             data-margin-top="1" data-margin-right="1"
             data-margin-bottom="1" data-margin-left="1">
          <p class="body">body <sup data-footnote-id="f1">1</sup></p>
          <div data-page-break="true"></div>
        </div>
      </div>
      <div id="container"></div>`);
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

    const result = await page.evaluate(() => {
      const staging = document.getElementById('staging') as HTMLElement;
      const container = document.getElementById('container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const pagination = new PaginationEngine(staging, container, {
        showPageNumbers: false,
      }).paginate();
      const pageContents = Array.from(container.querySelectorAll('.page-content')) as HTMLElement[];

      return {
        totalPages: pagination.totalPages,
        pageBoxes: container.querySelectorAll('.page-box').length,
        emptyPageContents: pageContents.filter(pageContent =>
          !(pageContent.textContent || '').trim()
        ).length,
        blankPages: Array.from(container.querySelectorAll<HTMLElement>('.page-box')).filter(box =>
          !(box.textContent || '').trim()).length,
        notePages: Array.from(container.querySelectorAll<HTMLElement>('.page-box')).map(box =>
          (box.querySelector('.page-footnotes')?.textContent || '').replace(/\s+/g, ' ').trim()),
      };
    });

    // The note, 90pt with its padding, does not fit beside its citing line, but it fits a fresh
    // page's note band, which may take the whole band (issue #1178): it starts there, whole, rather
    // than being clipped beside its citation. No page is blank.
    expect(result.totalPages).toBe(2);
    expect(result.pageBoxes).toBe(2);
    expect(result.blankPages).toBe(0);
    expect(result.notePages).toEqual(['', '1 note']);
  });
});
