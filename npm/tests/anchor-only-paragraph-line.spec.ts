import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { outOfFlowDocx, textBoxRun } from './docx-out-of-flow-fixture.js';

/**
 * A paragraph whose only content is a floating text box keeps its paragraph-mark line in the
 * paginated view, as Word and LibreOffice lay it out: only the box leaves the text flow. The
 * paginator used to leave such a paragraph with a zero-height line, so the text after it moved up
 * by one line (issue #880). LibreOffice, measured on the same three shapes, puts the following
 * paragraph exactly where it puts it after an empty paragraph.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const BEFORE = '<w:p><w:r><w:t>Before the box.</w:t></w:r></w:p>';
const AFTER = '<w:p><w:r><w:t>After the box.</w:t></w:r></w:p>';

/** How far below the top of the "Before" paragraph the "After" paragraph starts, in points. */
async function afterParagraphOffset(page: Page, bytes: Uint8Array): Promise<number> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const D = (window as any).Docxodus;
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, false, false, false, false, null, false,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
      showPageNumbers: false,
      fragmentParagraphs: true,
      layoutToken: { documentVersion: 0, rendererFingerprint: 'anchor-only-line' },
    }).paginate();
    const paragraph = (text: string) => Array.from(container.querySelectorAll<HTMLElement>('p'))
      .find((p) => p.textContent?.includes(text))!;
    const pointsPerPixel = 72 / 96;
    return (paragraph('After the box.').getBoundingClientRect().top
      - paragraph('Before the box.').getBoundingClientRect().top) * pointsPerPixel;
  }, Array.from(bytes));
}

test.describe('a paragraph holding only a floating text box keeps its line', () => {
  for (const verticalFrom of ['paragraph', 'page']) {
    test(`text after it sits where it sits after an empty paragraph (box anchored to ${verticalFrom})`, async ({ page }) => {
      const withBox = await afterParagraphOffset(page, outOfFlowDocx(
        `${BEFORE}<w:p>${textBoxRun(verticalFrom, ['Inside the text box'])}</w:p>${AFTER}`));
      const withEmpty = await afterParagraphOffset(page, outOfFlowDocx(`${BEFORE}<w:p/>${AFTER}`));
      const withNothing = await afterParagraphOffset(page, outOfFlowDocx(`${BEFORE}${AFTER}`));

      expect(Math.abs(withBox - withEmpty), JSON.stringify({ withBox, withEmpty })).toBeLessThan(0.5);
      expect(withBox - withNothing, JSON.stringify({ withBox, withNothing })).toBeGreaterThan(5);
    });
  }
});
