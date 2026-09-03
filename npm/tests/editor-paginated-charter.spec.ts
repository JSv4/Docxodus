import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

/**
 * The paginated editor on the fixture it is documented against — NVCA's Model Certificate of
 * Incorporation, four sections, 94 footnote citations, lower-Roman front matter.
 *
 * Two things brought this spec into existence (issue #688). The first is that
 * `docs/architecture/editor_ui_surface.md` carried a page count nobody had re-measured, and
 * nothing failed when the real number moved: the caption was prose. The count below is the same
 * claim as a test, so drift is a red run rather than a stale sentence.
 *
 * The second is the footer. Its shape is the standard legal running foot —
 * `Last Updated October 2025 [tab] PAGE` over a centered tab stop at 4680 twips — and the page
 * number was in the DOM, at the correct value, painted *on top of the date*, because tab geometry
 * was only ever resolved for the body story. "Present in the DOM" is exactly what a naive
 * assertion would have checked and exactly what was true while the bug shipped, so the assertion
 * here is geometric: the number's painted box must clear the label's, and must sit on the stop.
 */

/**
 * The centered footer tab stop this charter declares: `w:pos="4680"` twips = 3.25in = 312 CSS px,
 * measured from the start of the text column — which is where the footer band's own left edge is.
 * That is the midpoint of the 468pt column, so the number sits centered under the body.
 */
const FOOTER_TAB_STOP_PX = 312;

/**
 * Settled page count for this fixture in the paginated editor at the viewport below.
 *
 * The editor deliberately does NOT fragment paragraphs across pages (`fragmentParagraphs: false`
 * in `mountPaginated`): a fragment has one addressable head, and the editor's model is one
 * addressable node per anchor. Fragmentation is worth 1 page on this document, so it is not the
 * lever anyone reaching for a smaller number is looking for. Re-measure and update deliberately.
 */
const PAGES = 53;
const FOOTNOTE_CITATIONS = 94;
const SECTIONS = 4;
/** Body-scope blocks the editor makes addressable. Header, footer and note blocks are not these. */
const BODY_BLOCKS = 234;

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 60000 });
}

test.describe('Paginated editor — NVCA charter', () => {
  test('settles at the documented page count with a legible running foot', async ({ page }) => {
    test.setTimeout(300000);

    const pageErrors: string[] = [];
    page.on('pageerror', (error) => pageErrors.push(String(error)));

    await page.setViewportSize({ width: 1440, height: 1100 });
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);

    const bytes = new Uint8Array(
      fs.readFileSync(path.join(TEST_FILES_DIR, 'NVCA-Model-COI.docx')));
    await page.evaluate((array: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(array);
    }, Array.from(bytes));

    const measured = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.style.cssText = 'width:1200px;margin:0 auto;background:white';
      document.body.appendChild(container);
      D.DocxEditor.open(container, (window as any).testDocxBytes, D, {
        paginated: true,
        headerFooter: true,
      });

      const all = (selector: string) =>
        Array.from(container.querySelectorAll<HTMLElement>(selector));

      // The first page's footer: the label runs, then a centered tab, then the PAGE field. The
      // converter emits the label-plus-tab as one aligned segment — an inline-flex box as wide as
      // the advance — with the label's runs in a nowrap span inside it.
      const footer = container.querySelector<HTMLElement>('.page-box .page-footer')!;
      const pageNumber = footer.querySelector<HTMLElement>('[data-field="PAGE"]')!;
      const segment = footer.querySelector<HTMLElement>('span[style*="inline-flex"]')!;
      const label = segment.querySelector<HTMLElement>('span[style*="nowrap"]')!;

      const box = (element: HTMLElement) => {
        const rect = element.getBoundingClientRect();
        return { left: rect.left, right: rect.right, top: rect.top, width: rect.width };
      };

      return {
        pages: all('.page-box').length,
        contents: all('.page-content').length,
        footers: all('.page-footer').length,
        sections: new Set(all('.page-box').map((box) => box.dataset.sectionIndex)).size,
        citations: all('.footnote-ref').length,
        // A citation that survived pagination is inside a page box; one that did not is
        // still in the document but attached to nothing the reader can reach.
        citationsOnPages: all('.footnote-ref').filter((ref) => ref.closest('.page-box')).length,
        // Body blocks only: header and footer stories are cloned onto every page and note
        // paragraphs live in their own scope, so both would inflate a bare anchor count.
        bodyBlocks: all('[data-source-anchor-id*=":body:"]').length,
        pageNumberText: (pageNumber.textContent || '').trim(),
        labelText: (label.textContent || '').replace(/\u00a0/g, ' ').trim(),
        pageNumberBox: box(pageNumber),
        labelBox: box(label),
        footerLeft: footer.getBoundingClientRect().left,
      };
    });

    expect(measured.pages).toBe(PAGES);
    expect(measured.contents).toBe(PAGES);
    expect(measured.footers).toBe(PAGES);
    expect(measured.sections).toBe(SECTIONS);
    expect(measured.bodyBlocks).toBe(BODY_BLOCKS);

    // Every citation still reaches a page.
    expect(measured.citations).toBe(FOOTNOTE_CITATIONS);
    expect(measured.citationsOnPages).toBe(FOOTNOTE_CITATIONS);

    // The section numbers its front matter in lower Roman, so page one is `i` — not `1`, which
    // is what the field's cached result and the fallback `.page-number` element both carry.
    expect(measured.pageNumberText).toBe('i');
    expect(measured.labelText).toContain('Last Updated');

    // Painted, not merely present: the number starts after the label ends.
    expect(measured.pageNumberBox.left).toBeGreaterThan(measured.labelBox.right);

    // And it sits on the stop Word declared, centered over it.
    const centre =
      (measured.pageNumberBox.left + measured.pageNumberBox.right) / 2 - measured.footerLeft;
    expect(centre).toBeGreaterThan(FOOTER_TAB_STOP_PX - 8);
    expect(centre).toBeLessThan(FOOTER_TAB_STOP_PX + 8);

    expect(pageErrors).toEqual([]);
  });
});
