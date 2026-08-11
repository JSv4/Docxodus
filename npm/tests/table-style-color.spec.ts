import { expect, test } from '@playwright/test';
import {
  ACCENT5_TINT_33,
  ACCENT5_TINT_99,
  STALE_BAND_FILL,
  STALE_BORDER_COLOR,
  STALE_HEADER_FILL,
  THEME_ACCENT5,
  generateTableColorDocx,
} from './docx-table-color-fixture.js';

/**
 * Issue #399 — where a table style's colours come from.
 *
 * The tracked `merged-table` case (HC029) draws every colour from the `Grid Table 4 Accent 5`
 * style: a header-row fill, banded-row fills, and a border colour, each declaring BOTH a
 * `w:themeColor`/`w:themeFill` reference and a cached literal. In a Word-written file those two
 * always agree, because Word rewrites the cache whenever it applies the theme — so the tracked
 * fixture cannot say which one a renderer used, and both engines paint identical pixels.
 *
 * This generated table makes the question decidable by letting them DISAGREE. Per ECMA-376 the
 * theme reference is the authority and `w:color`/`w:fill` is the cached last resolution, so a
 * document whose theme changed without a cache rewrite must follow the theme.
 */

const rgb = (hex: string) =>
  `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;

interface TableColors {
  headerFill: string;
  bandFills: string[];
  plainFill: string;
  borderColor: string;
}

async function renderTable(
  page: import('@playwright/test').Page,
  options: { staleCache?: boolean } = {},
): Promise<TableColors> {
  const bytes = Array.from(generateTableColorDocx(options));
  return page.evaluate(async (input) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', true, '', -1, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', true, true, false, false, false);
    document.body.innerHTML = '<main id="tbl-root"></main>';
    (window as any).DocxodusPagination.paginateHtml(html, document.getElementById('tbl-root'), {
      scale: 1, showPageNumbers: false, pageGap: 0, fragmentParagraphs: false,
    });
    await document.fonts.ready;
    await new Promise<void>(res => requestAnimationFrame(() => requestAnimationFrame(() => res())));

    const pageBox = document.querySelector('#tbl-root .page-box') as HTMLElement;
    const rows = Array.from(pageBox.querySelectorAll('tr'));
    const fillOf = (row: Element) => {
      const cell = row.querySelector('td, th') as HTMLElement;
      return getComputedStyle(cell).backgroundColor;
    };
    const headerCell = rows[0].querySelector('td, th') as HTMLElement;
    return {
      headerFill: fillOf(rows[0]),
      bandFills: [fillOf(rows[1]), fillOf(rows[3])],
      plainFill: fillOf(rows[2]),
      borderColor: getComputedStyle(headerCell).borderTopColor,
    };
  }, bytes);
}

test.beforeEach(async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined,
    { timeout: 30_000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });
});

test.describe('table-style colours', () => {
  test('conditional formatting paints the theme-derived header, band and border colours',
    async ({ page }) => {
      const colors = await renderTable(page);

      // These are exactly the values LibreOffice paints for the tracked HC029 fixture.
      expect(colors.headerFill, 'first-row fill = accent5').toBe(rgb(THEME_ACCENT5));
      for (const band of colors.bandFills) {
        expect(band, 'banded-row fill = accent5 tint 33').toBe(rgb(ACCENT5_TINT_33));
      }
      expect(colors.borderColor, 'table border = accent5 tint 99').toBe(rgb(ACCENT5_TINT_99));
    });

  test('a row outside the header and the bands takes no fill', async ({ page }) => {
    const colors = await renderTable(page);
    // `w:tblLook` turns banding on with a band size of one row, so row 2 is unbanded. If a
    // renderer banded every row, the fill assertions above would still pass.
    expect(colors.plainFill).not.toBe(rgb(ACCENT5_TINT_33));
    expect(colors.plainFill).not.toBe(rgb(THEME_ACCENT5));
  });

  /**
   * The discriminating case. Each of the three declarations keeps its `w:themeColor`/`w:themeFill`
   * reference but carries a deliberately WRONG cached literal, and each literal is a different
   * colour so the failure names which property regressed.
   */
  test('a stale cached literal loses to the theme reference it accompanies', async ({ page }) => {
    const colors = await renderTable(page, { staleCache: true });

    expect(colors.headerFill, 'header fill must resolve w:themeFill, not the stale w:fill')
      .toBe(rgb(THEME_ACCENT5));
    expect(colors.headerFill).not.toBe(rgb(STALE_HEADER_FILL));

    for (const band of colors.bandFills) {
      expect(band, 'band fill must resolve w:themeFill with its tint').toBe(rgb(ACCENT5_TINT_33));
      expect(band).not.toBe(rgb(STALE_BAND_FILL));
    }

    // Borders used to be the odd one out: shading resolved the theme while the border read the
    // cache, so one style's fill and border disagreed about the same accent colour (issue #399).
    expect(colors.borderColor, 'border must resolve w:themeColor, not the stale w:color')
      .toBe(rgb(ACCENT5_TINT_99));
    expect(colors.borderColor).not.toBe(rgb(STALE_BORDER_COLOR));
  });

  test('the resolved colours are identical whether or not the cache agrees', async ({ page }) => {
    // The whole point of the attribution: for a Word-written file (cache in sync) the two paths
    // are indistinguishable, which is why the tracked fixture showed no colour disagreement.
    const fresh = await renderTable(page);
    const stale = await renderTable(page, { staleCache: true });
    expect(stale).toEqual(fresh);
  });
});
