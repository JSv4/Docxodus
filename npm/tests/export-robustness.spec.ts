import { expect, test, type Page } from '@playwright/test';
import { mixedContentDocx, tabAfterWrappingTextDocx } from './docx-mixed-content-fixture.js';

/**
 * The paginated export must not fail an ordinary renderable document (epic #846). Each seed builds a
 * synthetic document mixing the shapes that used to fail it — fonts named only through an absent
 * theme (#847), paragraphs holding only a floating text box (#849), tables after zero-spacing
 * paragraphs and exact line spacing below the font height (#848) — with lists, keep-with-next
 * chains, borders and running headers and footers. Every one must export with a complete report.
 */

const SEEDS = Array.from({ length: 24 }, (_, index) => index + 1);

async function ready(page: Page): Promise<void> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
}

test.describe('paginated export robustness', () => {
  test('exports every seeded mixed-content document', async ({ page }) => {
    test.setTimeout(600_000);
    await ready(page);

    const failures: string[] = [];
    for (const seed of SEEDS) {
      const outcome = await page.evaluate(async (bytes) => (window as any).DocxodusStandalone
        .convertFailure(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }),
      Array.from(mixedContentDocx(seed)));
      if (!outcome.unexpectedSuccess) failures.push(`seed ${seed}: ${outcome.code} ${outcome.message}`);
    }

    expect(failures).toEqual([]);
  });

  // #891: the text before a tab is measured as one unwrapped line and pinned in a no-wrap box that
  // wide, so a long sentence + tab in a table cell forces the cell past the page. Remove `fixme`
  // when #891 is fixed.
  test.fixme('exports a tab that follows text longer than its line (#891)', async ({ page }) => {
    await ready(page);

    const outcome = await page.evaluate(async (bytes) => (window as any).DocxodusStandalone
      .convertFailure(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }),
    Array.from(tabAfterWrappingTextDocx()));

    expect(outcome.unexpectedSuccess).toBe(true);
  });
});
