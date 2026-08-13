import { test, expect, Page } from '@playwright/test';

// THE DOCX ARCADE on a phone — the mobile-Chrome regression that garbled the
// GitHub Pages demo. On Android the canvas rows render WIDER than authored:
// the system has no Courier New (Chrome substitutes a wider monospace face)
// and mobile Chrome's text autosizer (the "Text scaling" accessibility
// setting) inflates text-heavy blocks outright — the ribbon chrome's existing
// ancestor-level `-webkit-text-size-adjust: 100%` demonstrably does not
// protect the document inside it on a real phone. Once a 92-cell row outgrows
// the 6.5in text column, `overflow-wrap: break-word` folds it onto a second
// line and the frame stacks into garbage — extra lines multiplying as the
// animation fills the grid (sparse early frames stay narrow, which is why the
// screen "initially looks good").
//
// Neither platform trigger can be reproduced faithfully in a Linux Chromium
// (its autosizer honors the ancestor rule that the phone ignores, and its
// fontconfig substitutes a Courier-metric face), so this spec recreates the
// CONDITION — text wider than the column was authored for — directly, by
// inflating the canvas font, and asserts the driver's guarantee that makes
// every trigger harmless: the canvas paragraph is pinned to
// `white-space: pre`, so a frame row can never wrap; the worst case is a
// clipped right edge, never stacked rows. Runs under the `chromium-pixel5`
// project so the whole phone-shaped path (mobile viewport, fit-to-width CSS
// zoom of the sheet) is the one exercised.

const OVERRIDE = 'engine=./embed.bundle.js';

async function waitForBoot(page: Page) {
  await page.waitForFunction(
    () =>
      (window as any).__arcade !== undefined ||
      (window as any).__arcadeError !== undefined,
    { timeout: 90000 },
  );
  const err = await page.evaluate(() => (window as any).__arcadeError);
  expect(err, `arcade boot failed: ${err}`).toBeUndefined();
}

/** Distinct line-box tops inside the canvas paragraph. The frame is 26 rows
 * joined by `w:br`, so an intact render is exactly 26 line boxes; every
 * over-wide row that wraps adds one more. Fragments on the same line share
 * their top (same font size throughout the canvas), while the paragraph's
 * exact 10pt line pitch is ≥ ~5px even under the phone's fit-to-width zoom,
 * so a small fixed tolerance separates lines reliably — including when the
 * inflated glyph boxes are TALLER than the line pitch. */
async function canvasLineBoxes(page: Page): Promise<number> {
  return page.evaluate(() => {
    const el = (window as any).__arcade.canvasElement() as HTMLElement;
    const range = document.createRange();
    range.selectNodeContents(el);
    const rects = Array.from(range.getClientRects()).filter((r) => r.height >= 2);
    const tops: number[] = [];
    for (const r of rects) {
      if (!tops.some((t) => Math.abs(t - r.top) < 2)) tops.push(r.top);
    }
    return tops.length;
  });
}

/** What the phone does to the document: render its text substantially wider
 * than the column was authored for. `!important` beats the converter's inline
 * run styles, exactly as a UA text inflation does. */
async function inflateDocumentText(page: Page) {
  await page.addStyleTag({
    content: '#app [data-anchor], #app [data-anchor] span { font-size: 175% !important; }',
  });
}

test.describe('Arcade on a phone-shaped viewport', () => {
  test('game frame keeps 26 rows with text inflated past the column', async ({ page }) => {
    // Every game row runs bezel-to-bezel (column 92): un-pinned, 175% text
    // folds each row onto a second line (52 line boxes) — the field garble.
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&intro=0&cart=quest`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 3, { timeout: 60000 });
    await inflateDocumentText(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 6, { timeout: 60000 });
    await page.evaluate(() => (window as any).__arcade.pause());

    expect(await canvasLineBoxes(page)).toBe(26);
  });

  test('attract screen keeps 26 rows as the title card fills in', async ({ page }) => {
    // The reported repro surface: the intro looked fine at first, then grew
    // extra lines as the logo sweep filled the grid. Let it run well past the
    // full title card (t ≈ 4.5s at 10fps) with the text inflated throughout.
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await inflateDocumentText(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 50, { timeout: 60000 });
    await page.evaluate(() => (window as any).__arcade.pause());

    expect(await canvasLineBoxes(page)).toBe(26);
  });
});
