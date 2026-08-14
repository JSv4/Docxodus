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

// The second phone regression, same page, different grid property: the frame
// kept its 26 rows but the ART TILTED — the title card's X read as a K and the
// logo smeared off the right edge. The canvas draws with box drawing, block
// elements (█ ░ ▓) and geometric shapes (▶ ◀), and Android's monospace face
// covers NONE of them: each one lands in a proportional fallback whose advance
// is not the cell, so every cell after one is displaced — by a different amount
// on each row, worst exactly where the art is densest. The fix pins the canvas
// to a self-hosted 17 KB subset whose every glyph advances identically
// (docs/demo/fonts/, built by docs/demo/tools/build-canvas-font.sh).
//
// The specs below recreate that condition honestly: rather than overriding what
// the page asks for, they redefine what the NAME "Courier New" RESOLVES TO on
// this platform — a monospace face for Latin-1, a proportional face above it,
// which is Roboto Mono plus the Noto Sans Symbols 2 fallback. The pin is then
// free to win, and aborting the font's request is the control that shows the
// pin is what's doing the work.

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

/** What a phone actually has, expressed as font resolution rather than as an
 * override: "Courier New" covers Latin-1 and nothing above it, so the canvas's
 * box drawing / block elements / geometric shapes fall through to a
 * PROPORTIONAL face — exactly what Android does with Roboto Mono and the Noto
 * Sans Symbols 2 fallback. Installed before the document so the very first
 * frame is laid out under it. */
async function emulateAndroidFontCoverage(page: Page) {
  await page.addInitScript(() => {
    const css = `
      @font-face { font-family: "Courier New"; src: local("Liberation Mono");
                   unicode-range: U+0000-00FF; }
      @font-face { font-family: "Courier New"; src: local("DejaVu Sans");
                   unicode-range: U+0100-10FFFF; }`;
    const inject = () => {
      const style = document.createElement('style');
      style.textContent = css;
      document.head.append(style);
    };
    if (document.head) inject();
    else document.addEventListener('DOMContentLoaded', inject);
  });
}

/** Every canvas row is COLS cells wide, so on an intact grid every row is the
 * same width — whatever glyphs it happens to hold. Returns the spread between
 * the widest and narrowest row, IN CELLS, which is the tilt measured directly:
 * a row 12 cells wide of its neighbours is a row whose art has slid 12 columns
 * out of true by its right edge. */
async function rowWidthSpreadInCells(page: Page): Promise<number> {
  return page.evaluate(() => {
    const el = (window as any).__arcade.canvasElement() as HTMLElement;
    // The frame is one paragraph of rows joined by <br>; text nodes between two
    // <br>s are one row, however many coloured runs it took to draw it.
    const rows: Text[][] = [];
    let current: Text[] = [];
    const walk = (node: Node) => {
      for (const child of Array.from(node.childNodes)) {
        if (child.nodeName === 'BR') { rows.push(current); current = []; }
        else if (child.nodeType === Node.TEXT_NODE) current.push(child as Text);
        else walk(child);
      }
    };
    walk(el);
    rows.push(current);

    const widths = rows.filter((r) => r.length).map((nodes) => {
      let left: number | null = null;
      let right = 0;
      for (const text of nodes) {
        const range = document.createRange();
        range.selectNodeContents(text);
        const box = range.getBoundingClientRect();
        if (left === null) left = box.left;
        right = box.right;
      }
      return right - (left ?? 0);
    });
    const min = Math.min(...widths);
    const max = Math.max(...widths);
    return (max - min) / (min / 92); // cell width = a full row / COLS
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

  test('the title card stays on its grid when the platform has no block glyphs', async ({ page }) => {
    // The reported tilt, at its worst: the attract screen's block-letter logo
    // is five rows of █ ░ ▓ with different glyph counts, so a fallback advance
    // displaces each of them by a different amount.
    await emulateAndroidFontCoverage(page);
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 55, { timeout: 60000 });
    await page.evaluate(() => (window as any).__arcade.pause());

    expect(await page.evaluate(() =>
      document.fonts.check('8pt "Docxodus Canvas Mono"'))).toBe(true);
    // Sub-cell: rows differ only by sub-pixel layout rounding, never by a glyph.
    expect(await rowWidthSpreadInCells(page)).toBeLessThan(0.1);
    expect(await canvasLineBoxes(page)).toBe(26);
  });

  test('every cartridge keeps its grid on the same platform', async ({ page }) => {
    test.setTimeout(180000); // one boot, then three cartridges animated in turn
    // The games draw a box-drawing bezel on every row and the raycaster shades
    // its walls with ▒ █, so the tilt is not an attract-screen-only property.
    await emulateAndroidFontCoverage(page);
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&intro=0&cart=quest`);
    await page.locator('#boot').click();
    await waitForBoot(page);

    for (const cart of ['quest', 'dungeon', 'e1m1']) {
      await page.evaluate((name) => {
        (window as any).__arcade.setCart(name);
        (window as any).__arcade.resume();
      }, cart);
      const from = await page.evaluate(() => (window as any).__arcade.frames());
      await page.waitForFunction(
        (n) => (window as any).__arcade.frames() >= n, from + 20, { timeout: 60000 });
      await page.evaluate(() => (window as any).__arcade.pause());

      expect.soft(await rowWidthSpreadInCells(page), `${cart} row widths`).toBeLessThan(0.1);
      expect.soft(await canvasLineBoxes(page), `${cart} line boxes`).toBe(26);
    }
    expect(test.info().errors).toHaveLength(0);
  });

  test('without the pinned font the same platform tilts the art — the pin is the fix', async ({ page }) => {
    // The control. If this ever stops failing to hold the grid, the emulation
    // above has stopped being adverse and the two specs before it prove nothing.
    await emulateAndroidFontCoverage(page);
    await page.route('**/docxodus-canvas-mono.woff2', (route) => route.abort());
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 55, { timeout: 60000 });
    await page.evaluate(() => (window as any).__arcade.pause());

    expect(await page.evaluate(() =>
      document.fonts.check('8pt "Docxodus Canvas Mono"'))).toBe(false);
    // Whole cells of drift, concentrated in the logo's five rows.
    expect(await rowWidthSpreadInCells(page)).toBeGreaterThan(1);
    // …and note the rows still do not WRAP: PR #424's `white-space: pre` pin
    // holds independently, which is why the field report was a tilt and not
    // the earlier stacking garble.
    expect(await canvasLineBoxes(page)).toBe(26);
  });

  test('the cabinet hands a phone thumb controls, and a way to fire', async ({ page }) => {
    // The controls the cabinet drew before this were one wrapped bar under the
    // document: four 44px arrows on a row of their own beneath the cartridge
    // chips, the transport row and the telemetry — a stack that ate the bottom
    // of a phone screen, sat nowhere near a thumb, and offered no Space at all,
    // so Freedoom could be walked but never fought on a touch screen.
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&intro=0&cart=e1m1`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 3, { timeout: 60000 });

    await expect(page.locator('.dxa-controls')).toHaveAttribute('data-compact', 'true');
    await expect(page.locator('#pad .dxa-fire')).toBeVisible();
    await expect(page.locator('#dockcarts')).toBeHidden();

    // Thumb reach: bottom corners, clear of the game screen's middle, and hit
    // targets no smaller than the 44px both platform guidelines ask for.
    const geometry = await page.evaluate(() => {
      const box = (selector: string) => document.querySelector(selector)!.getBoundingClientRect();
      const dpad = box('.dxa-dpad');
      const fire = box('.dxa-fire');
      return {
        bothLow: dpad.top > window.innerHeight / 2 && fire.top > window.innerHeight / 2,
        opposedCorners: dpad.right < window.innerWidth / 2 && fire.left > window.innerWidth / 2,
        onScreen: dpad.bottom <= window.innerHeight && fire.bottom <= window.innerHeight,
        tapTargets: Math.min(dpad.width / 3, dpad.height / 3, fire.width, fire.height),
      };
    });
    expect(geometry.bothLow).toBe(true);
    expect(geometry.opposedCorners).toBe(true);
    expect(geometry.onScreen).toBe(true);
    expect(geometry.tapTargets).toBeGreaterThanOrEqual(40);

    // Real taps, on the touch rig: turning is the raycaster's own input path.
    const heading = await page.evaluate(() => (window as any).__arcade.game().player.dx as number);
    await page.locator('.dxa-right').dispatchEvent('pointerdown');
    await page.waitForFunction(
      (dx0) => Math.abs((window as any).__arcade.game().player.dx - dx0) > 0.05,
      heading,
      { timeout: 30000 },
    );
    await page.locator('.dxa-right').dispatchEvent('pointerup');

    // And the button a phone never had: Space, held, which is the trigger.
    await page.locator('.dxa-fire').dispatchEvent('pointerdown');
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Space'))).toBe(true);
    await page.locator('.dxa-fire').dispatchEvent('pointerup');
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Space'))).toBe(false);
  });

  test('every character on the canvas advances exactly one cell', async ({ page }) => {
    // Directly the grid's contract, measured glyph by glyph in the font the
    // canvas actually resolved to — the property that makes the art align at
    // all, rather than a proxy for it.
    await emulateAndroidFontCoverage(page);
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 55, { timeout: 60000 });
    await page.evaluate(() => (window as any).__arcade.pause());

    const measured = await page.evaluate(() => {
      const el = (window as any).__arcade.canvasElement() as HTMLElement;
      const style = getComputedStyle(el);
      const context = document.createElement('canvas').getContext('2d')!;
      context.font = `${style.fontSize} ${style.fontFamily}`;
      const cell = context.measureText('M'.repeat(20)).width / 20;
      // The converter emits bidi formatting marks (U+200E and friends) into the
      // rendered text. They are not cells and are zero-width in every font, so
      // they are measured separately rather than waved through.
      const invisible = /[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/;
      const name = (ch: string) =>
        `U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`;
      const cells: string[] = [];
      const marks: string[] = [];
      for (const ch of new Set((el.textContent ?? '').split(''))) {
        const width = context.measureText(ch.repeat(20)).width / 20;
        const bucket = invisible.test(ch) ? marks : cells;
        if (Math.abs(width - (invisible.test(ch) ? 0 : cell)) > cell * 0.005) {
          bucket.push(`${name(ch)} advances ${(width / cell).toFixed(3)} cells`);
        }
      }
      return { cells, marks };
    });
    expect(measured.cells, 'characters that do not advance one cell tilt every row they appear in')
      .toEqual([]);
    expect(measured.marks, 'formatting marks must stay zero-width').toEqual([]);
  });
});
