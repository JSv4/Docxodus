import { test, expect, Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';

// Match the capture script's optional offline mirror, including its pin check.
test.beforeEach(async ({ page }) => {
  if (!process.env.DOOM_ENGINE_PATH) return;
  const bytes = readFileSync(process.env.DOOM_ENGINE_PATH);
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(
    'efd7c34714e3753a84cf6873f2c2ae0e1a41bde1f10aaa2b0a95cb6935e8cce9');
  await page.route('https://cdn.jsdelivr.net/gh/grubbyplaya/**', route =>
    route.fulfill({ body: bytes, contentType: 'text/javascript' }));
});

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

// wad= is the webroot copy fetch-doom-iwad.mjs provides; the grid specs
// below animate the Doom cartridge and want its real framebuffer.
const OVERRIDE = 'engine=./embed.bundle.js&sound=0'
  + '&wad=' + encodeURIComponent('./vendor/freedoom1.wad.gz');

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

/** Visual line boxes inside the canvas paragraph. Each frame has a fixed row
 * count joined by `w:br`; every over-wide row
 * that WRAPS adds one more. Rows are read from the `<br>` structure (the
 * same segmentation `rowWidthSpreadInCells` uses) and each row contributes
 * 1 + how many full LINE PITCHES its own fragments' tops span: a real fold
 * displaces fragments by a whole pitch, while sub-pitch offsets are
 * rendering artifacts, not wraps — inflated glyph boxes overflow the exact
 * 10pt line box by half a leading, and mixed fallback faces differ in
 * ascent, offsets that scale with the fit zoom and so cannot be separated
 * from real folds by any fixed pixel tolerance (the margin-cropped phone
 * presentation zooms ~40% past where 2px was safe). The pitch is measured
 * from the rows themselves: the median gap between consecutive row tops. */
async function canvasLineBoxes(page: Page): Promise<number> {
  return page.evaluate(() => {
    const el = (window as any).__arcade.canvasElement() as HTMLElement;
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

    const measured = rows
      .map((nodes) => {
        let first = Infinity;
        let last = -Infinity;
        for (const text of nodes) {
          const range = document.createRange();
          range.selectNodeContents(text);
          for (const r of Array.from(range.getClientRects())) {
            if (r.height < 2) continue;
            first = Math.min(first, r.top);
            last = Math.max(last, r.top);
          }
        }
        return first === Infinity ? null : { top: first, spread: last - first };
      })
      .filter((row): row is { top: number; spread: number } => row !== null);

    const gaps = measured
      .slice(1)
      .map((row, i) => row.top - measured[i].top)
      .filter((gap) => gap > 0)
      .sort((a, b) => a - b);
    const pitch = gaps[Math.floor(gaps.length / 2)] || 1;

    return measured.reduce((total, row) => total + 1 + Math.round(row.spread / pitch), 0);
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
    const columns = rows[0].reduce((n, text) => n + (text.textContent?.length ?? 0), 0);
    const min = Math.min(...widths);
    const max = Math.max(...widths);
    return (max - min) / (min / columns); // normalize the spread to actual cells
  });
}

test.describe('Arcade on a phone-shaped viewport', () => {

  test('attract screen keeps 26 rows as the title card fills in', async ({ page }) => {
    // The reported repro surface: the intro looked fine at first, then grew
    // extra lines as the logo sweep filled the grid. Let it run well past the
    // full title card (t ≈ 4.5s at 10fps) with the text inflated throughout.
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap`);
    await page.locator('#boot').click();
    await waitForBoot(page);
    await inflateDocumentText(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 50, { timeout: 60000 });

    // The attract screen has one live control — the coin — so the pad offers
    // that and nothing else: a D-pad steering a title card is a lie about what
    // the buttons do.
    await expect(page.locator('#pad .dxa-fire')).toHaveText('START');
    await expect(page.locator('.dxa-dpad button:not([hidden])')).toHaveCount(0);

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

// ─── The pad's control map ────────────────────────────────────────────
// The thumb pad shipped with four arrows and FIRE, which is every control the
// platformer has and about half of the ones a Doom-format level needs. On a
// phone that meant no strafe, no sprint, no weapon switch, no automap, no way
// into Doom's own menu — and, decisively, no USE: E1M1's first door is thirty
// seconds in, and a door is a key press. The pad now takes its buttons from
// the cartridge (`cart.touch`), so what a thumb can reach is what that
// cartridge's keyboard map says exists.

test.describe('The pad takes its controls from the cartridge', () => {
  /** The dock alone, on a host of its own: no engine, no document, no game —
   * just the control surface, driven through the same `setPad` the arcade
   * driver calls on every cartridge change. */
  async function mountBareDock(page: Page) {
    await page.goto('/demo-arcade.html?boot=tap'); // boot=tap: nothing streams a runtime
    return page.evaluate(async () => {
      document.getElementById('boot')?.remove(); // …and nothing over the controls
      // Loaded through a computed specifier, as the CDN specs do: these modules
      // are served to the browser beside the page, not resolvable from here.
      const load = (name: string) => import(/* @vite-ignore */ `./${name}.js`);
      const [{ mountArcadeDock }, { DOOM_TOUCH }] = await Promise.all([
        load('arcade-dock'), load('doom-cart'),
      ]);
      const host = document.createElement('div');
      host.style.cssText = 'position:relative;width:100%;height:420px';
      document.body.append(host);
      const dock = mountArcadeDock(host, { anchor: 'host', ids: false });
      dock.show();
      (window as any).__dock = dock;
      (window as any).__profiles = { intro: {}, doom: DOOM_TOUCH };
      return dock.isCompact();
    });
  }

  /** Which slots the pad is offering, and what each one would send. */
  async function padMap(page: Page): Promise<Record<string, string | null>> {
    return page.evaluate(() => {
      const map: Record<string, string | null> = {};
      for (const button of Array.from(document.querySelectorAll('.dxa-pad button'))) {
        // Tray chips are minted per cartridge and carry no slot class.
        const name = Array.from(button.classList).find((c) => c.startsWith('dxa-'))?.slice(4);
        if (!name) continue;
        map[name] = (button as HTMLElement).hidden ? null : button.getAttribute('data-code');
      }
      return map;
    });
  }

  const setProfile = (page: Page, cart: string) => page.evaluate(
    (name) => (window as any).__dock.setPad((window as any).__profiles[name]), cart);

  test('FIRE touch taps confirm Doom menus and holding it still fires in a level', async ({ page }) => {
    test.setTimeout(240000);
    await page.goto(`/demo-arcade.html?${OVERRIDE}&intro=0&cart=doom`);
    await waitForBoot(page);
    await page.waitForFunction(() => {
      const state = (window as any).__arcade.game();
      if (state.status === 'error') throw new Error(state.error);
      return state.doomFrames >= 4;
    }, null, { timeout: 180000 });

    const fire = page.locator('#pad .dxa-fire');
    await expect(fire).toBeVisible();
    await expect(page.locator('#padextras')).toBeHidden();

    // Actual Chromium touch events, including pointer capture and release.
    // A synthetic pointerdown or an input.held() check cannot prove that a
    // quick tap survives until Doom processes its menu input on the next frame.
    for (let i = 0; i < 4; i++) {
      await fire.tap(); // title → New Game → episode → skill → level
      await page.waitForTimeout(700);
    }
    await page.waitForTimeout(1500); // level wipe and weapon raise

    const sample = (x0: number, y0: number, x1: number, y1: number) =>
      page.evaluate(({ x0, y0, x1, y1 }) => {
        const state = (window as any).__arcade.game();
        const pixels: number[] = [];
        for (let y = y0; y < y1; y += 2) {
          for (let x = x0; x < x1; x += 2) pixels.push(...state.pixel(x, y));
        }
        return pixels;
      }, { x0, y0, x1, y1 });
    const motion = (a: number[], b: number[]) =>
      a.reduce((sum, value, i) => sum + Math.abs(value - b[i]), 0) / a.length;
    const view = await sample(0, 35, 320, 135);
    const hud = await sample(0, 175, 320, 198);
    await page.waitForTimeout(600);
    expect(motion(view, await sample(0, 35, 320, 135)), 'an idle level holds still').toBeLessThan(1);

    // The view must actually turn while the HUD stays fixed. An opening menu
    // or attract demo cannot satisfy both the idle and controlled-motion checks.
    const cdp = await page.context().newCDPSession(page);
    const hold = async (selector: string) => {
      const box = await page.locator(selector).boundingBox();
      expect(box).not.toBeNull();
      await cdp.send('Input.dispatchTouchEvent', {
        type: 'touchStart', touchPoints: [{ x: box!.x + box!.width / 2, y: box!.y + box!.height / 2 }],
      });
    };
    const release = () => cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    await hold('#pad .dxa-right');
    await page.waitForTimeout(900);
    await release();
    expect(motion(view, await sample(0, 35, 320, 135)), 'FIRE must get past the opening menus')
      .toBeGreaterThan(6);
    expect(motion(hud, await sample(0, 175, 320, 198))).toBeLessThan(1);

    const ammo = await sample(2, 171, 42, 194);
    await hold('#pad .dxa-fire');
    await page.waitForTimeout(1400);
    await release();
    await page.waitForTimeout(700);
    expect(motion(ammo, await sample(2, 171, 42, 194)), 'holding FIRE spends ammunition')
      .toBeGreaterThan(1);
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Space'))).toBe(false);
    expect(await page.evaluate(() => (window as any).__arcade.playing())).toBe(true);
    await cdp.detach();
  });

  test('the rarer Doom keys ride in a tray instead of under a thumb', async ({ page }) => {
    await mountBareDock(page);

    // Nothing to open for a cartridge with no extra keys…
    await setProfile(page, 'intro');
    await expect(page.locator('.dxa-keys')).toBeHidden();
    await expect(page.locator('.dxa-extras')).toBeHidden();

    // …and for Doom, one tap away rather than crowding the game screen.
    await setProfile(page, 'doom');
    await expect(page.locator('.dxa-keys')).toBeVisible();
    await expect(page.locator('.dxa-extras')).toBeHidden();
    await page.locator('.dxa-keys').click();
    await expect(page.locator('.dxa-extras')).toBeVisible();
    expect(await page.locator('.dxa-extras button').evaluateAll(
      (buttons) => buttons.map((b) => b.getAttribute('data-code')))).toEqual(
      ['Digit1', 'Digit2', 'Digit3', 'Digit4', 'Digit5', 'Digit6', 'Digit7',
        'KeyM', 'KeyQ', 'Enter']);
    // It stays open: working Doom's menu is several taps in a row.
    await expect(page.locator('.dxa-extras')).toBeVisible();
    await page.locator('.dxa-keys').click();
    await expect(page.locator('.dxa-extras')).toBeHidden();

    // Switching to a cartridge without extras withdraws the toggle with them.
    await page.locator('.dxa-keys').click();
    await setProfile(page, 'intro');
    await expect(page.locator('.dxa-keys')).toBeHidden();
    await expect(page.locator('.dxa-extras')).toBeHidden();
    expect(await page.locator('.dxa-extras button').count()).toBe(0);
  });

  test('Doom hands the phone its door key, and its weapons', async ({ page }) => {
    // The reported gap, at the cartridge it actually blocked: USE. The pad is
    // pointed at Doom's keys the moment the cartridge is selected, so this
    // asserts the input path rather than waiting on a 10 MB IWAD to render.
    test.setTimeout(150000);
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&intro=0&cart=doom`);
    await page.locator('#boot').click();
    await waitForBoot(page);

    const use = page.locator('#pad .dxa-use');
    await expect(use).toBeVisible();
    await expect(use).toHaveAttribute('aria-label', /door/i);
    await use.dispatchEvent('pointerdown');
    expect(await page.evaluate(() => (window as any).__arcade.input.held('KeyE'))).toBe(true);
    await use.dispatchEvent('pointerup');
    expect(await page.evaluate(() => (window as any).__arcade.input.held('KeyE'))).toBe(false);

    // A weapon is a digit; the tray sends it the same way, through the same
    // transition log Doom's own input layer drains.
    await page.locator('#padkeys').click();
    const shotgun = page.locator('#padextras button[data-code="Digit3"]');
    await shotgun.dispatchEvent('pointerdown');
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Digit3'))).toBe(true);
    await shotgun.dispatchEvent('pointerup');
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Digit3'))).toBe(false);

    // Pausing releases every held key before the editor takes the keyboard back.
    await page.evaluate(() => (window as any).__arcade.pause());
    expect(await page.evaluate(() => (window as any).__arcade.input.held('KeyE'))).toBe(false);
  });
});
