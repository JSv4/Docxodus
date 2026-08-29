import { test, expect, Page } from '@playwright/test';

// Proof that the ACTUAL Doom engine runs inside a live Word document.
//
// Cartridge 3 of THE DOCX ARCADE used to be a hand-written ASCII raycaster
// walking Freedoom's E1M1 rasterized to a character grid. It is now id
// Software's own engine — doomgeneric, GPL-2.0, compiled to JavaScript — on
// Freedoom's BSD-licensed IWAD, with its 320×200 framebuffer downsampled every
// frame into one Word paragraph as half-block runs.
//
// That change moves what a spec can honestly claim. The old one steered a BFS
// autopilot through geometry it could read out of the grid; this one cannot,
// because Doom's world is BSP data in a WebAssembly heap and the only thing
// the document holds is a picture of it. So this spec proves the four things
// that actually matter about the swap:
//
//   1. it is really Doom     — the engine identifies its own IWAD and its own
//                              frame counter runs;
//   2. the picture is in the DOCUMENT — half-block glyphs and per-run `w:shd`
//                              shading in the canvas paragraph, repainted
//                              incrementally, one block at a time;
//   3. the keyboard reaches Doom — driven into a real level through Doom's own
//                              menu, the view is static when no key is held
//                              and swings when one is, while the status bar
//                              stays nailed down. A recorded attract demo
//                              would move the view on its own; a level under
//                              our control does not, which is what makes the
//                              contrast a proof and not a coincidence;
//   4. it is still a document — pause, save, reopen, and the frame is there.
//
// sound=0 keeps CI off the audio hardware; the cartridge's WebAudio path is
// guarded everywhere and never load-bearing for play.
const OVERRIDE = 'engine=./embed.bundle.js&intro=0&sound=0';

/** Boot the cabinet on the Doom cartridge and wait for the engine to come up.
 *  The budget is generous on purpose: this is a 3 MB engine plus a 10 MB
 *  IWAD, inflated in the browser, before Doom has drawn anything. */
async function bootDoom(page: Page) {
  await page.goto(`/demo-arcade.html?${OVERRIDE}&cart=doom`);
  await page.waitForFunction(
    () => (window as any).__arcade !== undefined || (window as any).__arcadeError !== undefined,
    null,
    { timeout: 120000 },
  );
  const err = await page.evaluate(() => (window as any).__arcadeError);
  expect(err, `arcade boot failed: ${err}`).toBeUndefined();
  await page.selectOption('#pace', '0'); // unthrottled: the honest frame rate
  await page.waitForFunction(
    () => {
      const state = (window as any).__arcade.game();
      if (state.status === 'error') throw new Error(`doom failed to start: ${state.error}`);
      return state.status === 'playing' && state.doomFrames > 0;
    },
    null,
    { timeout: 180000 },
  );
}

/** Mean absolute change in one horizontal band of Doom's framebuffer over
 *  `ms`, sampled through the cartridge's own state hook. Bands are in Doom's
 *  320×200 coordinates: the 3-D view lives above y=150, the status bar below
 *  y=168. */
async function bandMotion(page: Page, y0: number, y1: number, ms: number) {
  const sample = () => page.evaluate(({ y0, y1 }) => {
    const state = (window as any).__arcade.game();
    const out: number[] = [];
    for (let y = y0; y < y1; y += 2) {
      for (let x = 0; x < 320; x += 4) out.push(state.pixel(x, y)![0]);
    }
    return out;
  }, { y0, y1 });

  const before = await sample();
  await page.waitForTimeout(ms);
  const after = await sample();
  let total = 0;
  for (let i = 0; i < before.length; i++) total += Math.abs(before[i] - after[i]);
  return total / before.length;
}

/** Doom's own path into a level: the first key raises the main menu, then
 *  New Game → episode → skill, each on the default cursor position. */
async function startLevel(page: Page) {
  for (let i = 0; i < 4; i++) {
    await page.keyboard.press('Enter');
    await page.waitForTimeout(700);
  }
  await page.waitForTimeout(1500);
}

test.describe('DOOM inside a Word document', () => {
  test('boots the real engine: it names its own IWAD and its frame counter runs', async ({ page }) => {
    test.setTimeout(240000);
    await bootDoom(page);

    const first = await page.evaluate(() => ({
      cart: (window as any).__arcade.cart() as string,
      title: (window as any).__arcade.game().title as string | null,
      doomFrames: (window as any).__arcade.game().doomFrames as number,
      fallback: (window as any).__arcade.editor.lastReconcileFallback as string | null,
    }));

    expect(first.cart).toBe('doom');
    // Doom reads the IWAD's own identity out of the file and reports it
    // through DG_SetWindowTitle. Nothing but the engine produces this string.
    expect(first.title).toBe('Freedoom: Phase 1');
    expect(first.fallback).toBeNull(); // frames stay incremental

    // Doom's own counter, not the arcade's: it must keep climbing on its own.
    await page.waitForFunction(
      (n) => (window as any).__arcade.game().doomFrames > n + 60,
      first.doomFrames,
      { timeout: 60000 },
    );
  });

  test('the framebuffer is in the paragraph: half-block runs, shaded, one block repainted', async ({ page }) => {
    test.setTimeout(240000);
    await bootDoom(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 4, null, { timeout: 60000 });

    const screen = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      const element = arcade.canvasElement() as HTMLElement;
      const spans = Array.from(element.querySelectorAll('span'));
      const shaded = spans.filter((span) => {
        const fill = getComputedStyle(span).backgroundColor;
        return fill && fill !== 'transparent' && fill !== 'rgba(0, 0, 0, 0)';
      });
      return {
        text: arcade.canvasText() as string,
        rows: (arcade.canvasText() as string).split('\n').length,
        spans: spans.length,
        shaded: shaded.length,
        inks: new Set(spans.map((span) => getComputedStyle(span).color)).size,
        fallback: arcade.editor.lastReconcileFallback as string | null,
      };
    });

    // The picture: solid cells where a character row's two half-pixels agree,
    // upper-half blocks where they differ.
    expect(screen.text).toMatch(/█/);
    expect(screen.text).toMatch(/▀/);
    // Still the arcade's markdown-safe bezel — no row can open a heading or a
    // bullet, and no row is blank.
    expect(screen.text).toContain('┌');
    for (const row of screen.text.split('\n')) expect(row.trim()).not.toBe('');
    // Doom's palette arriving as real run formatting: many inks, and run
    // shading actually rendered rather than silently dropped.
    expect(screen.inks).toBeGreaterThan(20);
    expect(screen.shaded).toBeGreaterThan(50);
    expect(screen.fallback).toBeNull();
  });

  test('the keyboard reaches Doom: turning swings the view, the status bar holds still', async ({ page }) => {
    test.setTimeout(300000);
    await bootDoom(page);
    await startLevel(page);

    // Standing still in a level under our own control, the view barely moves.
    // This is the control, and it is what separates "we started a game" from
    // "we are watching one of Doom's recorded attract demos" — a demo drives
    // the player, so its view would never sit this still.
    const idleView = await bandMotion(page, 0, 150, 600);
    expect(idleView, 'the view should be near-static with no key held').toBeLessThan(5);

    // Hold a turn: Doom's own input layer has to see the key and rotate the
    // camera. Measured at roughly 24 against 0.4 standing, so the threshold is
    // nowhere near either edge.
    await page.keyboard.down('ArrowLeft');
    const turningView = await bandMotion(page, 0, 150, 600);
    const turningStatusBar = await bandMotion(page, 168, 200, 600);
    await page.keyboard.up('ArrowLeft');

    expect(turningView, 'holding a turn key should swing the 3-D view').toBeGreaterThan(5);
    expect(turningView).toBeGreaterThan(idleView * 3);
    // The status bar is drawn by the engine every frame and does not move when
    // the camera does — the frame really is Doom's composed screen, not just
    // an animated field of color.
    expect(turningStatusBar, 'the status bar should not move while turning').toBeLessThan(2);
  });

  test('pause hands the frame back as an ordinary paragraph, and it saves as a real .docx', async ({ page }) => {
    test.setTimeout(240000);
    await bootDoom(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 4, null, { timeout: 60000 });

    const result = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      arcade.pause();
      const text = arcade.canvasText() as string;
      const bytes: Uint8Array = arcade.save();
      const handle = arcade.bridge.OpenSession(bytes, '');
      const html = arcade.bridge.RenderHtml(handle, 'doom-', false, false, 1) as string;
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      const canvas = Array.from(parsed.querySelectorAll<HTMLElement>('p[data-anchor]'))
        .find((paragraph) => (paragraph.textContent ?? '').includes('DOOM')) ?? null;
      arcade.bridge.CloseSession(handle);
      return {
        playing: arcade.playing() as boolean,
        magic: Array.from(bytes.slice(0, 2)),
        text,
        reopenedText: canvas?.textContent ?? '',
        reopenedHtml: canvas?.innerHTML ?? '',
      };
    });

    expect(result.playing).toBe(false);
    expect(result.magic).toEqual([0x50, 0x4b]); // a real ZIP, i.e. a real .docx
    // The frame survives the round trip through the file, shading and all.
    expect(result.reopenedText).toBe(result.text);
    expect(result.reopenedHtml).toMatch(/background/i);
  });
});
