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
//
// wad= points at the copy npm/scripts/fetch-doom-iwad.mjs puts in the webroot.
// The shipped pages fetch their IWAD from a pinned CDN, but a browser suite
// that depends on a CDN being up fails for reasons that have nothing to do
// with the change under test — and the local copy is same-origin, so it goes
// through the cartridge's URL gate exactly as a self-hosted IWAD would.
const LOCAL_WAD = 'wad=' + encodeURIComponent('./vendor/freedoom1.wad.gz');
const OVERRIDE = `engine=./embed.bundle.js&intro=0&sound=0&${LOCAL_WAD}`;

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
    //
    // The margin is small on purpose. Each painted frame is a whole paragraph
    // of coloured OOXML runs through replaceXml + refresh, so the frame RATE
    // is a property of how loaded the machine is — on a busy CI runner it lands near 1 fps,
    // which made a +60 margin sit right on this timeout and flake. The claim
    // being made here is liveness, not throughput: a frozen or crashed engine
    // produces no further frames at all, so any sustained advance falsifies
    // it. Sustained real-time play is proved separately, by the turning test
    // below.
    await page.waitForFunction(
      (n) => (window as any).__arcade.game().doomFrames > n + 10,
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
      const rows = [''];
      const readRows = (node: Node) => {
        for (const child of Array.from(node.childNodes)) {
          if (child.nodeName === 'BR') rows.push('');
          else if (child.nodeType === Node.TEXT_NODE) {
            // The HTML converter emits one zero-width bidi guard after each
            // <br>; it is generated layout chrome, not a document grid cell.
            rows[rows.length - 1] += (child.textContent ?? '').replace(/[\u200e\u200f]/g, '');
          }
          else readRows(child);
        }
      };
      readRows(element);
      const controls = arcade.controlsElements() as HTMLElement[];
      return {
        text: arcade.canvasText() as string,
        rows,
        columns: rows.map((row) => row.length),
        spans: spans.length,
        shaded: shaded.length,
        inks: new Set(spans.map((span) => getComputedStyle(span).color)).size,
        controls: controls.map((line) => line.innerText).join(' '),
        controlsCount: controls.length,
        controlsFontPx: Math.min(...controls.map((line) =>
          Number.parseFloat(getComputedStyle(line).fontSize))),
        controlsOverflow: controls.some((line) => line.scrollWidth > line.clientWidth + 1),
        controlsInsideFrame: controls.some((line) => element.contains(line)),
        fallback: arcade.editor.lastReconcileFallback as string | null,
      };
    });

    // The picture uses solid and partial quadrant blocks from the live frame.
    expect(screen.text).toMatch(/█/);
    expect(screen.text).toMatch(/[▀▄]/);
    // Still the arcade's markdown-safe bezel — no row can open a heading or a
    // bullet, and no row is blank.
    expect(screen.text).toContain('┌');
    expect(screen.rows).toHaveLength(32);
    expect(screen.columns).toEqual(new Array(32).fill(96));
    for (const row of screen.rows) expect(row.trim()).not.toBe('');
    // Controls are large document text, not 8pt framebuffer cells. This is a
    // display-size contract, not merely a source-format check: even if the
    // captured editor is reduced to 60% in an embed, the keys must remain at
    // least 14px. The normal fitted desktop page renders them at 24px.
    expect(screen.controls).toContain('MOVE W/S');
    expect(screen.controls).toContain('FIRE SPACE');
    expect(screen.controls).toContain('PAUSE/EDIT ESC');
    expect(screen.controlsFontPx).toBeGreaterThanOrEqual(23.5);
    expect(screen.controlsFontPx * 0.6).toBeGreaterThanOrEqual(14);
    expect(screen.controlsCount).toBe(4);
    expect(screen.controlsOverflow).toBe(false);
    expect(screen.controlsInsideFrame).toBe(false);
    // The playable projection deliberately holds one stable high-contrast
    // endpoint pair across the picture; that is what keeps each row to one
    // rendered span while the quadrant glyphs carry the structure.
    expect(screen.inks).toBeGreaterThanOrEqual(2);
    expect(screen.inks).toBeLessThanOrEqual(3);
    expect(screen.shaded).toBeGreaterThanOrEqual(29);
    expect(screen.fallback).toBeNull();
  });

  test('the playable projection sustains ten document repaints per second', async ({ page }) => {
    test.setTimeout(300000);
    await bootDoom(page);
    await startLevel(page);

    // Warm the converter, then count COMPLETED document refreshes over a real
    // wall-clock window. Doom's internal tics do not count here: the screen has
    // to make it through replaceXml and editor.refresh() to advance `frames`.
    await page.waitForTimeout(1500);
    const start = await page.evaluate(() => ({
      frames: (window as any).__arcade.frames() as number,
      at: performance.now(),
    }));
    await page.waitForTimeout(5000);
    const result = await page.evaluate((sample) => {
      const arcade = (window as any).__arcade;
      const elapsed = performance.now() - sample.at;
      const frames = arcade.frames() - sample.frames;
      return {
        fps: frames * 1000 / elapsed,
        runs: arcade.timings().runs as number,
        spans: (arcade.canvasElement() as HTMLElement).querySelectorAll('span').length,
      };
    }, start);

    expect(result.runs).toBeLessThanOrEqual(4);
    expect(result.spans).toBeLessThan(40);
    // Half a frame of scheduling tolerance keeps the guard about the 10 fps
    // design point instead of timer quantisation at either edge of the window.
    expect(result.fps).toBeGreaterThanOrEqual(9.5);
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

  test('P switches projection, and both stay inside their run budget', async ({ page }) => {
    test.setTimeout(300000);
    await bootDoom(page);
    await startLevel(page);

    /** Spans in the canvas paragraph are the rendered form of OOXML runs, so
     *  this counts the thing the frame budget is actually made of. */
    const shape = () => page.evaluate(() => {
      const el = (window as any).__arcade.canvasElement() as HTMLElement;
      const spans = Array.from(el.querySelectorAll('span'));
      const rows = [''];
      const readRows = (node: Node) => {
        for (const child of Array.from(node.childNodes)) {
          if (child.nodeName === 'BR') rows.push('');
          else if (child.nodeType === Node.TEXT_NODE) {
            rows[rows.length - 1] += (child.textContent ?? '').replace(/[\u200e\u200f]/g, '');
          }
          else readRows(child);
        }
      };
      readRows(el);
      const shaded = spans.filter((s) => {
        const f = getComputedStyle(s).backgroundColor;
        return f && f !== 'transparent' && f !== 'rgba(0, 0, 0, 0)';
      });
      return {
        spans: spans.length,
        runs: (window as any).__arcade.timings().runs as number,
        shaded: shaded.length,
        inks: new Set(spans.map((s) => getComputedStyle(s).color)).size,
        text: (window as any).__arcade.canvasText() as string,
        rows,
        projection: (window as any).__arcade.game().projection as string,
      };
    });

    // High contrast is the default, because it is the playable projection.
    const eight = await shape();
    expect(eight.projection).toBe('8bit');
    // The whole picture shares one endpoint pair. The 29 picture rows are one
    // visible span each, while the authored XML remains three runs across all
    // line breaks; either ceiling moving means the 10 fps budget regressed.
    expect(eight.spans).toBeLessThan(40);
    expect(eight.runs).toBeLessThanOrEqual(4);
    expect(eight.shaded).toBeGreaterThanOrEqual(29);
    expect(eight.inks).toBeLessThanOrEqual(3);
    // Quadrant blocks are where the resolution comes from: they carry four
    // sub-pixels per cell instead of two, for no extra runs. Seeing them is
    // the evidence the picture is sampled at 188 x 58 rather than 94 x 58.
    expect(eight.text).toMatch(/[▘▝▖▗▌▐▚▞▛▜▙▟]/);
    // Solid quadrants only. The shade characters approximate TONE rather than
    // carrying detail, and at the shipped cell size they read as dots.
    expect(eight.text).not.toMatch(/[░▒▓]/);

    // The key is claimed by the arcade and handled inside the cartridge — it
    // must never reach Doom, which has its own meaning for most letters.
    await page.keyboard.press('KeyP');
    await page.waitForFunction(
      () => (window as any).__arcade.game().projection === 'bitmap', null, { timeout: 60000 });
    // Let a frame paint in the new projection before measuring.
    const before = await page.evaluate(() => (window as any).__arcade.frames());
    await page.waitForFunction((n) => (window as any).__arcade.frames() > n + 1, before,
      { timeout: 120000 });

    const bitmap = await shape();
    expect(bitmap.projection).toBe('bitmap');
    expect(bitmap.shaded).toBeGreaterThanOrEqual(29);   // every picture row shades
    // Edge-aware segment merging is this projection's whole frame budget:
    // without it a photographic downsample gives nearly every cell its own
    // run. The painter allocates 900 picture runs, leaving the document chrome
    // inside this established rendered-span ceiling.
    expect(bitmap.spans).toBeLessThan(1200);
    // The detailed projection derives unrestricted endpoints from the
    // framebuffer, so its palette is open where the 8-bit one is closed.
    expect(bitmap.inks).toBeGreaterThan(eight.inks);
    // The old fixed-tolerance merge satisfied the span ceiling by turning
    // near-colour wall and floor texture into horizontal bars. Ink/shading
    // boundaries cost; glyph changes do not. The replacement spends that free
    // channel so neighbouring half-pixels can keep choosing between each
    // segment's endpoints. A live E1M1 frame carries thousands of those texture
    // decisions and both orientations — this is the regression guard for the
    // actual visible picture, not merely its dimensions and run count.
    const glyphTransitions = bitmap.rows.slice(2, -1).reduce((total, row) => {
      const picture = row.slice(1, -1);
      for (let x = 1; x < picture.length; x++) total += Number(picture[x] !== picture[x - 1]);
      return total;
    }, 0);
    expect(glyphTransitions).toBeGreaterThan(500);
    expect(bitmap.text).toContain('▄');
    // Still a document, still the markdown-safe bezel.
    expect(bitmap.text).toContain('┌');
    expect(bitmap.rows).toHaveLength(32);
    for (const row of bitmap.rows) {
      expect(row).toHaveLength(96);
      expect(row.trim()).not.toBe('');
    }

    // And back again.
    await page.keyboard.press('KeyP');
    await page.waitForFunction(
      () => (window as any).__arcade.game().projection === '8bit', null, { timeout: 60000 });
  });

  test('the paused frame copies and pastes as a real block, runs and shading intact', async ({ page }) => {
    test.setTimeout(300000);
    await bootDoom(page);
    await startLevel(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 4, null, { timeout: 60000 });

    // Pause first: the frame has to stop moving before it can be selected, and
    // pausing is the cabinet's whole claim — it was a document the whole time.
    await page.evaluate(() => (window as any).__arcade.pause());
    await page.waitForTimeout(500);

    const before = await page.evaluate(() =>
      (window as any).__arcade.editor.root.querySelectorAll('[data-anchor]').length as number);

    // Select the screen paragraph and copy it, exactly as a reader would.
    await page.evaluate(() => {
      const element = (window as any).__arcade.canvasElement() as HTMLElement;
      const range = document.createRange();
      range.selectNodeContents(element);
      const selection = getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
    });
    await page.keyboard.press('Control+C');
    await page.waitForTimeout(300);

    // Paste with the caret in the last block. The editor commits TEXT diffs, so
    // a native paste would drop the colour; the cabinet inserts the paragraph's
    // own OOXML instead, which is what this is checking for.
    await page.evaluate(() => {
      const blocks = (window as any).__arcade.editor.root.querySelectorAll('[data-anchor]');
      const last = blocks[blocks.length - 1] as HTMLElement;
      const range = document.createRange();
      range.selectNodeContents(last);
      range.collapse(false);
      const selection = getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
    });
    await page.keyboard.press('Control+V');
    await page.waitForFunction((n) =>
      (window as any).__arcade.editor.root.querySelectorAll('[data-anchor]').length > n,
    before, { timeout: 60000 });

    const pasted = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      const screen = arcade.canvasElement() as HTMLElement;
      const blocks = arcade.editor.root.querySelectorAll('[data-anchor]') as NodeListOf<HTMLElement>;
      const copy = Array.from(blocks)
        .filter((block) => block !== screen && block.querySelectorAll('span').length > 25)
        .pop();
      const spans = Array.from(copy?.querySelectorAll('span') ?? []) as HTMLElement[];
      return {
        text: copy?.textContent ?? '',
        screenText: screen.textContent ?? '',
        spans: spans.length,
        shaded: spans.filter((span) => {
          const fill = getComputedStyle(span).backgroundColor;
          return fill && fill !== 'transparent' && fill !== 'rgba(0, 0, 0, 0)';
        }).length,
        inks: new Set(spans.map((span) => getComputedStyle(span).color)).size,
      };
    });

    // The copy is the frame: same characters, and the runs came across as runs.
    expect(pasted.text).toBe(pasted.screenText);
    expect(pasted.shaded).toBeGreaterThanOrEqual(29);
    expect(pasted.inks).toBeGreaterThanOrEqual(2);

    // And it is in the document, not just in the DOM — it survives a save and
    // reopen with its shading, which is the difference between a real block and
    // a browser paste the next commit would flatten.
    const saved = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      const bytes: Uint8Array = arcade.save();
      const handle = arcade.bridge.OpenSession(bytes, '');
      const html = arcade.bridge.RenderHtml(handle, 'doom-', false, false, 1) as string;
      arcade.bridge.CloseSession(handle);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      const screens = Array.from(parsed.querySelectorAll<HTMLElement>('p[data-anchor]'))
        .filter((paragraph) => (paragraph.textContent ?? '').includes('DOOM'));
      return { count: screens.length, html: screens.map((s) => s.innerHTML).join('') };
    });
    expect(saved.count).toBeGreaterThanOrEqual(2);
    expect(saved.html).toMatch(/background/i);
  });

  test('a cross-origin IWAD is refused rather than fetched', async ({ page }) => {
    test.setTimeout(120000);
    // `import()` executes whatever it fetches, on this page's origin, so the
    // cartridge takes no engine URL from a link at all and gates the one URL
    // it does take. This is the regression guard on that: a crafted link must
    // land in the cartridge's error state without a request going out.
    let requested = false;
    await page.route('https://example.com/**', (route) => {
      requested = true;
      return route.abort();
    });

    // Deliberately NOT the OVERRIDE above: this one supplies the hostile wad=
    // as the only IWAD, so nothing else can satisfy the load.
    await page.goto(
      '/demo-arcade.html?engine=./embed.bundle.js&intro=0&sound=0&cart=doom'
      + `&wad=${encodeURIComponent('https://example.com/evil.wad.gz')}`,
    );
    await page.waitForFunction(
      () => (window as any).__arcade !== undefined || (window as any).__arcadeError !== undefined,
      null,
      { timeout: 120000 },
    );
    await page.waitForFunction(
      () => (window as any).__arcade.game().status === 'error',
      null,
      { timeout: 60000 },
    );

    const state = await page.evaluate(() => (window as any).__arcade.game());
    expect(state.error).toContain('same-origin');
    expect(requested, 'no request should have been made to the cross-origin host').toBe(false);
    // The cabinet stays a document even when the cartridge refuses to load.
    expect(await page.evaluate(() => (window as any).__arcade.canvasText() as string))
      .toContain('DOOM');
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
