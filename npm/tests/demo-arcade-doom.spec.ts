import { test, expect, Page } from '@playwright/test';

// Proof that the ACTUAL Doom engine runs inside a live Word document.
//
// Cartridge 4 of THE DOCX ARCADE grew out of a hand-written ASCII raycaster
// walking Freedoom's E1M1 rasterized to a character grid. It is now id
// Software's own engine — doomgeneric, GPL-2.0, compiled to JavaScript — on
// Freedoom's BSD-licensed IWAD, with its lossless 320×200 framebuffer stored
// every frame as the media payload of one inline image in a Word paragraph.
//
// That change moves what a spec can honestly claim. The old one steered a BFS
// autopilot through geometry it could read out of the grid; this one cannot,
// because Doom's world is BSP data in a WebAssembly heap and the only thing
// the document holds is a picture of it. So this spec proves the four things
// that actually matter about the swap:
//
//   1. it is really Doom     — the engine identifies its own IWAD and its own
//                              frame counter runs;
//   2. the picture is in the DOCUMENT — a mutable image occurrence and PNG
//                              relationship, repainted incrementally one block
//                              at a time, never an overlay;
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
    // through PNG encoding, replaceImage, and refresh, so the frame RATE is a
    // property of how loaded the machine is,
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

  test('the lossless framebuffer is a legible inline DOCX image, one block repainted', async ({ page }) => {
    test.setTimeout(240000);
    await bootDoom(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 4, null, { timeout: 60000 });

    await page.evaluate(() => (window as any).__arcade.pause());
    const screen = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      const element = arcade.canvasElement() as HTMLElement;
      const image = element.querySelector('img')!;
      const rect = image.getBoundingClientRect();
      const sample = document.createElement('canvas');
      sample.width = 320; sample.height = 200;
      const context = sample.getContext('2d')!;
      context.drawImage(image, 0, 0, 320, 200);
      const rgba = context.getImageData(0, 0, 320, 200).data;
      let mismatches = 0;
      const colors = new Set<string>();
      const chromaticColors = new Set<string>();
      for (let y = 0; y < 200; y += 10) {
        for (let x = 0; x < 320; x += 10) {
          const expected = arcade.game().pixel(x, y)!;
          const o = (y * 320 + x) * 4;
          const actual = [rgba[o], rgba[o + 1], rgba[o + 2]];
          if (actual.some((channel, i) => channel !== expected[i])) mismatches++;
          colors.add(actual.join(','));
          if (Math.max(...actual) - Math.min(...actual) > 12) chromaticColors.add(actual.join(','));
        }
      }
      const controls = arcade.controlsElements() as HTMLElement[];
      const occurrences = arcade.session.listImages();
      return {
        source: image.src.slice(0, 32),
        alt: image.alt,
        complete: image.complete,
        naturalWidth: image.naturalWidth,
        naturalHeight: image.naturalHeight,
        renderedWidth: rect.width,
        renderedHeight: rect.height,
        hudHeight: rect.height * 32 / 200,
        digitHeight: rect.height * 11 / 200,
        mismatches,
        sampledColors: colors.size,
        chromaticColors: chromaticColors.size,
        images: element.querySelectorAll('img').length,
        spans: element.querySelectorAll('span').length,
        canvasesInDocument: arcade.editor.root.querySelectorAll('canvas').length,
        occurrences,
        controls: controls.map((line) => line.innerText).join(' '),
        controlsCount: controls.length,
        controlsFontPx: Math.min(...controls.map((line) =>
          Number.parseFloat(getComputedStyle(line).fontSize))),
        controlsOverflow: controls.some((line) => line.scrollWidth > line.clientWidth + 1),
        controlsInsideFrame: controls.some((line) => element.contains(line)),
        fallback: arcade.editor.lastReconcileFallback as string | null,
      };
    });

    expect(screen.source).toContain('data:image/png;base64,');
    expect(screen.alt).toContain('Live Doom framebuffer');
    expect(screen.complete).toBe(true);
    expect([screen.naturalWidth, screen.naturalHeight]).toEqual([320, 200]);
    expect(screen.renderedWidth).toBeGreaterThanOrEqual(590);
    expect(screen.renderedHeight).toBeGreaterThanOrEqual(440);
    // Guidepost: Doom's 32-pixel HUD and 11-pixel numerals must be physically
    // readable in the actual document, not merely present in source pixels.
    expect(screen.hudHeight).toBeGreaterThanOrEqual(70);
    expect(screen.digitHeight).toBeGreaterThanOrEqual(24);
    // The document image is pixel-exact against the engine framebuffer and
    // retains a real palette; a grayscale/downsampled projection cannot pass.
    expect(screen.mismatches).toBe(0);
    expect(screen.sampledColors).toBeGreaterThan(64);
    expect(screen.chromaticColors).toBeGreaterThan(20);
    expect(screen.images).toBe(1);
    expect(screen.spans).toBe(1);
    expect(screen.canvasesInDocument).toBe(0);
    expect(screen.occurrences).toHaveLength(1);
    expect(screen.occurrences[0].contentType).toBe('image/png');
    expect(screen.occurrences[0].isBroken).toBe(false);
    // Controls are large document text, not framebuffer pixels. This is a
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
    expect(screen.fallback).toBeNull();
  });

  test('turning sustains ten decoded, visibly presented document frames per second', async ({ page }) => {
    test.setTimeout(300000);
    await bootDoom(page);
    await startLevel(page);

    // Warm the converter, then count both completed session/editor refreshes
    // and distinct decoded <img> sources observed on animation frames. Doom's
    // internal tics and repeated static PNGs do not count.
    await page.waitForTimeout(1500);
    const result = await page.evaluate(async () => {
      const arcade = (window as any).__arcade;
      const first = arcade.frames();
      const started = performance.now();
      let visible = 0;
      let priorSource = '';
      arcade.input.set('ArrowRight', true);
      while (performance.now() - started < 5000) {
        await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
        const image = arcade.canvasElement()?.querySelector('img') as HTMLImageElement | null;
        if (image?.complete && image.naturalWidth === 320 && image.src !== priorSource) {
          priorSource = image.src;
          visible++;
        }
      }
      arcade.input.set('ArrowRight', false);
      const elapsed = performance.now() - started;
      const frames = arcade.frames() - first;
      const element = arcade.canvasElement() as HTMLElement;
      return {
        fps: frames * 1000 / elapsed,
        visibleFps: visible * 1000 / elapsed,
        runs: arcade.timings().runs as number,
        spans: element.querySelectorAll('span').length,
        images: element.querySelectorAll('img').length,
      };
    });

    expect(result.runs).toBe(1);
    expect(result.spans).toBe(1);
    expect(result.images).toBe(1);
    // Half a frame of scheduling tolerance keeps the guard about the 10 fps
    // design point instead of timer quantisation at either edge of the window.
    expect(result.fps).toBeGreaterThanOrEqual(9.5);
    expect(result.visibleFps).toBeGreaterThanOrEqual(9.5);
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

  test('the paused inline frame copies, pastes, saves, and reopens as a real image block', async ({ page }) => {
    test.setTimeout(300000);
    await bootDoom(page);
    await startLevel(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 4, null, { timeout: 60000 });

    // Leave several visibly distinct turns in the session history. Each
    // replaceImage is one normal undo unit, which makes the paused document a
    // tiny frame scrubber as well as a still image.
    const turnFrom = await page.evaluate(() => (window as any).__arcade.frames());
    await page.keyboard.down('ArrowRight');
    await page.waitForFunction(
      (n) => (window as any).__arcade.frames() >= n + 5,
      turnFrom,
      { timeout: 60000 },
    );
    await page.keyboard.up('ArrowRight');

    // Pause first: the frame has to stop moving before it can be selected, and
    // pausing is the cabinet's whole claim — it was a document the whole time.
    await page.evaluate(() => (window as any).__arcade.pause());
    await page.waitForTimeout(500);

    const scrubbed = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      const source = () => (arcade.canvasElement() as HTMLElement)
        .querySelector<HTMLImageElement>('img')!.src;
      const forward = source();
      arcade.editor.undo();
      const backward = source();
      (window as any).__scrubbedDoomSource = backward;
      (window as any).__forwardDoomSource = forward;
      return {
        movedBackward: backward !== forward,
        imageCount: arcade.session.listImages().length,
      };
    });
    expect(scrubbed.movedBackward, 'Undo should reveal the prior framebuffer').toBe(true);
    expect(scrubbed.imageCount).toBe(1);

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

    // Copy captured the backward frame. Redo now restores the later one, so a
    // pasted copy can prove it came from what the editor displayed at copy
    // time rather than from the game loop's last-frame cache.
    const movedForward = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      arcade.editor.redo();
      return (arcade.canvasElement() as HTMLElement)
        .querySelector<HTMLImageElement>('img')!.src === (window as any).__forwardDoomSource;
    });
    expect(movedForward, 'Redo should restore the later framebuffer').toBe(true);

    // Paste with the caret in the last block. The editor commits TEXT diffs, so
    // a native paste would flatten/drop document structure; the cabinet inserts
    // a fresh native drawing through the public image API instead.
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
        .filter((block) => block !== screen && block.querySelector('img[alt^="Live Doom framebuffer"]'))
        .pop();
      const screenImage = screen.querySelector('img') as HTMLImageElement;
      const copiedImage = copy?.querySelector('img') as HTMLImageElement | null;
      return {
        copied: !!copy,
        imageCount: copy?.querySelectorAll('img').length ?? 0,
        sameSource: copiedImage?.src === screenImage.src,
        matchesCopied: copiedImage?.src === (window as any).__scrubbedDoomSource,
        natural: [copiedImage?.naturalWidth, copiedImage?.naturalHeight],
      };
    });

    expect(pasted.copied).toBe(true);
    expect(pasted.imageCount).toBe(1);
    expect(pasted.sameSource, 'the copied older frame must differ from the redone live frame').toBe(false);
    expect(pasted.matchesCopied, 'paste must preserve the frame visible when Copy ran').toBe(true);
    expect(pasted.natural).toEqual([320, 200]);

    // And it is in the document, not just in the DOM — it survives a save and
    // reopen as embedded PNG data, which is the difference between a real block
    // and a browser-only paste the next commit would flatten.
    const saved = await page.evaluate(() => {
      const arcade = (window as any).__arcade;
      const bytes: Uint8Array = arcade.save();
      const handle = arcade.bridge.OpenSession(bytes, '');
      const html = arcade.bridge.RenderHtml(handle, 'doom-', false, false, 1) as string;
      arcade.bridge.CloseSession(handle);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      const images = Array.from(parsed.querySelectorAll<HTMLImageElement>(
        'img[alt^="Live Doom framebuffer"]'));
      return { count: images.length, sources: images.map((image) => image.src.slice(0, 32)) };
    });
    expect(saved.count).toBeGreaterThanOrEqual(2);
    expect(saved.sources.every((source: string) => source.includes('data:image/png;base64,'))).toBe(true);
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
      const source = (arcade.canvasElement() as HTMLElement).querySelector('img')?.getAttribute('src') ?? '';
      const bytes: Uint8Array = arcade.save();
      const handle = arcade.bridge.OpenSession(bytes, '');
      const html = arcade.bridge.RenderHtml(handle, 'doom-', false, false, 1) as string;
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      const image = parsed.querySelector<HTMLImageElement>('img[alt^="Live Doom framebuffer"]');
      arcade.bridge.CloseSession(handle);
      return {
        playing: arcade.playing() as boolean,
        magic: Array.from(bytes.slice(0, 2)),
        source: source.slice(0, 32),
        reopenedSource: image?.src.slice(0, 32) ?? '',
        reopenedAlt: image?.alt ?? '',
      };
    });

    expect(result.playing).toBe(false);
    expect(result.magic).toEqual([0x50, 0x4b]); // a real ZIP, i.e. a real .docx
    // The frame survives the round trip through the file as embedded PNG data.
    expect(result.source).toContain('data:image/png;base64,');
    expect(result.reopenedSource).toContain('data:image/png;base64,');
    expect(result.reopenedAlt).toContain('Live Doom framebuffer');
  });
});
