import { test, expect, Page } from '@playwright/test';

// docs/demo/arcade.html — THE DOCX ARCADE (copied into the webroot as
// demo-arcade.html by pretest). In production it imports the pinned CDN
// engine, while the games (`ascii-arcade.js`) are demo content living beside
// the page — in docs/demo/ on Pages, and in the webroot here (pretest copies
// the same file). `?engine=./embed.bundle.js` retargets only the library,
// which is how this spec drives the page locally.
//
// The pure game logic (physics, raycaster, level parse round-trips) is
// exercised headlessly against the module; this spec guards what only a real
// browser proves: the createRibbonEditor boot path, the per-frame
// replaceXml + refresh loop staying incremental, keyboard capture reaching
// the simulation, and the signature trick — pause, TYPE terrain into the
// document, resume, and the level re-parses from the paragraph.

// intro=0 skips the attract screen: these specs test the cartridges, and the
// intro has its own dedicated coverage.
// wad= is the webroot copy fetch-doom-iwad.mjs provides, so the Doom
// cartridge shows its real picture here rather than a failed load.
const OVERRIDE = 'engine=./embed.bundle.js&intro=0&sound=0'
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

/** Hold the named game-loop timeout after every completed frame. Calling
 * `release()` runs exactly one next frame synchronously, then captures its
 * newly scheduled timeout. This makes frame assertions genuinely consecutive. */
async function installFrameGate(page: Page) {
  await page.evaluate(() => {
    const nativeSetTimeout = window.setTimeout.bind(window);
    let pending: (() => void) | null = null;
    let sequence = 0;
    (window as any).__arcadeFrameGate = {
      sequence: () => sequence,
      release: () => {
        if (!pending) throw new Error('no arcade frame is pending');
        const next = pending;
        pending = null;
        next();
      },
    };
    window.setTimeout = ((handler: TimerHandler, timeout?: number, ...args: any[]) => {
      if (typeof handler === 'function' && handler.name === 'loop') {
        pending = () => handler(...args);
        sequence++;
        return -sequence;
      }
      return nativeSetTimeout(handler, timeout, ...args);
    }) as typeof window.setTimeout;
  });
}

/** The HUD word that proves a given cartridge's frame is on screen. */
export const CART_HUD: Record<string, string> = {
  // Doom's HUD row says so while the engine and its IWAD are still
  // downloading, which is the frame these gated specs land on.
  doom: 'DOOM',
};

async function bootGatedCartridge(page: Page, cart: 'doom') {
  await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&cart=${cart}`);
  await installFrameGate(page);
  await page.locator('#boot').click();
  await waitForBoot(page);
  await page.waitForFunction(() =>
    (window as any).__arcade.frames() === 1 &&
    (window as any).__arcadeFrameGate.sequence() === 1,
  );
}

async function replaceCanvasCharacter(page: Page, row: number, column: number, text: string) {
  return page.evaluate(({ row, column, text }) => {
    const a = (window as any).__arcade;
    a.pause();
    const el = a.canvasElement() as HTMLElement;
    el.focus();
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_ALL);
    let currentRow = 0;
    let currentColumn = 0;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (node.nodeName === 'BR') {
        currentRow++;
        currentColumn = 0;
        continue;
      }
      if (node.nodeType !== Node.TEXT_NODE || currentRow !== row) continue;
      const length = node.textContent?.length ?? 0;
      if (column < currentColumn + length) {
        const offset = column - currentColumn;
        const range = document.createRange();
        range.setStart(node, offset);
        range.setEnd(node, offset + 1);
        const selection = window.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        return document.execCommand('insertText', false, text);
      }
      currentColumn += length;
    }
    return false;
  }, { row, column, text });
}

test.describe('THE DOCX ARCADE page', () => {
  for (const cart of ['doom'] as const) {
    test(`${cart}: its very first frame reconciles incrementally`, async ({ page }) => {
      await bootGatedCartridge(page, cart);
      const state = await page.evaluate(() => {
        const a = (window as any).__arcade;
        return {
          frames: a.frames() as number,
          fallback: a.editor.lastReconcileFallback as string | null,
          notes: document.querySelectorAll('section.footnotes > ol > li').length,
          text: a.canvasText() as string,
        };
      });
      expect(state.frames).toBe(1);
      expect(state.fallback).toBeNull();
      expect(state.notes).toBe(1);
      expect(state.text).toContain(CART_HUD[cart]);
    });

    test(`${cart}: ten consecutive frame saves reopen with stable canvas content`, async ({ page }) => {
      test.setTimeout(120000);
      await bootGatedCartridge(page, cart);

      const observations: Array<{
        frame: number;
        anchor: string;
        reopenedAnchor: string | null;
        text: string;
        reopenedText: string;
        image: string;
        reopenedImage: string;
        magic: number[];
      }> = [];
      for (let i = 0; i < 10; i++) {
        observations.push(await page.evaluate((hudWord) => {
          const a = (window as any).__arcade;
          const anchor = a.canvasAnchor() as string;
          const canvas = a.canvasElement() as HTMLElement;
          const text = a.canvasText() as string;
          const image = canvas.querySelector<HTMLImageElement>('img')?.src ?? '';
          const bytes: Uint8Array = a.save();
          const reopened = a.bridge.OpenSession(bytes, '');
          const html = a.bridge.RenderHtml(reopened, 'stress-', false, false, 1) as string;
          const parsed = new DOMParser().parseFromString(html, 'text/html');
          // Quest/Dungeon and Doom's loading frame are text. Once the engine
          // starts, Doom's complete framebuffer is a native drawing, so find
          // the reopened block by its image contract rather than pretending
          // an image has HUD text nodes.
          const reopenedImageElement = image
            ? parsed.querySelector<HTMLImageElement>('img[alt^="Live Doom framebuffer"]')
            : null;
          const reopenedCanvas = reopenedImageElement?.closest<HTMLElement>('p[data-anchor]') ??
            Array.from(parsed.querySelectorAll<HTMLElement>('p[data-anchor]'))
              .find((paragraph) => (paragraph.textContent ?? '').includes(hudWord)) ?? null;
          const reopenedAnchor = reopenedCanvas?.getAttribute('data-anchor') ?? null;
          const reopenedText = reopenedCanvas?.textContent ?? '';
          const reopenedImage = reopenedImageElement?.src ?? '';
          a.bridge.CloseSession(reopened);
          return {
            frame: a.frames() as number,
            anchor,
            reopenedAnchor,
            text,
            reopenedText,
            image,
            reopenedImage,
            magic: Array.from(bytes.slice(0, 2)),
          };
        }, CART_HUD[cart]));
        if (i < 9) {
          await page.evaluate(() => (window as any).__arcadeFrameGate.release());
          await page.waitForFunction((frame) => (window as any).__arcade.frames() === frame, i + 2);
        }
      }

      expect(observations.map((o) => o.frame)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
      expect(new Set(observations.map((o) => o.anchor)).size).toBe(1);
      for (const observation of observations) {
        expect(observation.magic).toEqual([0x50, 0x4b]);
        expect(observation.reopenedAnchor).toMatch(/^[0-9a-f]{32}$/);
        if (observation.image) {
          expect(observation.reopenedImage).toBe(observation.image);
          expect(observation.reopenedImage).toContain('data:image/png;base64,');
        } else {
          expect(observation.reopenedText).toBe(observation.text);
          expect(observation.reopenedText).toContain(CART_HUD[cart]);
        }
      }
    });
  }

  test('attract screen: OS LEGAL presents DOCXODUS, and Space drops the coin', async ({ page }) => {
    // No intro=0 here — the title card animates on the same canvas paragraph
    // the games use, so the whole per-frame path is already under test.
    await page.goto('/demo-arcade.html?engine=./embed.bundle.js&cart=doom');
    await waitForBoot(page);
    // Wait past the sweep reveal until the blinking coin prompt is on screen.
    // The rendered DOM preserves space runs as NBSP+space pairs and carries
    // invisible direction marks between runs — normalize both before matching.
    await page.waitForFunction(
      () => ((window as any).__arcade.canvasText() as string)
        .replace(/[\u200B-\u200F\uFEFF]/g, '')
        .replace(/\u00A0/g, ' ')
        .includes('PRESS  SPACE'),
      null,
      { timeout: 45000 },
    );
    const state = await page.evaluate(() => {
      const a = (window as any).__arcade;
      return {
        intro: a.introActive() as boolean,
        text: (a.canvasText() as string)
          .replace(/[\u200B-\u200F\uFEFF]/g, '')
          .replace(/\u00A0/g, ' '),
        fallback: a.editor.lastReconcileFallback as string | null,
      };
    });
    expect(state.intro).toBe(true);
    expect(state.text).toContain('OS LEGAL');   // the credit line finished typing
    expect(state.text).toMatch(/█/);            // the block title is on screen
    expect(state.fallback).toBeNull();          // attract frames stay incremental
    // Space is the coin drop: the selected cartridge takes over the same canvas.
    await page.keyboard.press('Space');
    await page.waitForFunction(
      () => !(window as any).__arcade.introActive(),
      null,
      { timeout: 15000 },
    );
    await page.waitForFunction(
      () => ((window as any).__arcade.canvasText() as string).includes('DOOM'),
      null,
      { timeout: 15000 },
    );
    expect(await page.evaluate(() => (window as any).__arcade.cart())).toBe('doom');
  });
});
