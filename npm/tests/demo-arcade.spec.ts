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

async function bootGatedCartridge(page: Page, cart: 'quest' | 'dungeon') {
  await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&cart=${cart}`);
  await installFrameGate(page);
  await page.locator('#boot').click();
  await waitForBoot(page);
  await page.waitForFunction(() =>
    (window as any).__arcade.frames() === 1 &&
    (window as any).__arcadeFrameGate.sequence() === 1,
  );
}

test.describe('THE DOCX ARCADE page', () => {
  for (const cart of ['quest', 'dungeon'] as const) {
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
      expect(state.text).toContain(cart === 'quest' ? 'PILCROW' : 'DUNGEON');
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
        magic: number[];
      }> = [];
      for (let i = 0; i < 10; i++) {
        observations.push(await page.evaluate(() => {
          const a = (window as any).__arcade;
          const anchor = a.canvasAnchor() as string;
          const text = a.canvasText() as string;
          const bytes: Uint8Array = a.save();
          const reopened = a.bridge.OpenSession(bytes, '');
          const html = a.bridge.RenderHtml(reopened, 'stress-', false, false, 1) as string;
          const parsed = new DOMParser().parseFromString(html, 'text/html');
          const reopenedCanvas = Array.from(parsed.querySelectorAll<HTMLElement>('p[data-anchor]'))
            .find((paragraph) => (paragraph.textContent ?? '')
              .includes(a.cart() === 'quest' ? 'PILCROW' : 'DUNGEON')) ?? null;
          const reopenedAnchor = reopenedCanvas?.getAttribute('data-anchor') ?? null;
          const reopenedText = reopenedCanvas?.textContent ?? '';
          a.bridge.CloseSession(reopened);
          return {
            frame: a.frames() as number,
            anchor,
            reopenedAnchor,
            text,
            reopenedText,
            magic: Array.from(bytes.slice(0, 2)),
          };
        }));
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
        expect(observation.reopenedText).toBe(observation.text);
        expect(observation.reopenedText).toContain(cart === 'quest' ? 'PILCROW' : 'DUNGEON');
      }
    });
  }

  test('boots the shipped surface, animates incrementally, and steers with the keyboard', async ({ page }) => {
    await page.goto(`/demo-arcade.html?${OVERRIDE}`);
    await waitForBoot(page);

    await page.waitForFunction(() => (window as any).__arcade.frames() >= 4, { timeout: 30000 });
    const state = await page.evaluate(() => {
      const a = (window as any).__arcade;
      return {
        anchor: a.canvasAnchor() as string,
        text: a.canvasText() as string,
        cart: a.cart() as string,
        fallback: a.editor.lastReconcileFallback as string | null,
        cartButtons: document.querySelectorAll('#dockcarts button').length,
        playerX: a.game().player.x as number,
      };
    });
    expect(state.anchor).toMatch(/^p:body:/);
    expect(state.cart).toBe('quest');
    expect(state.text).toContain('¶');            // the player is on screen
    expect(state.text).toContain('PILCROW');      // HUD present
    expect(state.text).toContain('│');            // bezel present
    expect(state.fallback).toBeNull();            // per-frame path stayed incremental
    expect(state.cartButtons).toBe(2);
    // Computed visibility, not the `hidden` attribute: the overlay/dock carry
    // explicit display values, which would silently defeat the attribute.
    await expect(page.locator('#dock')).toBeVisible();
    await expect(page.locator('#boot')).toBeHidden(); // auto-boot: no coin screen

    // Hold ArrowRight: the capture-phase input must reach the simulation.
    const before = state.playerX;
    await page.keyboard.down('ArrowRight');
    await page.waitForFunction(
      (x0) => (window as any).__arcade.game().player.x > x0 + 3,
      before,
      { timeout: 15000 },
    );
    await page.keyboard.up('ArrowRight');
    const after = await page.evaluate(() => (window as any).__arcade.game().player.x);
    expect(after).toBeGreaterThan(before + 3);

    // The ribbon chrome is the shipped surface, not page-local UI.
    await expect(page.locator('[data-dxr-surface], .dxr').first()).toBeVisible();
  });

  test('pause → type terrain into the document → resume makes it real; save yields a real DOCX', async ({ page }) => {
    await page.goto(`/demo-arcade.html?${OVERRIDE}`);
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 2, { timeout: 30000 });

    // Pause, then type $$$$ into a sky row of the frozen frame exactly as a
    // player would: caret into the contenteditable paragraph, insertText.
    const typed = await page.evaluate(() => {
      const a = (window as any).__arcade;
      a.pause();
      const el = a.canvasElement() as HTMLElement;
      el.focus();
      // Walk to display row 8 (level row 6 — open sky): rows are separated
      // by <br>, so the row starts after the 8th <br>.
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_ALL);
      let brs = 0;
      let target: Text | null = null;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (n.nodeName === 'BR') brs++;
        else if (brs === 8 && n.nodeType === Node.TEXT_NODE && (n.textContent ?? '').length > 3) {
          target = n as Text;
          break;
        }
      }
      if (!target) return { ok: false as const, reason: 'row text node not found' };
      const range = document.createRange();
      range.setStart(target, 3); // past the │ bezel column
      range.collapse(true);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      const ok = document.execCommand('insertText', false, '$$$$');
      return { ok, reason: 'execCommand' };
    });
    expect(typed.ok, typed.reason).toBe(true);

    // Resume: blur commits the edit through markdown ReplaceText, then the
    // driver re-parses the level from the session's XML.
    await page.evaluate(() => (window as any).__arcade.resume());
    const levelRow = await page.evaluate(() => (window as any).__arcade.game().levelRow(6) as string);
    expect(levelRow).toContain('§§§§'); // typed $ aliases the § coin tile

    // And the whole thing is still just a document: save + reopen + project.
    const result = await page.evaluate(() => {
      const a = (window as any).__arcade;
      a.pause();
      const bytes: Uint8Array = a.save();
      const handle = a.bridge.OpenSession(bytes, '');
      const markdown = JSON.parse(a.bridge.Project(handle)).markdown as string;
      a.bridge.CloseSession(handle);
      return { magic: Array.from(bytes.slice(0, 2)), markdown };
    });
    expect(result.magic).toEqual([0x50, 0x4b]);
    expect(result.markdown).toContain('THE DOCX ARCADE');
    expect(result.markdown).toContain('│');
  });

  test('power-on-tap embed path boots the dungeon cartridge and turns', async ({ page }) => {
    // boot=tap is what an iframe embed gets: no runtime streams until the
    // visitor inserts a coin.
    await page.goto(`/demo-arcade.html?${OVERRIDE}&boot=tap&cart=dungeon`);
    await expect(page.locator('#boot')).toBeVisible();
    expect(await page.evaluate(() => (window as any).__arcade)).toBeUndefined();

    await page.locator('#boot').click();
    await waitForBoot(page);
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 3, { timeout: 30000 });

    const state = await page.evaluate(() => {
      const a = (window as any).__arcade;
      return {
        cart: a.cart() as string,
        text: a.canvasText() as string,
        player: a.game().player as { x: number; y: number; dx: number; dy: number },
      };
    });
    expect(state.cart).toBe('dungeon');
    expect(state.text).toContain('DUNGEON');
    expect(state.text).toContain('MAP');
    expect(state.text).toMatch(/[█▓▒░]/); // the raycast view is on screen

    // W walks forward (spawn faces +x down the entry hall).
    await page.keyboard.down('KeyW');
    await page.waitForFunction(
      (x0) => (window as any).__arcade.game().player.x > x0 + 1,
      state.player.x,
      { timeout: 15000 },
    );
    await page.keyboard.up('KeyW');

    // Arrows turn: the direction vector must rotate.
    await page.keyboard.down('ArrowLeft');
    await page.waitForFunction(
      (dy0) => Math.abs((window as any).__arcade.game().player.dy - dy0) > 0.3,
      state.player.dy,
      { timeout: 15000 },
    );
    await page.keyboard.up('ArrowLeft');
  });
});
