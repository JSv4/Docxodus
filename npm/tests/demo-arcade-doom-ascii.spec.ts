import { test, expect, Page } from '@playwright/test';

async function boot(page: Page, rendering = 'image') {
  await page.goto('/demo-arcade.html?engine=./embed.bundle.js&intro=0&sound=0&cart=doom'
    + `&wad=./vendor/freedoom1.wad.gz&render=${rendering}`);
  await page.waitForFunction(() => (window as any).__arcade?.game().doomFrames >= 4,
    null, { timeout: 180000 });
  await page.selectOption('#pace', '0');
  for (let i = 0; i < 4; i++) {
    await page.keyboard.press('Enter');
    await page.waitForTimeout(700);
  }
  await page.waitForTimeout(800);
}

test.describe('DOOM as native ASCII document text', () => {
  test.setTimeout(240000);

  test('projects the same paused frame, preserves every cell, and round-trips as text', async ({ page }) => {
    await boot(page);
    await page.evaluate(() => (window as any).__arcade.pause());
    const before = await page.evaluate(() => (window as any).__arcade.game().doomFrames);
    await page.selectOption('#rendering', 'ascii');
    const proof = await page.evaluate(async () => {
      const a = (window as any).__arcade;
      const { asciiFramebuffer } = await import(/* @vite-ignore */ '/doom-ascii.js' as string);
      const { rowsFromXml } = await import(/* @vite-ignore */ '/ascii-arcade.js' as string);
      const fb = new Uint8Array(320 * 200 * 4);
      for (let y = 0; y < 200; y++) for (let x = 0; x < 320; x++) {
        const [r, g, b] = a.game().pixel(x, y);
        fb.set([b, g, r, 255], (y * 320 + x) * 4);
      }
      const expected = asciiFramebuffer(fb).chars.map((r: string[]) => r.join(''));
      const xml = a.session.raw.getXml(a.canvasAnchor());
      const rows = rowsFromXml(xml);
      const el = a.canvasElement() as HTMLElement;
      const saved = a.save();
      const handle = a.bridge.OpenSession(saved, '');
      const html = a.bridge.RenderHtml(handle, 'ascii-', false, false, 1);
      a.bridge.CloseSession(handle);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      const reopened = Array.from(parsed.querySelectorAll('p')).find(p =>
        p.textContent!.replace(/\u00a0/g, ' ').length === 64200);
      const spans = Array.from(el.querySelectorAll('span'));
      const colors = new Set(spans.map(s => getComputedStyle(s).color));
      return {
        rows: rows.length, width: rows.every((r: string) => r.length === 321),
        exact: JSON.stringify(rows) === JSON.stringify(expected),
        ascii: rows.every((r: string) => /^[\x20-\x7e]+$/.test(r)),
        images: el.querySelectorAll('img, canvas, svg').length,
        drawings: xml.includes('w:drawing'), colors: colors.size,
        perCellShading: (xml.match(/<w:shd/g) ?? []).length,
        sameReopened: reopened?.textContent?.replace(/\u00a0/g, ' ') === rows.join(''),
        breaks: reopened?.querySelectorAll('br').length,
        doomFrames: a.game().doomFrames, playing: a.playing(),
        fallback: a.editor.lastReconcileFallback,
        domText: el.textContent!.replace(/\u00a0/g, ' ') === rows.join(''),
      };
    });
    expect(proof).toMatchObject({ rows: 200, width: true, exact: true, ascii: true,
      images: 0, drawings: false, perCellShading: 1, sameReopened: true,
      breaks: 199, doomFrames: before, playing: false, fallback: null, domText: true });
    expect(proof.colors).toBeGreaterThan(5);

    // Each direction of the renderer switch is one document history entry.
    await page.evaluate(() => (window as any).__arcade.editor.undo());
    expect(await page.evaluate(() => (window as any).__arcade.canvasElement().querySelectorAll('img').length)).toBe(1);
    await page.evaluate(() => (window as any).__arcade.editor.redo());
    expect(await page.evaluate(() => (window as any).__arcade.canvasElement().textContent.length)).toBe(64200);
    await page.selectOption('#rendering', 'image');
    await page.evaluate(() => (window as any).__arcade.editor.undo());
    expect(await page.evaluate(() => (window as any).__arcade.canvasElement().textContent.length)).toBe(64200);
    // Resuming after undoing an image switch must not reuse its deleted ID.
    const frames = await page.evaluate(() => {
      const a = (window as any).__arcade;
      const before = a.frames(); a.resume(); return before;
    });
    await page.waitForFunction(n => (window as any).__arcade.frames() >= n + 3, frames);
    expect(await page.evaluate(() => (window as any).__arcade.canvasElement().querySelectorAll('img').length)).toBe(1);
    await page.evaluate(() => (window as any).__arcade.pause());
  });

  test('accepts gameplay input, copies the scrubbed frame, and preserves the pasted text', async ({ page }) => {
    await boot(page, 'ascii');
    await page.keyboard.down('ArrowRight');
    await page.waitForTimeout(1500);
    await page.keyboard.up('ArrowRight');
    await page.evaluate(() => (window as any).__arcade.pause());
    const result = await page.evaluate(() => {
      const a = (window as any).__arcade;
      const current = a.canvasElement().textContent;
      a.editor.undo();
      const screen = a.canvasElement();
      const older = screen.textContent;
      const range = document.createRange(); range.selectNodeContents(screen);
      const selection = getSelection()!; selection.removeAllRanges(); selection.addRange(range);
      const clipboardData = new DataTransfer();
      screen.dispatchEvent(new ClipboardEvent('copy', { clipboardData, bubbles: true, cancelable: true }));
      const plain = clipboardData.getData('text/plain');
      a.editor.redo();
      const last = Array.from(a.editor.root.querySelectorAll('[data-anchor]')).pop() as HTMLElement;
      last.dispatchEvent(new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true }));
      const copies = Array.from(a.editor.root.querySelectorAll('[data-anchor]')) as HTMLElement[];
      const pasted = copies.find(e => e !== a.canvasElement() && e.textContent === older);
      return {
        moved: current !== older,
        plainAscii: /^[\x20-\x7e\n]+$/.test(plain), lines: plain.split('\n').length,
        copiedOlder: plain.replace(/\n/g, '') === older.replace(/\u00a0/g, ' '),
        pasted: !!pasted, pastedImages: pasted?.querySelectorAll('img,svg,canvas').length,
        replayed: a.canvasElement().textContent === current,
      };
    });
    expect(result).toEqual({ moved: true, plainAscii: true, lines: 200, copiedOlder: true,
      pasted: true, pastedImages: 0, replayed: true });
  });
});

test.describe('DOOM ASCII on a phone', () => {
  test.use({ viewport: { width: 393, height: 851 }, isMobile: true,
    hasTouch: true, deviceScaleFactor: 2 });
  test('keeps all rows on one line and exposes the renderer in the controls sheet', async ({ page }) => {
    test.setTimeout(240000);
    await boot(page, 'ascii');
    await page.evaluate(() => (window as any).__arcade.pause());
    const geometry = await page.evaluate(() => {
      const el = (window as any).__arcade.canvasElement() as HTMLElement;
      const range = document.createRange(); range.selectNodeContents(el);
      const box = el.getBoundingClientRect();
      const rects = Array.from(range.getClientRects());
      return { breaks: el.querySelectorAll('br').length, chars: el.textContent!.length,
        pre: getComputedStyle(el).whiteSpace,
        overflow: Math.max(...rects.map(r => r.right)) - box.right };
    });
    expect(geometry).toMatchObject({ breaks: 199, chars: 64200, pre: 'pre' });
    expect(geometry.overflow).toBeLessThan(2);
    await page.locator('#dockmore').click();
    await expect(page.locator('#rendering')).toBeVisible();
    await page.selectOption('#rendering', 'image');
    expect(await page.evaluate(() => (window as any).__arcade.canvasElement().querySelectorAll('img').length)).toBe(1);
    await page.selectOption('#rendering', 'ascii');
    expect(await page.evaluate(() => (window as any).__arcade.canvasElement().textContent.length)).toBe(64200);
  });
});
