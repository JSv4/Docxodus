import { test, expect, Page } from '@playwright/test';

// examples/ascii-animation.html — DOCX as a rendering canvas. A live DocxSession
// paragraph is the framebuffer: every frame is raw.replaceXml (Unid-preserving,
// so the anchor survives) + renderBlock + one DOM replaceWith. These tests pin
// the demo's contract: it boots, frames advance on a STABLE anchor, every scene
// draws, and a mid-animation save() is a valid DOCX holding the caught frame.

async function bootObservatory(page: Page) {
  await page.goto('/ascii-animation.html');
  await page.waitForFunction(
    () =>
      (window as any).__observatory !== undefined ||
      (window as any).__observatoryError !== undefined,
    { timeout: 90000 },
  );
  const err = await page.evaluate(() => (window as any).__observatoryError);
  expect(err, `demo boot failed: ${err}`).toBeUndefined();
}

test.describe('ASCII animation demo — DOCX as a rendering canvas', () => {
  test('boots, animates, and keeps the canvas anchor stable across frames', async ({ page }) => {
    await bootObservatory(page);

    // Let a few frames land.
    await page.waitForFunction(() => (window as any).__observatory.frames() >= 5, { timeout: 30000 });

    const before = await page.evaluate(() => {
      const o = (window as any).__observatory;
      return { anchor: o.canvasAnchor(), text: o.canvasText() as string, frames: o.frames() as number };
    });
    expect(before.anchor).toMatch(/^p:body:/);
    // The ocean scene paints from frame one: surface glyphs must be present.
    expect(before.text).toMatch(/[~^\\/]/);

    // More frames → different pixels, same anchor. That pairing is the whole
    // point: replaceXml kept the Unid, so renderBlock stayed addressable.
    await page.waitForFunction(
      (n) => (window as any).__observatory.frames() >= n + 3,
      before.frames,
      { timeout: 30000 },
    );
    const after = await page.evaluate(() => {
      const o = (window as any).__observatory;
      return { anchor: o.canvasAnchor(), text: o.canvasText() as string };
    });
    expect(after.anchor).toBe(before.anchor);
    expect(after.text).not.toBe(before.text);

    // The rendered canvas block is a real data-anchor element in the DOM.
    const unid = before.anchor.split(':')[2];
    await expect(page.locator(`[data-anchor="${unid}"]`)).toHaveCount(1);
  });

  test('every scene draws its own phenomenon', async ({ page }) => {
    await bootObservatory(page);
    await page.evaluate(() => (window as any).__observatory.pause());

    const frames: Record<string, string> = {};
    for (const scene of ['ocean', 'ripples', 'rain', 'fire']) {
      frames[scene] = await page.evaluate((name) => {
        const o = (window as any).__observatory;
        o.setScene(name);
        o.step();
        return o.canvasText() as string;
      }, scene);
      expect(frames[scene].length, `${scene} drew an empty frame`).toBeGreaterThan(500);
    }

    // Distinct phenomena, distinct glyph vocabularies.
    expect(frames.ocean).toContain('~');
    expect(frames.rain).toContain('|');
    expect(frames.fire).toMatch(/[@%#]/);
    expect(frames.ocean).not.toBe(frames.fire);
    expect(frames.ripples).not.toBe(frames.rain);
  });

  test('save() mid-animation yields a real DOCX holding the caught frame', async ({ page }) => {
    await bootObservatory(page);

    const result = await page.evaluate(() => {
      const o = (window as any).__observatory;
      o.pause();
      o.setScene('ocean');
      o.step();
      const bytes: Uint8Array = o.save();

      // Round-trip: open the saved bytes as a FRESH session and project it.
      const handle = o.bridge.OpenSession(bytes, '');
      const fresh = new (window as any).DocxodusSession.DocxSession(handle, o.bridge);
      const projection = fresh.project();
      fresh.close();
      return {
        magic: Array.from(bytes.slice(0, 2)),
        size: bytes.length,
        markdown: projection.markdown as string,
      };
    });

    // A ZIP container of plausible size…
    expect(result.magic).toEqual([0x50, 0x4b]); // "PK"
    expect(result.size).toBeGreaterThan(2000);
    // …that is the document we seeded, with the frozen ocean inside it.
    expect(result.markdown).toContain('THE DOCX OBSERVATORY');
    expect(result.markdown).toContain('~');
    expect(result.markdown).toContain('single monospaced paragraph');
  });
});
