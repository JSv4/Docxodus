import { test, expect, Page } from '@playwright/test';

// examples/ascii-animation-editor.html — the Observatory hosted by the REAL
// ribbon editor. The animation drives the editor's own session (Unid-preserving
// raw.replaceXml on the canvas paragraph) and repaints through the public
// DocxEditor.refresh(), so pausing needs no mode switch: the canvas is an
// ordinary editable block, the ribbon works on it, and Save is the editor's own.
// These tests pin that contract: frames advance INCREMENTALLY on a stable
// anchor inside the editor DOM, a paused document accepts real typing that
// coexists with the frozen frame in the saved bytes, and every scene draws.

async function bootMoneyshot(page: Page) {
  await page.goto('/ascii-animation-editor.html');
  await page.waitForFunction(
    () =>
      (window as any).__moneyshot !== undefined ||
      (window as any).__moneyshotError !== undefined,
    { timeout: 90000 },
  );
  const err = await page.evaluate(() => (window as any).__moneyshotError);
  expect(err, `editor demo boot failed: ${err}`).toBeUndefined();
}

test.describe('ASCII animation inside the real editor surface', () => {
  test('animates editor blocks incrementally on a stable anchor', async ({ page }) => {
    await bootMoneyshot(page);

    await page.waitForFunction(() => (window as any).__moneyshot.frames() >= 5, { timeout: 30000 });
    const before = await page.evaluate(() => {
      const m = (window as any).__moneyshot;
      return {
        anchor: m.canvasAnchor() as string,
        text: m.canvasText() as string,
        frames: m.frames() as number,
        fallback: m.editor.lastReconcileFallback as string | null,
      };
    });
    expect(before.anchor).toMatch(/^p:body:/);
    expect(before.text).toMatch(/[~^\\/]/);
    // The frame path must be the incremental one — a remount per frame would
    // mean refresh() lost the Unid-preserving single-block repaint.
    expect(before.fallback).toBeNull();

    await page.waitForFunction(
      (n) => (window as any).__moneyshot.frames() >= n + 3,
      before.frames,
      { timeout: 30000 },
    );
    const after = await page.evaluate(() => {
      const m = (window as any).__moneyshot;
      return { anchor: m.canvasAnchor() as string, text: m.canvasText() as string };
    });
    expect(after.anchor).toBe(before.anchor);
    expect(after.text).not.toBe(before.text);

    // The canvas is not an exhibit — it is a wired, editable block of the editor.
    const unid = before.anchor.split(':')[2];
    const canvas = page.locator(`[data-anchor="${unid}"]`);
    await expect(canvas).toHaveCount(1);
    await expect(canvas).toHaveAttribute('contenteditable', 'true');
  });

  test('pause, type into the document, save: one real DOCX holds both', async ({ page }) => {
    await bootMoneyshot(page);
    await page.waitForFunction(() => (window as any).__moneyshot.frames() >= 3, { timeout: 30000 });
    await page.evaluate(() => (window as any).__moneyshot.pause());

    // Real editing gestures on the REAL surface: click the title block, type,
    // then click the caption to commit the title on blur.
    const anchors = await page.evaluate(() => {
      const m = (window as any).__moneyshot;
      const blocks = Array.from(
        m.editor.root.querySelectorAll('[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[];
      const title = blocks.find((b) => (b.textContent ?? '').includes('THE DOCX OBSERVATORY'));
      const caption = blocks.find((b) => (b.textContent ?? '').includes('single monospaced paragraph'));
      return { title: title?.getAttribute('data-anchor'), caption: caption?.getAttribute('data-anchor') };
    });
    expect(anchors.title).toBeTruthy();
    expect(anchors.caption).toBeTruthy();

    await page.locator(`[data-anchor="${anchors.title}"]`).click();
    await page.keyboard.press('End');
    await page.keyboard.type(' — CAUGHT MID-WAVE');
    await page.locator(`[data-anchor="${anchors.caption}"]`).click(); // commits the title

    const result = await page.evaluate(() => {
      const m = (window as any).__moneyshot;
      const bytes: Uint8Array = m.save();
      const handle = m.bridge.OpenSession(bytes, '');
      const fresh = new (window as any).DocxodusSession.DocxSession(handle, m.bridge);
      const projection = fresh.project();
      fresh.close();
      return {
        magic: Array.from(bytes.slice(0, 2)),
        markdown: projection.markdown as string,
      };
    });
    expect(result.magic).toEqual([0x50, 0x4b]); // a ZIP container
    // The projection escapes '-' as '\-', so match either form.
    expect(result.markdown).toMatch(/CAUGHT MID\\?-WAVE/); // the human's edit
    expect(result.markdown).toMatch(/[~^\\/]/); // the frozen sea
    expect(result.markdown).toContain('single monospaced paragraph'); // the seeded doc
  });

  test('every scene draws inside the editor', async ({ page }) => {
    await bootMoneyshot(page);
    await page.evaluate(() => (window as any).__moneyshot.pause());

    const frames: Record<string, string> = {};
    for (const scene of ['ocean', 'ripples', 'rain', 'fire']) {
      frames[scene] = await page.evaluate((name) => {
        const m = (window as any).__moneyshot;
        m.setScene(name);
        m.step();
        return m.canvasText() as string;
      }, scene);
      expect(frames[scene].length, `${scene} drew an empty frame`).toBeGreaterThan(500);
    }
    expect(frames.ocean).toContain('~');
    expect(frames.rain).toContain('|');
    expect(frames.fire).toMatch(/[@%#]/);
    expect(frames.ocean).not.toBe(frames.fire);
  });
});
