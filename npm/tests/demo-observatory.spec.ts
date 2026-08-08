import { test, expect, Page } from '@playwright/test';

// docs/demo/observatory.html — the GitHub Pages observatory page (copied into
// the webroot as demo-observatory.html by pretest). In production it imports
// the pinned CDN engine AND the phenomena module from the same docxodus
// package; `?engine=` retargets both (ascii-scenes.js is resolved beside the
// engine bundle), which is how this spec drives the page fully locally.
// Deep animation/editing behavior is pinned by ascii-animation-editor.spec.ts;
// this spec guards the PAGE wiring: the createRibbonEditor boot path, the
// scenes-URL derivation, the dock, and the observatory running in the surface.

const OVERRIDE = 'engine=./embed.bundle.js';

async function bootPage(page: Page) {
  await page.goto(`/demo-observatory.html?${OVERRIDE}`);
  await page.waitForFunction(
    () =>
      (window as any).__moneyshot !== undefined ||
      (window as any).__moneyshotError !== undefined,
    { timeout: 90000 },
  );
  const err = await page.evaluate(() => (window as any).__moneyshotError);
  expect(err, `pages observatory boot failed: ${err}`).toBeUndefined();
}

test.describe('GitHub Pages observatory page', () => {
  test('boots the shipped surface and animates incrementally on a stable anchor', async ({ page }) => {
    await bootPage(page);

    await page.waitForFunction(() => (window as any).__moneyshot.frames() >= 4, { timeout: 30000 });
    const state = await page.evaluate(() => {
      const m = (window as any).__moneyshot;
      return {
        anchor: m.canvasAnchor() as string,
        text: m.canvasText() as string,
        fallback: m.editor.lastReconcileFallback as string | null,
        dockHidden: (document.getElementById('dock') as HTMLElement).hidden,
        sceneButtons: document.querySelectorAll('#dockscenes button').length,
      };
    });
    expect(state.anchor).toMatch(/^p:body:/);
    expect(state.text).toMatch(/[~^\\/]/);
    expect(state.fallback).toBeNull(); // per-frame path stayed incremental
    expect(state.dockHidden).toBe(false);
    expect(state.sceneButtons).toBe(4);

    // The ribbon chrome is the shipped surface, not page-local UI.
    await expect(page.locator('[data-dxr-surface], .dxr').first()).toBeVisible();
  });

  test('pause + save yields the observatory as a real DOCX', async ({ page }) => {
    await bootPage(page);
    await page.waitForFunction(() => (window as any).__moneyshot.frames() >= 2, { timeout: 30000 });

    const result = await page.evaluate(() => {
      const m = (window as any).__moneyshot;
      m.pause();
      m.step();
      const bytes: Uint8Array = m.save();
      const handle = m.bridge.OpenSession(bytes, '');
      const markdown = JSON.parse(m.bridge.Project(handle)).markdown as string;
      m.bridge.CloseSession(handle);
      return { magic: Array.from(bytes.slice(0, 2)), markdown };
    });
    expect(result.magic).toEqual([0x50, 0x4b]);
    expect(result.markdown).toContain('THE DOCX OBSERVATORY');
    expect(result.markdown).toMatch(/[~^\\/]/);
  });
});
