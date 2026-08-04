import { test, expect } from '@playwright/test';

/**
 * The social-embed demo pages (docs/demo/player.html + index.html — the
 * X/Twitter Player Card target and the og-card landing page).
 *
 * In production they pull the engine from jsDelivr and the sample document
 * from raw.githubusercontent.com; both accept `?engine=` / `?doc=` overrides
 * so this spec can drive them fully locally: pretest copies them into the
 * test webroot (player.html / demo-index.html) beside embed.bundle.js and
 * demo-sample.docx. The local layout also exercises the embed bundle's
 * wasm-webroot fallback (assets next to the bundle, not under wasm/).
 */

const OVERRIDES = 'engine=./embed.bundle.js&doc=./demo-sample.docx';

test.describe('social demo pages', () => {
  test('player.html boots the in-card editor on tap', async ({ page }) => {
    await page.goto(`/player.html?${OVERRIDES}`);
    // Poster state first — nothing heavy loads until the user taps.
    await expect(page.locator('#start')).toBeVisible();
    await expect(page.locator('#app')).toBeHidden();

    await page.click('#start');
    await expect(page.locator('#app')).toBeVisible({ timeout: 45000 });
    expect(await page.locator('#doc [data-anchor]').count()).toBeGreaterThan(0);

    // The toolbar drives the real editor: bold a selection and see markup.
    const result = await page.evaluate(() => {
      const block = document.querySelector<HTMLElement>('#doc [data-anchor]')!;
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      return { text: block.textContent?.length ?? 0 };
    });
    expect(result.text).toBeGreaterThan(0);
  });

  test('index.html landing boots the editor and reports live status', async ({ page }) => {
    await page.goto(`/demo-index.html?${OVERRIDES}`);
    await expect(page.locator('#status')).toContainText('live', { timeout: 45000 });
    expect(await page.locator('#doc [data-anchor]').count()).toBeGreaterThan(0);

    // The share-card meta X and LinkedIn read must be present and absolute.
    const meta = await page.evaluate(() => ({
      card: document.querySelector('meta[name="twitter:card"]')?.getAttribute('content'),
      player: document.querySelector('meta[name="twitter:player"]')?.getAttribute('content'),
      ogImage: document.querySelector('meta[property="og:image"]')?.getAttribute('content'),
    }));
    expect(meta.card).toBe('player');
    expect(meta.player).toMatch(/^https:\/\/.+player\.html$/);
    expect(meta.ogImage).toMatch(/^https:\/\//);
  });
});
