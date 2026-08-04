import { test, expect } from '@playwright/test';

/**
 * The social-embed demo pages (docs/demo/player.html + index.html — an
 * experimental historical Player Card target and the Open Graph landing page).
 *
 * In production they pull the engine from jsDelivr and the sample document
 * from raw.githubusercontent.com; both accept `?engine=` / `?doc=` overrides
 * so this spec can drive them fully locally: pretest copies them into the
 * test webroot (player.html / demo-index.html) beside embed.bundle.js and
 * demo-sample.docx. The local layout also exercises the embed bundle's
 * wasm-webroot fallback (assets next to the bundle, not under wasm/).
 */

const OVERRIDES = 'engine=./embed.bundle.js&doc=./demo-sample.docx';
const RELEASE_ENGINE = 'docxodus@9.1.0/dist/embed.bundle.js';

test.describe('social demo pages', () => {
  test('player.html boots its editor on tap', async ({ page }) => {
    await page.goto(`/player.html?${OVERRIDES}`);
    expect(await page.locator('script[type="module"]').textContent()).toContain(RELEASE_ENGINE);
    // Poster state first — nothing heavy loads until the user taps.
    await expect(page.locator('#start')).toBeVisible();
    await expect(page.locator('#app')).toBeHidden();

    await page.click('#start');
    await expect(page.locator('#app')).toBeVisible({ timeout: 45000 });
    expect(await page.locator('#doc [data-anchor]').count()).toBeGreaterThan(0);

    // The toolbar drives the real editor: select an initially-normal block,
    // click Bold through the actual button handler, and verify the remounted
    // block now carries bold computed styling.
    const result = await page.evaluate(() => {
      const isBold = (root: HTMLElement) =>
        [root, ...Array.from(root.querySelectorAll<HTMLElement>('*'))].some((el) => {
          const weight = getComputedStyle(el).fontWeight;
          return weight === 'bold' || Number.parseInt(weight, 10) >= 600;
        });
      const block = Array.from(
        document.querySelectorAll<HTMLElement>('#doc [data-anchor][contenteditable="true"]'),
      ).find((candidate) => (candidate.textContent?.trim().length ?? 0) > 0 && !isBold(candidate));
      if (!block) throw new Error('No non-bold editable block found');
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      return {
        anchor: block.getAttribute('data-anchor'),
        text: block.textContent?.trim() ?? '',
      };
    });
    expect(result.text.length).toBeGreaterThan(0);

    await page.locator('#bar [data-fmt="bold"]').click();
    const formatted = await page.evaluate(({ anchor, text }) => {
      const block = Array.from(document.querySelectorAll<HTMLElement>('#doc [data-anchor]'))
        .find((candidate) => candidate.getAttribute('data-anchor') === anchor);
      if (!block) return { found: false, bold: false, textPreserved: false };
      const elements = [block, ...Array.from(block.querySelectorAll<HTMLElement>('*'))];
      const bold = elements.some((el) => {
        const weight = getComputedStyle(el).fontWeight;
        return weight === 'bold' || Number.parseInt(weight, 10) >= 600;
      });
      return { found: true, bold, textPreserved: block.textContent?.trim() === text };
    }, result);
    expect(formatted).toEqual({ found: true, bold: true, textPreserved: true });
  });

  test('index.html landing boots the editor and reports live status', async ({ page }) => {
    await page.goto(`/demo-index.html?${OVERRIDES}`);
    expect(await page.locator('script[type="module"]').textContent()).toContain(RELEASE_ENGINE);
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
