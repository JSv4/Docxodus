import { test, expect, type Page } from '@playwright/test';

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

async function formattingState(page: Page, anchor: string, text: string) {
  return page.evaluate(({ anchor, text }) => {
    const block = Array.from(document.querySelectorAll<HTMLElement>('#doc [data-anchor]'))
      .find((candidate) => candidate.getAttribute('data-anchor') === anchor);
    if (!block) return { found: false, bold: false, italic: false, textPreserved: false };
    const elements = [block, ...Array.from(block.querySelectorAll<HTMLElement>('*'))];
    const bold = elements.some((element) => {
      const weight = getComputedStyle(element).fontWeight;
      return weight === 'bold' || Number.parseInt(weight, 10) >= 600;
    });
    const italic = elements.some((element) => getComputedStyle(element).fontStyle === 'italic');
    return { found: true, bold, italic, textPreserved: block.textContent?.trim() === text };
  }, { anchor, text });
}

async function selectUnformattedBlock(page: Page) {
  return page.evaluate(() => {
    const isBold = (root: HTMLElement) =>
      [root, ...Array.from(root.querySelectorAll<HTMLElement>('*'))].some((element) => {
        const weight = getComputedStyle(element).fontWeight;
        return weight === 'bold' || Number.parseInt(weight, 10) >= 600;
      });
    const block = Array.from(
      document.querySelectorAll<HTMLElement>('#doc [data-anchor][contenteditable="true"]'),
    ).find((candidate) => (candidate.textContent?.trim().length ?? 0) > 0 && !isBold(candidate));
    if (!block) throw new Error('No non-bold editable block found');
    const range = document.createRange();
    range.selectNodeContents(block);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    return {
      anchor: block.getAttribute('data-anchor')!,
      text: block.textContent?.trim() ?? '',
    };
  });
}

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
    const result = await selectUnformattedBlock(page);
    expect(result.text.length).toBeGreaterThan(0);

    await page.locator('#bar [data-fmt="bold"]').click();
    expect(await formattingState(page, result.anchor, result.text)).toMatchObject({
      found: true, bold: true, textPreserved: true,
    });

    await page.locator('#undo').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).bold).toBe(false);
    await page.locator('#redo').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).bold).toBe(true);

    const downloadPromise = page.waitForEvent('download');
    await page.locator('#dl').click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toBe('docxodus-demo.docx');
  });

  test('index.html landing boots the editor and reports live status', async ({ page }) => {
    await page.goto(`/demo-index.html?${OVERRIDES}`);
    expect(await page.locator('script[type="module"]').textContent()).toContain(RELEASE_ENGINE);
    await expect(page.locator('#loader')).toBeVisible();
    await expect(page.locator('#featureTitle')).not.toBeEmpty();
    await expect(page.locator('#status')).toContainText(/live/i, { timeout: 45000 });
    await expect(page.locator('#loader')).toBeHidden();
    expect(await page.locator('#doc [data-anchor]').count()).toBeGreaterThan(0);

    // Main-page controls are not decorative: formatting, page view, and save all
    // call the live DocxEditor instance.
    const result = await selectUnformattedBlock(page);
    await page.locator('#bar [data-fmt="italic"]').click();
    expect(await formattingState(page, result.anchor, result.text)).toMatchObject({
      found: true, italic: true, textPreserved: true,
    });

    await page.locator('#pages').click();
    await expect(page.locator('#pages')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.locator('#doc #pagination-container')).toBeVisible();
    await page.locator('#pages').click();
    await expect(page.locator('#pages')).toHaveAttribute('aria-pressed', 'false');

    const downloadPromise = page.waitForEvent('download');
    await page.locator('#dl').click();
    expect((await downloadPromise).suggestedFilename()).toBe('docxodus-demo.docx');

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
