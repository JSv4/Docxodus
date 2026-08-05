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

const OVERRIDES = 'engine=./embed.bundle.js&doc=./docxodus-demo-guide.docx';
const RELEASE_ENGINE = 'docxodus@9.1.1/dist/embed.bundle.js';

async function formattingState(page: Page, anchor: string, text: string) {
  return page.evaluate(({ anchor, text }) => {
    const block = Array.from(document.querySelectorAll<HTMLElement>('#doc [data-anchor]'))
      .find((candidate) => candidate.getAttribute('data-anchor') === anchor);
    if (!block) return {
      found: false, bold: false, italic: false, strike: false, fontSize: 0,
      alignment: '', textPreserved: false,
    };
    const elements = [block, ...Array.from(block.querySelectorAll<HTMLElement>('*'))];
    const bold = elements.some((element) => {
      const weight = getComputedStyle(element).fontWeight;
      return weight === 'bold' || Number.parseInt(weight, 10) >= 600;
    });
    const italic = elements.some((element) => getComputedStyle(element).fontStyle === 'italic');
    const strike = elements.some((element) => getComputedStyle(element).textDecorationLine.includes('line-through'));
    const fontSize = Math.max(...elements.map((element) => Number.parseFloat(getComputedStyle(element).fontSize)));
    const alignment = getComputedStyle(block).textAlign;
    return { found: true, bold, italic, strike, fontSize, alignment, textPreserved: block.textContent?.trim() === text };
  }, { anchor, text });
}

async function selectPracticeBlock(page: Page) {
  return page.evaluate(() => {
    const block = Array.from(
      document.querySelectorAll<HTMLElement>('#doc [data-anchor][contenteditable="true"]'),
    ).find((candidate) => candidate.textContent?.includes('This is a real Word document'));
    if (!block) throw new Error('The editable style-playground block was not rendered');
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

async function reselectBlock(page: Page, anchor: string) {
  await page.evaluate((anchor) => {
    const block = document.querySelector<HTMLElement>(`#doc [data-anchor="${CSS.escape(anchor)}"]`);
    if (!block) throw new Error(`Block ${anchor} was not remounted`);
    const range = document.createRange();
    range.selectNodeContents(block);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  }, anchor);
}

async function dragAcrossDemoBlocks(page: Page, firstNeedle: string, lastNeedle: string) {
  const editable = page.locator('#doc [data-anchor][contenteditable="true"]');
  const first = editable.filter({ hasText: firstNeedle }).first();
  const last = editable.filter({ hasText: lastNeedle }).first();
  await first.scrollIntoViewIfNeeded();
  const textPoint = (element: HTMLElement, fraction: number) => {
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      if (!(node.textContent ?? '').trim()) continue;
      const range = document.createRange();
      range.selectNodeContents(node);
      const rect = range.getBoundingClientRect();
      return { x: rect.left + rect.width * fraction, y: rect.top + rect.height / 2 };
    }
    throw new Error('Editable block has no visible text');
  };
  const start = await first.evaluate(textPoint, 0.2);
  const end = await last.evaluate(textPoint, 0.8);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 24 });
  await page.mouse.up();

  return page.evaluate(({ firstNeedle, lastNeedle }) => {
    const blocks = Array.from(
      document.querySelectorAll<HTMLElement>('#doc [data-anchor][contenteditable="true"]'),
    );
    const describe = (needle: string) => {
      const block = blocks.find((candidate) => candidate.textContent?.includes(needle));
      if (!block) throw new Error(`Could not find editable block containing ${needle}`);
      return { anchor: block.getAttribute('data-anchor')!, text: block.textContent?.trim() ?? '' };
    };
    const selection = window.getSelection()!;
    const blockOf = (node: Node | null) => {
      const element = node?.nodeType === Node.ELEMENT_NODE ? node as Element : node?.parentElement;
      return element?.closest('#doc [data-anchor][contenteditable="true"]')?.getAttribute('data-anchor');
    };
    return {
      first: describe(firstNeedle),
      last: describe(lastNeedle),
      selectionText: selection.rangeCount ? selection.getRangeAt(0).toString() : '',
      anchorBlock: blockOf(selection.anchorNode),
      focusBlock: blockOf(selection.focusNode),
    };
  }, { firstNeedle, lastNeedle });
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
    await expect(page.locator('#doc')).toContainText('Edit this document');

    // The toolbar drives the real editor: select an initially-normal block,
    // click Bold through the actual button handler, and verify the remounted
    // block now carries bold computed styling.
    const result = await selectPracticeBlock(page);
    expect(result.text.length).toBeGreaterThan(0);
    expect(await formattingState(page, result.anchor, result.text)).toMatchObject({
      found: true, bold: false, textPreserved: true,
    });

    await page.locator('#bar [data-fmt="bold"]').click();
    expect(await formattingState(page, result.anchor, result.text)).toMatchObject({
      found: true, bold: true, textPreserved: true,
    });

    await page.locator('#undo').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).bold).toBe(false);
    await page.locator('#redo').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).bold).toBe(true);

    // The compact surface keeps advanced controls in a 4 × 3 overflow palette.
    await reselectBlock(page, result.anchor);
    await page.locator('#more').click();
    await expect(page.locator('#morePanel')).toBeVisible();
    await page.locator('#morePanel [data-fmt="strike"]').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).strike).toBe(true);
    await expect(page.locator('#morePanel')).toBeHidden();

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
    await expect(page.locator('#doc')).toContainText('Edit this document');

    // Main-page controls are not decorative: formatting, page view, and save all
    // call the live DocxEditor instance.
    const result = await dragAcrossDemoBlocks(
      page,
      'SELECT ONLY THIS SENTENCE',
      'This is a real Word document',
    );
    expect(result.anchorBlock).not.toBe(result.focusBlock);
    expect(result.selectionText).toContain('THIS SENTENCE');
    expect(result.selectionText).toContain('This is a real Word document');
    await page.locator('#bar [data-fmt="italic"]').click();
    expect(await formattingState(page, result.first.anchor, result.first.text)).toMatchObject({
      found: true, italic: true, textPreserved: true,
    });
    expect(await formattingState(page, result.last.anchor, result.last.text)).toMatchObject({
      found: true, italic: true, textPreserved: true,
    });

    await page.locator('#fontSize').selectOption('18');
    await expect.poll(async () => (await formattingState(page, result.last.anchor, result.last.text)).fontSize).toBeGreaterThanOrEqual(23);
    await page.locator('#alignment').selectOption('center');
    await expect.poll(async () => (await formattingState(page, result.last.anchor, result.last.text)).alignment).toBe('center');

    const tableCount = await page.locator('#doc table').count();
    await page.locator('#insertTable').click();
    await expect(page.locator('#doc table')).toHaveCount(tableCount + 1);

    await page.locator('#pages').click();
    await expect(page.locator('#pages')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.locator('#doc #pagination-container')).toBeVisible();
    await page.locator('#pages').click();
    await expect(page.locator('#pages')).toHaveAttribute('aria-pressed', 'false');

    const downloadPromise = page.waitForEvent('download');
    await page.locator('#dl').click();
    expect((await downloadPromise).suggestedFilename()).toBe('docxodus-demo.docx');

    // X and LinkedIn get an honest link preview. X no longer documents the
    // historical Player Card, so the live player remains a website iframe.
    const meta = await page.evaluate(() => ({
      card: document.querySelector('meta[name="twitter:card"]')?.getAttribute('content'),
      player: document.querySelector('meta[name="twitter:player"]')?.getAttribute('content'),
      ogImage: document.querySelector('meta[property="og:image"]')?.getAttribute('content'),
    }));
    expect(meta.card).toBe('summary_large_image');
    expect(meta.player).toBeUndefined();
    expect(meta.ogImage).toMatch(/^https:\/\//);

    // The page explains both supported embedding modes and supplies code that
    // is pinned to the release instead of a mutable latest URL.
    await page.locator('#openEmbed').click();
    await expect(page.locator('#embedDialog')).toBeVisible();
    await expect(page.locator('#iframeCode')).toContainText('/demo/player.html');
    await expect(page.locator('#moduleCode')).toContainText(RELEASE_ENGINE);
  });
});
