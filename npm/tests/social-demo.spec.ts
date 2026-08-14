import { test, expect, type Page } from '@playwright/test';

/**
 * The GitHub Pages demo trio, all three of which now host the SAME editor
 * surface (`createRibbonEditor`) rather than a hand-rolled toolbar each:
 *
 *  - demo-index.html — the landing page, editor embedded in its dark frame
 *  - demo-app.html   — the full-bleed editor
 *  - player.html     — the compact, boot-on-tap iframe target
 *
 * In production they pull the engine from jsDelivr and the sample document from
 * the same Pages directory; all three accept `?engine=` / `?doc=` overrides so
 * this spec can drive them fully locally. pretest copies them into the test
 * webroot beside embed.bundle.js and the sample docx. The local layout also
 * exercises the embed bundle's wasm-webroot fallback (assets next to the
 * bundle, not under wasm/).
 *
 * Because the surface is shared, the control selectors below are the ribbon's
 * own (`button[data-cmd]`, `#fontsize`, `#paginated`, `#save`) — one contract
 * covering every host.
 */

const OVERRIDES = 'engine=./embed.bundle.js&doc=./docxodus-demo-guide.docx';
const RELEASE_ENGINE = 'docxodus@9.9.0/dist/embed.bundle.js';

async function formattingState(page: Page, anchor: string, text: string) {
  return page.evaluate(({ anchor, text }) => {
    const block = Array.from(document.querySelectorAll<HTMLElement>('#editor [data-anchor]'))
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
      document.querySelectorAll<HTMLElement>('#editor [data-anchor][contenteditable="true"]'),
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
    const block = document.querySelector<HTMLElement>(`#editor [data-anchor="${CSS.escape(anchor)}"]`);
    if (!block) throw new Error(`Block ${anchor} was not remounted`);
    const range = document.createRange();
    range.selectNodeContents(block);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  }, anchor);
}

async function dragAcrossDemoBlocks(page: Page, firstNeedle: string, lastNeedle: string) {
  const editable = page.locator('#editor [data-anchor][contenteditable="true"]');
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
      document.querySelectorAll<HTMLElement>('#editor [data-anchor][contenteditable="true"]'),
    );
    const describe = (needle: string) => {
      const block = blocks.find((candidate) => candidate.textContent?.includes(needle));
      if (!block) throw new Error(`Could not find editable block containing ${needle}`);
      return { anchor: block.getAttribute('data-anchor')!, text: block.textContent?.trim() ?? '' };
    };
    const selection = window.getSelection()!;
    const blockOf = (node: Node | null) => {
      const element = node?.nodeType === Node.ELEMENT_NODE ? node as Element : node?.parentElement;
      return element?.closest('#editor [data-anchor][contenteditable="true"]')?.getAttribute('data-anchor');
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
  test('player.html boots the shared surface on tap, in its compact layout', async ({ page }) => {
    await page.goto(`/player.html?${OVERRIDES}`);
    expect(await page.locator('script[type="module"]').textContent()).toContain(RELEASE_ENGINE);
    // Poster state first — nothing heavy loads until the user taps.
    await expect(page.locator('#start')).toBeVisible();
    await expect(page.locator('#app')).toBeHidden();

    await page.click('#start');
    await expect(page.locator('#app')).toBeVisible({ timeout: 45000 });
    await expect(page.locator('.dxr')).toHaveAttribute('data-state', 'ready', { timeout: 45000 });
    // Pinned compact: this page exists to be the small layout regardless of frame width.
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'compact');
    expect(await page.locator('#editor [data-anchor]').count()).toBeGreaterThan(0);
    await expect(page.locator('#editor')).toContainText('Edit this document');

    // The ribbon drives the real editor: select an initially-normal block, click
    // Bold through the actual control, and verify the remounted block is bold.
    const result = await selectPracticeBlock(page);
    expect(result.text.length).toBeGreaterThan(0);
    expect(await formattingState(page, result.anchor, result.text)).toMatchObject({
      found: true, bold: false, textPreserved: true,
    });

    await page.locator('button[data-cmd="bold"]').click();
    expect(await formattingState(page, result.anchor, result.text)).toMatchObject({
      found: true, bold: true, textPreserved: true,
    });

    await page.locator('#undo').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).bold).toBe(false);
    await page.locator('#redo').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).bold).toBe(true);

    // The compact layout keeps every control reachable — it scrolls the ribbon
    // strip rather than dropping commands, which the old mini toolbar did.
    await reselectBlock(page, result.anchor);
    await page.locator('button[data-cmd="strike"]').click();
    await expect.poll(async () => (await formattingState(page, result.anchor, result.text)).strike).toBe(true);

    const downloadPromise = page.waitForEvent('download');
    await page.locator('#save').click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toBe('docxodus-demo.docx');
  });

  test('index.html landing boots the editor and reports live status', async ({ page }) => {
    await page.goto(`/demo-index.html?${OVERRIDES}`);
    expect(await page.locator('script[type="module"]').textContent()).toContain(RELEASE_ENGINE);
    // A desktop-width frame still gets the plain editor; the arcade swap is a
    // phone default, not a change of what this page is.
    expect(await page.evaluate(() => document.documentElement.dataset.demo)).toBe('editor');
    // The loading overlay belongs to the surface now, and narrates the wait.
    await expect(page.locator('#loader')).toBeVisible();
    await expect(page.locator('#loaderAdTitle')).not.toBeEmpty();
    await expect(page.locator('#pageStatus')).toContainText(/live/i, { timeout: 45000 });
    await expect(page.locator('#loader')).toBeHidden();
    expect(await page.locator('#editor [data-anchor]').count()).toBeGreaterThan(0);
    await expect(page.locator('#editor')).toContainText('Edit this document');
    // A 1220px frame gets the roomy layout, tab strip and all.
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'full');
    await expect(page.locator('.dxr-tab[data-tab="insert"]')).toBeVisible();

    // Ribbon controls are not decorative: formatting, page view, and save all
    // call the live DocxEditor instance.
    const result = await dragAcrossDemoBlocks(
      page,
      'SELECT ONLY THIS SENTENCE',
      'This is a real Word document',
    );
    expect(result.anchorBlock).not.toBe(result.focusBlock);
    expect(result.selectionText).toContain('THIS SENTENCE');
    expect(result.selectionText).toContain('This is a real Word document');
    await page.locator('button[data-cmd="italic"]').click();
    expect(await formattingState(page, result.first.anchor, result.first.text)).toMatchObject({
      found: true, italic: true, textPreserved: true,
    });
    expect(await formattingState(page, result.last.anchor, result.last.text)).toMatchObject({
      found: true, italic: true, textPreserved: true,
    });

    // Font size is a combobox (any positive value), not a fixed preset list.
    await page.locator('#fontsize').fill('18');
    await page.locator('#fontsize').press('Enter');
    await expect.poll(async () => (await formattingState(page, result.last.anchor, result.last.text)).fontSize).toBeGreaterThanOrEqual(23);

    await reselectBlock(page, result.last.anchor);
    await page.locator('button[data-align="center"]').click();
    await expect.poll(async () => (await formattingState(page, result.last.anchor, result.last.text)).alignment).toBe('center');

    // Insert tab → visual grid picker → a real table in the document.
    const tableCount = await page.locator('#editor table').count();
    await page.locator('.dxr-tab[data-tab="insert"]').click();
    await page.locator('#table').click();
    await expect(page.locator('#gridpicker')).toBeVisible();
    await page.locator('#gridcells div').nth(11).click(); // row 2, column 2
    await expect(page.locator('#editor table')).toHaveCount(tableCount + 1);

    // Layout tab → page view, re-rendered from the LIVE session.
    await page.evaluate(() => (window as any).__ribbon.selectTab('layout'));
    await page.locator('#paginated').check();
    await expect(page.locator('#editor #pagination-container')).toBeVisible();
    await page.locator('#paginated').uncheck();
    await expect(page.locator('#editor #pagination-container')).toHaveCount(0);

    const downloadPromise = page.waitForEvent('download');
    await page.locator('#save').click();
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
    await expect(page.locator('#moduleCode')).toContainText('createRibbonEditor');
  });

  test('app.html serves the same surface full-bleed, and adapts to a phone', async ({ page }) => {
    await page.goto(`/demo-app.html?${OVERRIDES}`);
    expect(await page.locator('script[type="module"]').textContent()).toContain(RELEASE_ENGINE);
    await expect(page.locator('.dxr')).toHaveAttribute('data-state', 'ready', { timeout: 45000 });
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'full');
    expect(await page.locator('#editor [data-anchor]').count()).toBeGreaterThan(0);

    // The surface owns the viewport, and scrolls the DOCUMENT rather than the page,
    // so the ribbon stays reachable in a long contract.
    const layout = await page.evaluate(() => {
      const root = document.querySelector<HTMLElement>('.dxr')!;
      const scroll = root.querySelector<HTMLElement>('.dxr-scroll')!;
      return {
        fillsViewport: Math.abs(root.getBoundingClientRect().height - window.innerHeight) < 4,
        documentScrolls: scroll.scrollHeight > scroll.clientHeight,
        pageScrolls: document.documentElement.scrollHeight > window.innerHeight + 2,
      };
    });
    expect(layout.fillsViewport).toBe(true);
    expect(layout.documentScrolls).toBe(true);
    expect(layout.pageScrolls).toBe(false);

    // Same page, phone viewport: the chrome collapses without a reload, because
    // density is measured from the container rather than baked in at mount.
    await page.setViewportSize({ width: 390, height: 844 });
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'compact');
    await expect(page.locator('.dxr-rail')).toBeHidden();
    await expect(page.locator('button[data-cmd="bold"]')).toBeVisible();
    // Touch targets grow with the layout, not with the pointer alone.
    const tap = await page.locator('button[data-cmd="bold"]').evaluate((el) => el.getBoundingClientRect().height);
    expect(tap).toBeGreaterThanOrEqual(30);

    await page.setViewportSize({ width: 1280, height: 900 });
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'full');
    await expect(page.locator('.dxr-rail')).toBeVisible();
  });
});

/**
 * The landing page on a phone. Two changes, one theme — a phone visitor is not
 * going to draft a contract on a 390px screen, so the page stops offering them
 * one first and stops dead-ending them second:
 *
 *  - it mounts THE DOCX ARCADE into the same frame, on the same shipped
 *    surface, with the plain editor one tap away (`?demo=editor`);
 *  - the nav links, which used to collapse to `display: none` below 620px,
 *    stay reachable as a scroll strip — they point at the other demo pages,
 *    which is precisely what a phone visitor wants;
 *  - the arcade's controls float where the thumbs are instead of stacking into
 *    a wrapped bar over the game.
 *
 * A phone-shaped viewport is enough to drive all of it: the demo decision and
 * the controls' density are both taken from measured WIDTH, not from the
 * pointer type, so they are exercisable here rather than only on the Pixel 5
 * rig (which covers the cabinet page and the font/grid properties instead).
 */
test.describe('The landing page on a phone', () => {
  test.use({ viewport: { width: 390, height: 844 } });

  const NAV_LINKS = ['GitHub', 'npm', 'Full editor', 'Compact player', 'Observatory', 'Arcade'];

  test('mounts the arcade by default and keeps every nav destination reachable', async ({ page }) => {
    test.setTimeout(150000);
    await page.goto(`/demo-index.html?${OVERRIDES}&intro=0`);

    // Decided in <head>, so it is already true of the first paint — the copy
    // around the frame can never advertise the demo the page did not mount.
    expect(await page.evaluate(() => document.documentElement.dataset.demo)).toBe('arcade');
    await expect(page.locator('.section-title.for-arcade')).toBeVisible();
    await expect(page.locator('.section-title.for-editor')).toBeHidden();
    await expect(page.locator('.hero-button.for-arcade').first()).toBeVisible();
    // The escape hatch is on the page, not only in the nav.
    await expect(page.locator('.section-sub.for-arcade a[href="?demo=editor"]')).toBeVisible();

    // The regression this replaces: every link but the CTA was display:none.
    // "Reachable" is a laid-out box in a scroll strip, not necessarily one
    // already on screen — so measure the boxes and the scroll affordance.
    const nav = await page.evaluate((labels) => {
      const bar = document.querySelector<HTMLElement>('.topbar nav')!;
      const links = Array.from(bar.querySelectorAll<HTMLAnchorElement>('a'));
      return {
        overflowX: getComputedStyle(bar).overflowX,
        laidOut: labels.every((label) => {
          const link = links.find((a) => a.textContent?.trim() === label);
          return !!link && link.getBoundingClientRect().width > 0;
        }),
        collapsed: links.filter((a) => getComputedStyle(a).display === 'none').length,
        // Whatever does not fit is swipeable rather than lost.
        clipped: bar.scrollWidth > bar.clientWidth,
      };
    }, NAV_LINKS);
    expect(nav.laidOut, 'every nav destination must be laid out on a phone').toBe(true);
    expect(nav.collapsed, 'no nav link may be collapsed away').toBe(0);
    expect(nav.overflowX, 'anything past the edge must stay swipeable').toBe('auto');
    expect(nav.clipped, 'six chips do not fit 390px — the strip must overflow, not wrap away').toBe(true);

    // …and the arcade really boots, on the shipped surface, in the card.
    await page.waitForFunction(
      () => (window as any).__arcade !== undefined || (window as any).__demoError !== undefined,
      { timeout: 90000 },
    );
    expect(await page.evaluate(() => (window as any).__demoError)).toBeUndefined();
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 3, { timeout: 60000 });
    await expect(page.locator('#pageStatus')).toContainText(/live/i);
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'compact');
    expect(await page.evaluate(() => (window as any).__arcade.canvasText() as string))
      .toContain('│'); // the game screen is drawn in the document
  });

  test('the floating controls steer the game and keep the screen clear', async ({ page }) => {
    test.setTimeout(150000);
    await page.goto(`/demo-index.html?${OVERRIDES}&intro=0&cart=quest`);
    await page.waitForFunction(
      () => (window as any).__arcade !== undefined || (window as any).__demoError !== undefined,
      { timeout: 90000 },
    );
    expect(await page.evaluate(() => (window as any).__demoError)).toBeUndefined();
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 3, { timeout: 60000 });

    await expect(page.locator('.dxa-controls')).toHaveAttribute('data-compact', 'true');
    await expect(page.locator('#pad .dxa-dpad')).toBeVisible();
    await expect(page.locator('#pad .dxa-fire')).toBeVisible();

    // The HUD keeps only what you touch mid-frame; the rest is behind "⋯",
    // still one tap away — nothing is dropped, only re-placed.
    await expect(page.locator('#playpause')).toBeVisible();
    await expect(page.locator('#pace')).toBeVisible();
    await expect(page.locator('#dockcarts')).toBeHidden();
    await page.locator('#dockmore').click();
    await expect(page.locator('#dockcarts')).toBeVisible();
    expect(await page.locator('#dockcarts button').count()).toBe(3);
    await expect(page.locator('#restart')).toBeVisible();
    // The pad sits a fixed distance above the dock, so an opened sheet grows
    // the dock up underneath it: the D-pad covered the cartridge buttons and
    // the telemetry, and took the taps meant for them. It stands down instead.
    //
    // Assert on the CLUSTER, not on `#pad`: that wrapper spans the host with
    // both clusters absolutely positioned inside it, so it is 390 × 0 and
    // Playwright calls it hidden whether the pad is up or down.
    await expect(page.locator('.dxa-dpad')).toBeHidden();
    await page.locator('#dockmore').click();
    await expect(page.locator('#dockcarts')).toBeHidden();
    await expect(page.locator('.dxa-dpad')).toBeVisible();

    // Thumb reach, inside the card, and out of the middle of the game screen.
    const geometry = await page.evaluate(() => {
      const box = (selector: string) =>
        document.querySelector(selector)!.getBoundingClientRect();
      const frame = box('#frame');
      const dpad = box('.dxa-dpad');
      const fire = box('.dxa-fire');
      return {
        insideCard: dpad.left >= frame.left && fire.right <= frame.right
          && dpad.bottom <= frame.bottom && fire.bottom <= frame.bottom,
        dpadOnTheLeft: dpad.right < frame.left + frame.width / 2,
        fireOnTheRight: fire.left > frame.left + frame.width / 2,
        // Big enough to hit without looking (Apple/Google both say 44px).
        tapTargets: Math.min(dpad.width / 3, dpad.height / 3, fire.width, fire.height),
        clearOfCentre: dpad.top > frame.top + frame.height / 2,
      };
    });
    expect(geometry.insideCard, 'the controls overlay the card, not the page').toBe(true);
    expect(geometry.dpadOnTheLeft).toBe(true);
    expect(geometry.fireOnTheRight).toBe(true);
    expect(geometry.tapTargets).toBeGreaterThanOrEqual(40);
    expect(geometry.clearOfCentre).toBe(true);

    // Holding ▶ runs the pilcrow right, through the same input the keyboard
    // feeds — the control is wired to the simulation, not to a stub.
    const before = await page.evaluate(() => (window as any).__arcade.game().player.x as number);
    const right = (await page.locator('.dxa-right').boundingBox())!;
    await page.mouse.move(right.x + right.width / 2, right.y + right.height / 2);
    await page.mouse.down();
    await page.waitForFunction(
      (x0) => (window as any).__arcade.game().player.x > x0 + 2, before, { timeout: 30000 });
    await page.mouse.up();

    // FIRE is Space — jump here, the weapon in the raycasters, and the coin
    // drop on the attract screen. A phone had no way to send it at all before.
    const fire = (await page.locator('.dxa-fire').boundingBox())!;
    await page.mouse.move(fire.x + fire.width / 2, fire.y + fire.height / 2);
    await page.mouse.down();
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Space'))).toBe(true);
    await page.mouse.up();
    expect(await page.evaluate(() => (window as any).__arcade.input.held('Space'))).toBe(false);
  });

  test('?demo=editor pins the plain editor on the same phone viewport', async ({ page }) => {
    await page.goto(`/demo-index.html?${OVERRIDES}&demo=editor`);
    expect(await page.evaluate(() => document.documentElement.dataset.demo)).toBe('editor');
    await expect(page.locator('.section-title.for-editor')).toBeVisible();
    await expect(page.locator('#pageStatus')).toContainText(/live/i, { timeout: 90000 });
    expect(await page.locator('#editor [data-anchor]').count()).toBeGreaterThan(0);
    await expect(page.locator('.dxa-controls')).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__arcade)).toBeUndefined();
  });
});
