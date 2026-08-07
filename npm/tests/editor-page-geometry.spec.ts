import { test, expect } from '@playwright/test';

/**
 * Page geometry — the editor's continuous view lays the document out at the width its
 * `w:sectPr` defines, and zooms that page to fit a narrow window, instead of reflowing the
 * text column to the device.
 *
 * The bug this pins: on a phone the column collapsed to the viewport (~354 px), so the demo
 * document's 496.8 pt cover table could not fit — a table box never shrinks below its content's
 * minimum — and enlarging a heading from 36 pt to 66 pt pushed that minimum wider still. The
 * result was a heading that ran off the paper and was clipped by the window.
 */

const PHONE = { width: 390, height: 844 };
const DESKTOP = { width: 1280, height: 900 };

/** Every laid-out box that sticks out past the sheet it belongs to. */
async function overflowingBoxes(page: import('@playwright/test').Page) {
  return page.evaluate(() => {
    const sheet = document.querySelector('[data-dxr-surface] .docx-body-flow') as HTMLElement;
    const limit = sheet.getBoundingClientRect().right;
    return Array.from(sheet.querySelectorAll<HTMLElement>('*'))
      .filter((el) => el.getBoundingClientRect().right > limit + 1)
      .map((el) => ({ tag: el.tagName, text: (el.textContent || '').slice(0, 30) }));
  });
}

async function geometry(page: import('@playwright/test').Page) {
  return page.evaluate(() => {
    const surface = document.querySelector('[data-dxr-surface]') as HTMLElement;
    const sheet = surface.querySelector('.docx-body-flow') as HTMLElement;
    const section = sheet.querySelector<HTMLElement>('[data-section-index]');
    return {
      contentWidthPt: section ? parseFloat(section.dataset.contentWidth || '0') : 0,
      pageWidthPt: section ? parseFloat(section.dataset.pageWidth || '0') : 0,
      // Text-column width in points, independent of zoom: the inline width the viewport stamped.
      columnWidthPt: section ? parseFloat(section.style.width) : 0,
      zoom: (window as any).__ribbon.editor.zoom as number,
      surfaceRight: surface.getBoundingClientRect().right,
      sheetRight: sheet.getBoundingClientRect().right,
    };
  });
}

async function openDemo(page: import('@playwright/test').Page) {
  await page.goto('/demo-app.html?engine=./embed.bundle.js');
  await page.waitForFunction(() => !!(window as any).__ribbon?.editor, { timeout: 90000 });
  await page.waitForTimeout(400);
}

/** Select the document's cover heading and set it to `pt` via the ribbon's size combobox. */
async function setHeroFontSize(page: import('@playwright/test').Page, pt: number) {
  const anchor = await page.evaluate(() => {
    const blocks = Array.from(
      document.querySelectorAll<HTMLElement>('[data-dxr-surface] [data-anchor]'),
    );
    const hero = blocks.find((b) => /^Edit this document/.test((b.textContent || '').trim()));
    return hero?.getAttribute('data-anchor') ?? null;
  });
  expect(anchor, 'demo document should have the "Edit this document" cover heading').not.toBeNull();
  await page.evaluate((a) => {
    const p = document.querySelector(`[data-anchor="${a}"]`) as HTMLElement;
    p.focus();
    const r = document.createRange();
    r.selectNodeContents(p);
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
  }, anchor);
  await page.waitForTimeout(150);
  await page.fill('#fontsize', String(pt));
  await page.locator('#fontsize').dispatchEvent('change');
  await page.waitForTimeout(1200);
}

test.describe('Editor — page geometry and fit-to-width', () => {
  test.use({ viewport: PHONE });

  test('a phone lays the document out at the section width and zooms it to fit', async ({ page }) => {
    await openDemo(page);
    const g = await geometry(page);

    // The column is the document's, not the device's: US Letter minus its margins.
    expect(g.contentWidthPt).toBeGreaterThan(400);
    expect(g.columnWidthPt).toBeCloseTo(g.contentWidthPt, 1);
    // A full page cannot fit 390 px unscaled, so the view is zoomed out rather than reflowed.
    expect(g.zoom).toBeLessThan(1);
    // …and the zoomed sheet fits inside the surface.
    expect(g.sheetRight).toBeLessThanOrEqual(g.surfaceRight + 1);
  });

  test('enlarging a heading to 66pt does not push content off the sheet', async ({ page }) => {
    await openDemo(page);
    expect(await overflowingBoxes(page)).toEqual([]);

    await setHeroFontSize(page, 66);

    const heroPx = await page.evaluate(() => {
      const blocks = Array.from(
        document.querySelectorAll<HTMLElement>('[data-dxr-surface] [data-anchor]'),
      );
      const hero = blocks.find((b) => /^Edit this document/.test((b.textContent || '').trim()))!;
      return parseFloat(getComputedStyle(hero.querySelector('span') || hero).fontSize);
    });
    expect(heroPx).toBeGreaterThan(80); // 66pt ≈ 88px — the size really was applied

    expect(await overflowingBoxes(page)).toEqual([]);
  });
});

test.describe('Editor — page geometry on a wide window', () => {
  test.use({ viewport: DESKTOP });

  test('a window wider than the page shows it unscaled', async ({ page }) => {
    await openDemo(page);
    const g = await geometry(page);

    expect(g.zoom).toBe(1);
    expect(g.columnWidthPt).toBeCloseTo(g.contentWidthPt, 1);
    // The sheet is one page wide, not full-bleed.
    expect(g.sheetRight).toBeLessThan(g.surfaceRight + 1);
  });

  test('enlarging a heading to 66pt keeps the text inside the column', async ({ page }) => {
    await openDemo(page);
    await setHeroFontSize(page, 66);
    expect(await overflowingBoxes(page)).toEqual([]);
  });
});
