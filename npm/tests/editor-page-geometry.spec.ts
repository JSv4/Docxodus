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

  for (const bands of [false, true]) {
    test(`zoom scales the paper with its contents (header/footer bands ${bands ? 'on' : 'off'})`, async ({ page }) => {
      await openDemo(page);
      if (bands) {
        await page.locator('.dxr-tab[data-tab="layout"]').click();
        await page.locator('[data-dxr="headerfooter"]').check();
      }
      await page.locator('.dxr-tab[data-tab="view"]').click();

      const measure = () => page.evaluate(() => {
        const surface = document.querySelector<HTMLElement>('[data-dxr-surface]')!;
        const sheet = surface.querySelector<HTMLElement>('.docx-body-flow')!;
        const section = sheet.querySelector<HTMLElement>('[data-section-index]')!;
        const table = sheet.querySelector<HTMLElement>('table')!;
        const scroll = surface.closest<HTMLElement>('.dxr-scroll')!;
        const rect = (el: HTMLElement) => {
          const { left, right, width } = el.getBoundingClientRect();
          return { left, right, width };
        };
        return {
          pageWidth: Number(section.dataset.pageWidth) * 96 / 72,
          sheet: rect(sheet),
          table: rect(table),
          bands: Array.from(surface.querySelectorAll<HTMLElement>('.docx-hf-band'), rect),
          scrollWidth: scroll.scrollWidth,
          viewportWidth: scroll.clientWidth,
          chromeWidth: surface.closest('.dxr')!.getBoundingClientRect().width,
        };
      });

      const original = await measure();
      // Require real page geometry and a substantial cover table: an empty document or
      // a zoom control that does nothing must not satisfy this regression guard.
      expect(original.pageWidth).toBeGreaterThan(700);
      expect(original.table.width).toBeGreaterThan(500);
      expect(original.bands).toHaveLength(bands ? 2 : 0);

      for (const [control, scale] of [['zoom', 2], ['zoomlevel', 0.5], ['zoom', 1]] as const) {
        await page.locator(`[data-dxr="${control}"]`).selectOption(String(scale));
        await expect(page.locator('[data-dxr="zoom"]')).toHaveValue(String(scale));
        await expect(page.locator('[data-dxr="zoomlevel"]')).toHaveValue(String(scale));
        const current = await measure();
        expect(Math.abs(current.sheet.width - original.pageWidth * scale)).toBeLessThan(1);
        expect(Math.abs(current.table.width - original.table.width * scale)).toBeLessThan(1);
        expect(current.table.left).toBeGreaterThanOrEqual(current.sheet.left);
        expect(current.table.right).toBeLessThanOrEqual(current.sheet.right + 1);
        expect(current.chromeWidth).toBe(original.chromeWidth);
        expect(await overflowingBoxes(page)).toEqual([]);
        for (const band of current.bands) {
          expect(Math.abs(band.left - current.sheet.left)).toBeLessThan(1);
          expect(Math.abs(band.right - current.sheet.right)).toBeLessThan(1);
        }
        if (scale === 2) {
          expect(current.sheet.width).toBeGreaterThan(current.viewportWidth);
          expect(current.scrollWidth).toBeGreaterThan(current.viewportWidth);
          // The enlarged paper must be reachable by scrolling within the editor.
          const reached = await page.evaluate(() => {
            const scroll = document.querySelector<HTMLElement>('.dxr-scroll')!;
            scroll.scrollLeft = scroll.scrollWidth;
            const sheet = scroll.querySelector('.docx-body-flow')!.getBoundingClientRect();
            const viewport = scroll.getBoundingClientRect();
            const result = { offset: scroll.scrollLeft, right: sheet.right, limit: viewport.right };
            scroll.scrollLeft = 0;
            return result;
          });
          expect(reached.offset).toBeGreaterThan(0);
          expect(reached.right).toBeLessThanOrEqual(reached.limit + 1);
        }
      }

      // Remounting between views must preserve the same zoomed paper geometry.
      await page.locator('[data-dxr="zoom"]').selectOption('2');
      await page.locator('[data-dxr="viewpage"]').click();
      const firstPage = await page.locator('.page-box').first().boundingBox();
      expect(firstPage).not.toBeNull();
      expect(Math.abs(firstPage!.width - original.pageWidth * 2)).toBeLessThan(1);
      await page.locator('[data-dxr="viewweb"]').click();
      expect(Math.abs((await measure()).sheet.width - original.pageWidth * 2)).toBeLessThan(1);
      await page.locator('[data-dxr="zoomfit"]').click();
      expect(Math.abs((await measure()).sheet.width - original.pageWidth)).toBeLessThan(1);
    });
  }
});
