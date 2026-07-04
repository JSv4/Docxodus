import { test, expect, Page } from '@playwright/test';

// Issue 6 — visual table grid picker (demo). Driving the real demo (editor.html): the table
// button opens a hover-to-pick grid; clicking a cell inserts that rows×cols table, replacing
// the old freetext prompt().
test.describe('Demo — visual table grid picker', () => {
  test('picking 3×3 inserts a 3-row, 3-col table', async ({ page }) => {
    await page.goto('/editor.html');
    // Wait for the demo to boot (the test hook appears after WASM init).
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });

    // New blank document, then focus its paragraph so insertTable has an active block.
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      p.focus();
    });

    // Open the picker and choose 3×3 (the cell at row index 2, col index 2).
    await page.click('#table');
    await expect(page.locator('#gridpicker')).toBeVisible();
    await expect(page.locator('#gridcells [data-r="2"][data-c="2"]')).toBeVisible();
    await page.locator('#gridcells [data-r="2"][data-c="2"]').dispatchEvent('mousedown');

    const shape = await page.evaluate(() => {
      const tbl = document.querySelector('#editor table') as HTMLElement | null;
      if (!tbl) return null;
      const firstRow = tbl.querySelector('tr');
      return {
        rows: tbl.querySelectorAll('tr').length,
        cols: firstRow ? firstRow.querySelectorAll('td').length : 0,
      };
    });
    expect(shape).toEqual({ rows: 3, cols: 3 });
    // The picker closed after insertion.
    await expect(page.locator('#gridpicker')).toBeHidden();
  });

  // Gap 2 — the picker's alignment selector controls cell alignment instead of hard-coding
  // center, so an inserted table can be left-aligned (the S-1 filing line no longer comes out
  // centered).
  test('alignment selector controls cell alignment (left → not centered)', async ({ page }) => {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      p.focus();
    });

    await page.click('#table');
    await expect(page.locator('#gridpicker')).toBeVisible();
    await page.selectOption('#gridalign', 'left');
    // Insert a 1×2 table (cell at row 0, col 1).
    await page.locator('#gridcells [data-r="0"][data-c="1"]').dispatchEvent('mousedown');

    const aligns = await page.evaluate(() => {
      const cells = Array.from(
        document.querySelectorAll('#editor table td p[data-anchor]'),
      ) as HTMLElement[];
      return cells.map((c) => getComputedStyle(c).textAlign);
    });
    expect(aligns.length).toBe(2);
    expect(aligns.every((a) => a !== 'center')).toBe(true);
  });
});
