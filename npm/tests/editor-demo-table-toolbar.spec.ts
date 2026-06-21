import { test, expect, Page } from '@playwright/test';

// Round-5 UX fix — the floating table toolbar must not overlap (and intercept clicks on) the
// document content directly ABOVE a table (e.g. the S-1 "(Exact name…)" line). When an editable
// block sits immediately above the table, the toolbar drops below the table instead.
test.describe('Demo — table toolbar placement', () => {
  test('toolbar does not cover the editable line directly above the table', async ({ page }) => {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());

    // Put a non-empty body line first, then insert a table AFTER it (so a body line is above the table).
    await page.evaluate(() => {
      const editor = (window as any).__demo.getEditor();
      const p = document.querySelector('#editor p[data-anchor][data-editable="1"]') as HTMLElement;
      p.focus();
      const r = document.createRange(); r.selectNodeContents(p);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
      document.execCommand('insertText', false, 'Heading above the table');
      p.dispatchEvent(new Event('blur'));
      // Re-focus the (now non-empty) body line so the table inserts after it.
      const p2 = document.querySelector('#editor p[data-anchor][data-editable="1"]') as HTMLElement;
      p2.focus();
      editor.insertTable(1, 2, { borderless: true });
    });

    // Focus a cell with a real caret so the demo's selectionchange shows + positions the toolbar.
    await page.evaluate(() => {
      const cell = document.querySelector('#editor table td p[data-anchor]') as HTMLElement;
      cell.focus();
      const r = document.createRange(); r.selectNodeContents(cell); r.collapse(true);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await expect(page.locator('#tabletools')).toBeVisible();

    const probe = await page.evaluate(() => {
      const tools = document.getElementById('tabletools')!;
      const table = document.querySelector('#editor table') as HTMLElement;
      const line = Array.from(document.querySelectorAll('#editor p[data-editable="1"]'))
        .find((p) => (p.textContent || '').includes('Heading above the table')) as HTMLElement;
      const lr = line.getBoundingClientRect();
      const tr = tools.getBoundingClientRect();
      const tabr = table.getBoundingClientRect();
      // What element is at the center of the body line above the table? Must NOT be the toolbar.
      const hit = document.elementFromPoint(lr.left + lr.width / 2, lr.top + lr.height / 2) as HTMLElement;
      const overlapsLine = !(tr.bottom <= lr.top || tr.top >= lr.bottom);
      return {
        hitInToolbar: !!hit?.closest('#tabletools'),
        overlapsLine,
        toolbarBelowTable: tr.top >= tabr.top, // dropped below when content sits above
      };
    });

    expect(probe.hitInToolbar).toBe(false); // the line above is clickable, not blocked by the toolbar
    expect(probe.overlapsLine).toBe(false); // toolbar does not vertically overlap the line above
    expect(probe.toolbarBelowTable).toBe(true);
  });
});
