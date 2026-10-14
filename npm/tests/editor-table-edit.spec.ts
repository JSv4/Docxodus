import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 5 — table row/column CRUD. Add/remove rows and columns after insert, addressed by the
// focused cell, surviving a lossless save → reopen.
test.describe('DocxEditor — table row/column editing', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('insert/delete rows and columns reshape the table and round-trip', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertTable(2, 2, { borderless: true });

      const shape = (root: HTMLElement) => {
        const tbl = root.querySelector('table') as HTMLElement | null;
        if (!tbl) return { rows: 0, cols: 0 };
        const firstRow = tbl.querySelector('tr');
        return {
          rows: tbl.querySelectorAll('tr').length,
          cols: firstRow ? firstRow.querySelectorAll('td').length : 0,
        };
      };
      const focusFirstCell = () => {
        const c = container.querySelector('table td p[data-anchor]') as HTMLElement;
        c.focus();
      };

      focusFirstCell();
      editor.insertTableRow('below');
      const afterRow = shape(container);

      focusFirstCell();
      editor.insertTableColumn('right');
      const afterCol = shape(container);

      focusFirstCell();
      editor.deleteTableRow();
      const afterDelRow = shape(container);

      // Save → reopen.
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const reopen = shape(c2);

      editor.close();
      e2.close();
      container.remove();
      c2.remove();
      return { afterRow, afterCol, afterDelRow, reopen };
    });

    expect(out.afterRow).toEqual({ rows: 3, cols: 2 });
    expect(out.afterCol).toEqual({ rows: 3, cols: 3 });
    expect(out.afterDelRow).toEqual({ rows: 2, cols: 3 });
    expect(out.reopen).toEqual({ rows: 2, cols: 3 });
  });
});
