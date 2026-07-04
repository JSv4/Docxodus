import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Gap 4 — deleteBlock removes the active block (e.g. a stray empty paragraph left by a table).
// Guards: inert when the only editable block (don't empty the document) and inert inside a table
// (cells are removed via the table toolbar's delete row/column). Closes the "no block-delete
// affordance" S-1 smoke-test gap.
test.describe('DocxEditor — deleteBlock', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('removes a stray paragraph; inert on the sole block and inside a table', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      const editCount = () =>
        container.querySelectorAll('p[data-anchor][contenteditable="true"]').length;
      const tableCount = () => container.querySelectorAll('table').length;
      const firstEditable = () =>
        container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;

      // (a) Sole-block guard: deleting the only block is a no-op.
      let blk = firstEditable();
      blk.focus();
      editor.deleteBlock();
      const afterSole = editCount();

      // Seed "KEEP" then add an empty rule paragraph below it → 2 editable blocks.
      blk = firstEditable();
      blk.focus();
      const sel = window.getSelection()!;
      const r = document.createRange(); r.selectNodeContents(blk);
      sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('insertText', false, 'KEEP');
      blk.dispatchEvent(new Event('blur'));
      blk = firstEditable();
      blk.focus();
      editor.insertHorizontalRule(12, 'single'); // empty rule paragraph below KEEP
      const afterRule = editCount();

      // (b) Delete the stray empty rule paragraph (focus the empty one).
      const eds = Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[];
      const stray = eds.find((b) => (b.textContent || '').trim() === '')!;
      stray.focus();
      editor.deleteBlock();
      const afterDelete = editCount();
      const remainingText = (firstEditable().textContent || '').trim();

      // (c) Inside-a-table guard: insert a table, focus a cell, deleteBlock is inert.
      firstEditable().focus();
      editor.insertTable(1, 1, {});
      const cell = container.querySelector('table p[data-anchor][contenteditable="true"]') as HTMLElement;
      cell.focus();
      const tablesBefore = tableCount();
      editor.deleteBlock();
      const tablesAfter = tableCount();

      editor.close(); container.remove();
      return { afterSole, afterRule, afterDelete, remainingText, tablesBefore, tablesAfter };
    });

    expect(out.afterSole).toBe(1); // sole-block delete is inert
    expect(out.afterRule).toBe(2); // KEEP + empty rule paragraph
    expect(out.afterDelete).toBe(1); // stray paragraph removed
    expect(out.remainingText).toBe('KEEP'); // the right block survived
    expect(out.tablesBefore).toBe(1);
    expect(out.tablesAfter).toBe(1); // delete inside a table is inert
  });
});
