import { test, expect, Page } from '@playwright/test';

// Regression for the complex-filing smoke-test data-loss bug: text typed into a table cell and then
// left (focus moves to a body block) was NOT committed to the session, so the next structural insert
// re-rendered the document from the model and silently erased it. The fix commits a cell on blur and
// flushes all dirty blocks before a (non-undo) remount.

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('DocxEditor — table-cell commit / data loss', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('cell text survives a later structural insert (no silent data loss)', async ({ page }) => {
    const cellTextAfter = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});

      const firstCell = () =>
        container.querySelector('table [data-anchor][data-editable="1"]') as HTMLElement | null;
      const bodyBlock = () =>
        Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
          .find((b) => !(b as HTMLElement).closest('table')) as HTMLElement | undefined;

      // Insert a 1×2 borderless table relative to the focused first block.
      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      editor.insertTable(1, 2, { borderless: true });

      // Type into the first cell, then leave it for a body block WITHOUT an explicit commit.
      const cell = firstCell()!;
      cell.focus();
      cell.textContent = 'TexasTax';
      const body = bodyBlock()!;
      body.focus(); // cell -> body (historically did NOT commit the cell)

      // A structural insert re-renders the whole document from the session.
      editor.insertHorizontalRule(12, 'single', 'below');

      const text = (firstCell()?.textContent || '').trim();
      editor.close();
      return text;
    });
    expect(cellTextAfter).toBe('TexasTax');
  });

  test('undo discards an uncommitted edit — remount flush is gated off for undo', async ({ page }) => {
    const blockTextAfterUndo = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});

      const block = () => container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;

      // Commit one edit (recorded in the undo history).
      const b1 = block();
      b1.focus();
      b1.textContent = 'COMMITTED';
      b1.dispatchEvent(new Event('blur'));

      // Type a NEW edit but do NOT commit it.
      const b2 = block();
      b2.focus();
      b2.textContent = 'DIRTY-UNCOMMITTED';

      // Undo must restore the pre-COMMITTED state, NOT commit the dirty DOM on top of the snapshot.
      editor.undo();

      const text = (block()?.textContent || '').trim();
      editor.close();
      return text;
    });
    expect(blockTextAfterUndo).not.toContain('DIRTY-UNCOMMITTED');
  });
});
