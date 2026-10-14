import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 2 — multi-paragraph table cells. Enter inside a cell must split the cell
// paragraph into TWO paragraphs WITHIN the same cell (Word's behaviour), so a cell
// can hold stacked lines with independent per-paragraph formatting (the S-1
// value-over-label rows and multi-line law-firm address columns). The engine already
// splits correctly inside a w:tc; this verifies the editor surfaces it and that it
// survives a lossless save → reopen.
test.describe('DocxEditor — multi-paragraph table cells (Enter in a cell)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Enter in a cell stacks two paragraphs in the same cell, lossless round-trip', async ({ page }) => {
    // Build a blank doc, insert a 1×2 borderless table, focus the first cell.
    await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'cell-container';
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      (window as any).__cell = { editor, container };

      // Focus the body paragraph (sets the active block) then insert a table after it.
      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertTable(1, 2, { borderless: true });

      // Focus the first cell paragraph and select its placeholder.
      const cellP = container.querySelector('table td p[data-anchor]') as HTMLElement;
      cellP.focus();
      const r = document.createRange();
      r.selectNodeContents(cellP);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });

    // Type the first line, Enter (splits the cell paragraph), type the second line.
    await page.keyboard.type('George J. Sampas');
    await page.keyboard.press('Enter');
    await page.keyboard.type('Hillary H. Holmes');

    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const { editor, container } = (window as any).__cell;

      // Commit whatever cell paragraph is focused.
      const active = document.activeElement as HTMLElement | null;
      if (active && active.getAttribute('data-anchor')) active.dispatchEvent(new Event('blur'));

      const firstCell = container.querySelector('table td') as HTMLElement;
      const cellParas = Array.from(firstCell.querySelectorAll('p')).map(
        (p) => (p.textContent || '').replace(/\s+/g, ' ').trim(),
      );

      // Save → reopen → both lines survive in the SAME cell (two paragraphs).
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const reopenedCell = c2.querySelector('table td') as HTMLElement;
      const reopenedParas = reopenedCell
        ? Array.from(reopenedCell.querySelectorAll('p')).map((p) => (p.textContent || '').replace(/\s+/g, ' ').trim())
        : [];

      editor.close();
      e2.close();
      container.remove();
      c2.remove();

      return { cellParaCount: cellParas.length, cellParas, reopenedParaCount: reopenedParas.length, reopenedParas };
    });

    // Two stacked paragraphs in the first cell, both lines present, after editing...
    expect(out.cellParaCount).toBe(2);
    expect(out.cellParas.join('|')).toContain('George J. Sampas');
    expect(out.cellParas.join('|')).toContain('Hillary H. Holmes');
    // ...and after a lossless save → reopen.
    expect(out.reopenedParaCount).toBe(2);
    expect(out.reopenedParas[0]).toContain('George J. Sampas');
    expect(out.reopenedParas[1]).toContain('Hillary H. Holmes');
  });
});
