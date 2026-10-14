import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 4 — table column widths. A 2-column table can be sized wide-left / narrow-right
// (the S-1 filing-header row "As filed…" | "Registration No. 333-") via columnWidths (twips),
// instead of equal halves that wrap the long left line.
test.describe('DocxEditor — table column widths', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('columnWidths produce an unequal layout that survives save → reopen', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const cellWidths = (root: HTMLElement): number[] =>
        Array.from(root.querySelectorAll('table tr:first-child td')).map(
          (td) => parseFloat(getComputedStyle(td as HTMLElement).width),
        );

      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      // ~2.7 : 1 split.
      editor.insertTable(1, 2, { borderless: true, columnWidths: [7000, 2576] });
      const widths = cellWidths(container);

      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const widthsReopen = cellWidths(c2);

      editor.close();
      e2.close();
      container.remove();
      c2.remove();
      return { widths, widthsReopen };
    });

    // First column clearly wider than the second (not equal halves), both after insert...
    expect(out.widths.length).toBe(2);
    expect(out.widths[0]).toBeGreaterThan(out.widths[1] * 2);
    // ...and after a lossless save → reopen.
    expect(out.widthsReopen.length).toBe(2);
    expect(out.widthsReopen[0]).toBeGreaterThan(out.widthsReopen[1] * 2);
  });
});
