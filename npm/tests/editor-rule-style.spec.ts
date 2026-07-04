import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 3 — double / styled horizontal rules. The engine + editor API already accept a
// border style on InsertHorizontalRule; only the demo hard-coded "single". This locks the
// capability: a double rule renders with a double bottom border and survives save → reopen.
test.describe('DocxEditor — styled horizontal rules (double)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('insertHorizontalRule(weight, "double") renders a double rule that round-trips', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const countDouble = (root: HTMLElement) =>
        Array.from(root.querySelectorAll('*')).filter(
          (el) => getComputedStyle(el as HTMLElement).borderBottomStyle === 'double',
        ).length;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertHorizontalRule(12, 'double');
      const doubleAfterInsert = countDouble(container);

      // Save → reopen: the double rule persists.
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const doubleAfterReopen = countDouble(c2);

      editor.close();
      e2.close();
      container.remove();
      c2.remove();

      return { doubleAfterInsert, doubleAfterReopen };
    });

    expect(out.doubleAfterInsert).toBeGreaterThanOrEqual(1);
    expect(out.doubleAfterReopen).toBeGreaterThanOrEqual(1);
  });
});
