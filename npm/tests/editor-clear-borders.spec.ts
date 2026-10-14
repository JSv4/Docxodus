import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// S-1 smoke-test finding 1b: once a paragraph has an HR border there was no way to clear it
// through the editor surface (the ribbon only ADDS rules; applyParagraphFormat didn't accept
// clearBorders). DocxEditor.clearParagraphBorders() now removes them on the active block.
test.describe('DocxEditor — clear paragraph borders', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('clearParagraphBorders() removes an inserted rule border', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const countBordered = (root: HTMLElement) =>
        Array.from(root.querySelectorAll('*')).filter((el) => {
          const s = getComputedStyle(el as HTMLElement);
          return s.borderBottomStyle !== 'none' && s.borderBottomWidth !== '0px';
        }).length;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertHorizontalRule(24); // thick bottom border on a new empty paragraph after `body`
      const borderedBefore = countBordered(container);

      // Focus the rule paragraph (the last addressable paragraph) and clear its borders.
      const ps = container.querySelectorAll('p[data-anchor][contenteditable="true"]');
      const rule = ps[ps.length - 1] as HTMLElement;
      rule.focus();
      editor.clearParagraphBorders();
      const borderedAfter = countBordered(container);

      editor.close();
      container.remove();
      return { borderedBefore, borderedAfter };
    });

    expect(out.borderedBefore).toBeGreaterThanOrEqual(1); // the rule was there
    expect(out.borderedAfter).toBe(0);                    // and the editor removed it
  });
});
