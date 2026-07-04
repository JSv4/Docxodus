import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Gap 1 — run font family (w:rFonts). The editor's setFontFamily() applies a font to the
// selection (or whole paragraph) through DocxSession.ApplyFormat; it must render and survive
// save → reopen. Closes the "no font-family control" S-1 smoke-test omission (a real S-1 is a
// serif filing; the blank doc seeds Calibri).
test.describe('DocxEditor — setFontFamily', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('applies a font family to a run and round-trips through save/reopen', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const usesFont = (root: HTMLElement, name: string) =>
        Array.from(root.querySelectorAll('span')).some((el) =>
          getComputedStyle(el as HTMLElement).fontFamily.includes(name),
        );

      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      // Seed text into the body paragraph and commit it.
      let blk = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      blk.focus();
      const sel = window.getSelection()!;
      let r = document.createRange();
      r.selectNodeContents(blk);
      sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('insertText', false, 'Registrant');
      blk.dispatchEvent(new Event('blur'));

      // Select the run and set the font family.
      blk = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      blk.focus();
      r = document.createRange();
      r.selectNodeContents(blk);
      sel.removeAllRanges(); sel.addRange(r);
      editor.setFontFamily('Times New Roman');
      const appliedAfter = usesFont(container, 'Times New Roman');

      // Save → reopen: the font persists.
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const appliedAfterReopen = usesFont(c2, 'Times New Roman');

      editor.close(); e2.close(); container.remove(); c2.remove();
      return { appliedAfter, appliedAfterReopen };
    });

    expect(out.appliedAfter).toBe(true);
    expect(out.appliedAfterReopen).toBe(true);
  });
});
