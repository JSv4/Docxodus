import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Hanging / first-line indent via DocxEditor.setIndent → DocxSession.SetParagraphFormat
// (LeftIndent + signed FirstLineIndent). The converter maps w:hanging → negative text-indent and
// w:ind/@w:left → margin-left, so a hanging indent shows margin-left > 0 and text-indent < 0.
test.describe('DocxEditor — hanging / first-line indent', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('setIndent({left, firstLine<0}) applies a hanging indent that survives save→reopen', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const readIndent = (root: HTMLElement) => {
        const p = root.querySelector('p[data-anchor]') as HTMLElement;
        const cs = getComputedStyle(p);
        return {
          left: (parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.paddingLeft) || 0),
          textIndent: parseFloat(cs.textIndent) || 0,
        };
      };

      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.focus();
      const r = document.createRange(); r.selectNodeContents(body); r.collapse(false);
      const s = getSelection()!; s.removeAllRanges(); s.addRange(r);
      document.dispatchEvent(new Event('selectionchange'));

      editor.setIndent({ left: 720, firstLine: -360 });
      const afterInsert = readIndent(container);

      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div'); document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const afterReopen = readIndent(c2);

      editor.close(); e2.close(); container.remove(); c2.remove();
      return { afterInsert, afterReopen };
    });

    expect(out.afterInsert.left).toBeGreaterThan(0);
    expect(out.afterInsert.textIndent).toBeLessThan(0); // hanging → negative text-indent
    expect(out.afterReopen.left).toBeGreaterThan(0);
    expect(out.afterReopen.textIndent).toBeLessThan(0);
  });

  test('setIndent({firstLine>0}) applies a positive first-line indent', async ({ page }) => {
    const textIndent = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.focus();
      const r = document.createRange(); r.selectNodeContents(body); r.collapse(false);
      const s = getSelection()!; s.removeAllRanges(); s.addRange(r);
      document.dispatchEvent(new Event('selectionchange'));

      editor.setIndent({ firstLine: 720 });
      const p = container.querySelector('p[data-anchor]') as HTMLElement;
      const ti = parseFloat(getComputedStyle(p).textIndent) || 0;
      editor.close(); container.remove();
      return ti;
    });
    expect(textIndent).toBeGreaterThan(0); // first-line → positive text-indent
  });
});
