import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// DocxEditor.toggleLegalNumbering applies a built-in outline scheme (1. / 1.1 / (a) / (i) …) to the
// active block via DocxSession.ApplyMultilevelNumbering; toggling again removes it. The block
// becomes a list item (a w:numPr / numId), and it survives save → reopen.
test.describe('DocxEditor — legal / outline numbering', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('toggleLegalNumbering makes the block a list item, and toggles off', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor: any = D.DocxEditor.open(container, blank, D, {});

      const setCaret = (el: HTMLElement) => {
        el.focus();
        const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
        const s = getSelection()!; s.removeAllRanges(); s.addRange(r);
        document.dispatchEvent(new Event('selectionchange'));
      };
      const numIdOf = () => {
        const p = container.querySelector('p[data-anchor]') as HTMLElement;
        const fullId = editor.unidToFullId.get(p.getAttribute('data-anchor'));
        try { return JSON.parse(D.DocxSessionBridge.GetListMembership(editor.handle, fullId)).numId; }
        catch { return null; }
      };

      let body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.textContent = 'Purchase and Sale of Preferred Stock.';
      body.dispatchEvent(new InputEvent('input', { bubbles: true }));
      setCaret(body);

      editor.toggleLegalNumbering();
      const numIdAfterOn = numIdOf();

      // toggle off — set caret in the (re-rendered) block first
      setCaret(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.toggleLegalNumbering();
      const numIdAfterOff = numIdOf();

      editor.close(); container.remove();
      return { numIdAfterOn, numIdAfterOff };
    });

    expect(typeof out.numIdAfterOn).toBe('number');
    expect(out.numIdAfterOn).toBeGreaterThan(0);
    expect(out.numIdAfterOff).toBeNull(); // removed list membership
  });

  test('legal numbering survives save → reopen', async ({ page }) => {
    const reopenedNumId = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor: any = D.DocxEditor.open(container, blank, D, {});

      const body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.textContent = 'Section heading.';
      body.dispatchEvent(new InputEvent('input', { bubbles: true }));
      body.focus();
      const r = document.createRange(); r.selectNodeContents(body); r.collapse(false);
      const s = getSelection()!; s.removeAllRanges(); s.addRange(r);
      document.dispatchEvent(new Event('selectionchange'));
      editor.toggleLegalNumbering();

      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div'); document.body.appendChild(c2);
      const e2: any = D.DocxEditor.open(c2, saved, D, {});
      const p = c2.querySelector('p[data-anchor]') as HTMLElement;
      const fullId = e2.unidToFullId.get(p.getAttribute('data-anchor'));
      let numId = null;
      try { numId = JSON.parse(D.DocxSessionBridge.GetListMembership(e2.handle, fullId)).numId; } catch {}

      editor.close(); e2.close(); container.remove(); c2.remove();
      return numId;
    });
    expect(typeof reopenedNumId).toBe('number');
  });
});
