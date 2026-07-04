import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 1 — line-break fidelity. Shift+Enter must produce a REAL Word line break
// (w:br), not a literal '\n' in w:t (which Word renders as a space). After commit
// the block re-renders from the live session, where a w:br renders back as <br>;
// the round-tripped projection carries the canonical GFM hard break "  \n".
test.describe('DocxEditor — line-break fidelity (Shift+Enter → w:br)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Shift+Enter inserts a faithful line break that survives save/reopen', async ({ page }) => {
    // Set up a blank doc and focus its first paragraph (select the placeholder so the
    // first typed char replaces it). Keep the editor on window for the second step.
    await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'lb-container';
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      (window as any).__lb = { editor, container };
      const target = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      target.focus();
      const r = document.createRange();
      r.selectNodeContents(target);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });

    // REAL keyboard input: triggers the editor's Shift+Enter handler and native typing.
    await page.keyboard.type('AAA');
    await page.keyboard.press('Shift+Enter');
    await page.keyboard.type('BBB');

    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const { editor, container } = (window as any).__lb;
      const firstBlock = () =>
        container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;

      const target = firstBlock();
      const preCommitHTML = target.innerHTML;
      // Commit (blur). The block re-renders from the session.
      target.dispatchEvent(new Event('blur'));

      const committed = firstBlock();
      const brInCommitted = committed.querySelectorAll('br').length;
      const committedText = (committed.textContent || '').replace(/\s+/g, ' ').trim();

      // Save → reopen the bytes → project to markdown (the faithfulness oracle).
      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      // Re-open in a fresh editor: a w:br renders back to <br>.
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const reBlock = c2.querySelector('p[data-anchor]') as HTMLElement;
      const brAfterReopen = reBlock ? reBlock.querySelectorAll('br').length : 0;

      editor.close();
      e2.close();
      container.remove();
      c2.remove();

      return { brInCommitted, committedText, md, brAfterReopen, preCommitHTML };
    });

    // A real line break exists in the committed (re-rendered) block...
    expect(out.brInCommitted).toBeGreaterThanOrEqual(1);
    expect(out.committedText).toContain('AAA');
    expect(out.committedText).toContain('BBB');
    // ...the projected markdown carries the canonical GFM hard break "  \n"...
    expect(out.md).toContain('AAA  \nBBB');
    // ...and it survives a full save → reopen (w:br renders back as <br>).
    expect(out.brAfterReopen).toBeGreaterThanOrEqual(1);
  });
});
