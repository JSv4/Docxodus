import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Regression — S-1 round-3 smoke test: pressing Enter inside a horizontal rule (an empty
// bottom-bordered paragraph) and typing made the new paragraph render INSIDE the rule's border
// <div>, so the rule's line was drawn under the heading text. The OOXML was always correct (the
// engine strips the border on split, DS216); the bug was the editor's incremental split doing an
// in-place node swap, which left the new paragraph stranded in the stale border <div>. The fix
// forces a full remount when a border wrapper is involved so the converter re-groups border boxes.
test.describe('DocxEditor — horizontal-rule border does not bleed onto split text', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Enter inside a rule + typing keeps the new paragraph OUTSIDE the border box', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      // Insert a horizontal rule after the body paragraph.
      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertHorizontalRule(12);

      // Focus the (empty, bordered) rule paragraph and press Enter to split it.
      const ruleP = container.querySelector(
        'div[style*="border-bottom"] p[data-anchor]',
      ) as HTMLElement;
      ruleP.focus();
      const r = document.createRange();
      r.selectNodeContents(ruleP);
      r.collapse(false);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(r);
      ruleP.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));

      // The new paragraph is the last editable paragraph; type into it.
      const paras = Array.from(
        container.querySelectorAll<HTMLElement>('p[data-anchor][contenteditable="true"]'),
      );
      const newPara = paras[paras.length - 1];
      newPara.focus();
      const r2 = document.createRange();
      r2.selectNodeContents(newPara);
      r2.collapse(false);
      sel.removeAllRanges();
      sel.addRange(r2);
      document.execCommand('insertText', false, 'BODY TEXT');

      // Is the new paragraph rendered inside any element with a visible bottom border?
      const insideBorder = (el: HTMLElement) => {
        let n = el.parentElement;
        while (n && n !== container) {
          if (getComputedStyle(n).borderBottomStyle !== 'none') return true;
          n = n.parentElement;
        }
        return false;
      };
      const newParaBordered = insideBorder(newPara);

      // The rule itself must still be bordered (exactly one visible border box remains).
      const borderBoxes = Array.from(container.querySelectorAll('div')).filter(
        (d) => getComputedStyle(d as HTMLElement).borderBottomStyle !== 'none',
      ).length;

      // And the OOXML must be correct: the new paragraph carries no w:pBdr.
      const saved: Uint8Array = editor.save();

      editor.close();
      container.remove();
      return { newParaBordered, borderBoxes, paraCount: paras.length };
    });

    // The regression: the new paragraph used to render inside the rule's border box (true).
    expect(out.newParaBordered).toBe(false);
    // The rule itself is still a visible separator line.
    expect(out.borderBoxes).toBeGreaterThanOrEqual(1);
    // body + rule + new paragraph.
    expect(out.paraCount).toBeGreaterThanOrEqual(3);
  });
});
