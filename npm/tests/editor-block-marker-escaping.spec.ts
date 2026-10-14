import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// #571 — the blur-commit serializer must escape leading block markers.
//
// The projector escapes literal markers on the way out (a paragraph containing
// literal "## x" projects as "\#\# x"), and MarkdownPayloadParser unescapes
// "\X" generically, so the round-trip contract exists. But the editor's
// DOM→markdown serializer escaped only the inline set — a typed leading
// "## ", "> ", "- ", "+ ", or "1. " reached ReplaceText unescaped, was parsed
// as a block construct, and the marker characters were silently dropped from
// the committed document. A list marker after a soft break was worse: the
// splitter promotes it to its own block, and ReplaceText truncates to the
// first block, deleting the whole line.
test.describe('DocxEditor — typed block markers survive commit', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('leading markers typed into a paragraph commit as literal text', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      const firstEditable = () =>
        container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;

      const payloads = [
        '## 2.1 Termination',
        '> see clause 4 for exceptions',
        '- deliverables listed below',
        '+ plus a margin of 2%',
        '12. December invoice',
        '| amount | due |',
      ];
      const committed: string[] = [];
      for (const typed of payloads) {
        const blk = firstEditable();
        blk.focus();
        const sel = window.getSelection()!;
        const r = document.createRange();
        r.selectNodeContents(blk);
        sel.removeAllRanges();
        sel.addRange(r);
        document.execCommand('insertText', false, typed);
        blk.dispatchEvent(new Event('blur'));
        committed.push((firstEditable().textContent || '').trim());
      }

      editor.close();
      container.remove();
      return { payloads, committed };
    });

    for (let i = 0; i < out.payloads.length; i++) {
      expect(out.committed[i], `typed "${out.payloads[i]}" must commit unchanged`)
        .toBe(out.payloads[i]);
    }
  });

  test('a list marker after a soft break does not delete the line', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      const blk = container.querySelector(
        '[data-anchor][contenteditable="true"]',
      ) as HTMLElement;

      // "alpha⏎- beta" with a soft line break, the way Shift+Enter leaves the DOM.
      blk.focus();
      blk.textContent = '';
      blk.appendChild(document.createTextNode('alpha'));
      blk.appendChild(document.createElement('br'));
      blk.appendChild(document.createTextNode('- beta'));
      blk.dispatchEvent(new Event('blur'));

      const text = (container.querySelector(
        '[data-anchor][contenteditable="true"]',
      ) as HTMLElement).textContent || '';
      editor.close();
      container.remove();
      return { text };
    });

    expect(out.text).toContain('alpha');
    expect(out.text, 'the soft-broken second line must survive the commit').toContain('- beta');
  });
});
