import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Gap 3 — insert a horizontal rule ABOVE the active block. The S-1 heavy top bar sits between
// the filing table and "UNITED STATES", and rules previously only inserted below, so there was
// no reachable path. insertHorizontalRule(..., 'above') must place the bordered empty paragraph
// BEFORE the active block; the default ('below') keeps it after.
test.describe('DocxEditor — insertHorizontalRule position', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test("'above' places the rule before the active block; default keeps it after", async ({ page }) => {
    const above = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      let blk = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      blk.focus();
      const sel = window.getSelection()!;
      const r = document.createRange(); r.selectNodeContents(blk);
      sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('insertText', false, 'UNITED STATES');
      blk.dispatchEvent(new Event('blur'));
      blk = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      blk.focus();

      editor.insertHorizontalRule(12, 'single', 'above');
      const eds = Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[];
      const markerIdx = eds.findIndex((b) => (b.textContent || '').includes('UNITED STATES'));
      const prevEmpty =
        markerIdx > 0 && (eds[markerIdx - 1].textContent || '').trim() === '';
      editor.close(); container.remove();
      return { markerIdx, prevEmpty, total: eds.length };
    });

    const below = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      let blk = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      blk.focus();
      const sel = window.getSelection()!;
      const r = document.createRange(); r.selectNodeContents(blk);
      sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('insertText', false, 'UNITED STATES');
      blk.dispatchEvent(new Event('blur'));
      blk = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      blk.focus();

      editor.insertHorizontalRule(12, 'single'); // default 'below'
      const eds = Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[];
      const markerIdx = eds.findIndex((b) => (b.textContent || '').includes('UNITED STATES'));
      const nextEmpty =
        markerIdx < eds.length - 1 && (eds[markerIdx + 1].textContent || '').trim() === '';
      editor.close(); container.remove();
      return { markerIdx, nextEmpty, total: eds.length };
    });

    // Above: an empty (rule) paragraph now precedes the marker.
    expect(above.markerIdx).toBe(1);
    expect(above.prevEmpty).toBe(true);
    // Below: the marker stays first and the rule follows it.
    expect(below.markerIdx).toBe(0);
    expect(below.nextEmpty).toBe(true);
  });
});
