import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 7 — multi-block formatting. A selection spanning multiple paragraphs applies
// paragraph ops (alignment) and inline ops (bold) to EVERY block in range, not just the
// focused one (the S-1 has ~20 centered lines — one action should center a whole stack).
test.describe('DocxEditor — multi-block formatting', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('selecting three paragraphs centers and bolds all of them', async ({ page }) => {
    await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'mbf';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__mbf = { editor, container };
      const first = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      first.focus();
      const r = document.createRange();
      r.selectNodeContents(first);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });

    // Three paragraphs via real Enter splits.
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.keyboard.press('Enter');
    await page.keyboard.type('CCC');

    // Commit the just-typed last line (blur) so all three blocks are clean, like a user who
    // finished typing before selecting.
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());

    // Drag-style selection from the first line's text to the last line's text.
    const selectAllThree = () =>
      page.evaluate(() => {
        const { container } = (window as any).__mbf;
        const blocks = Array.from(
          container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
        ) as HTMLElement[];
        const firstText = (el: HTMLElement) => document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
        const lastText = (el: HTMLElement) => {
          const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          let t: Node | null, last: Node | null = null;
          while ((t = w.nextNode())) last = t;
          return last;
        };
        const fn = firstText(blocks[0]) || blocks[0];
        const ln = lastText(blocks[blocks.length - 1]) || blocks[blocks.length - 1];
        const r = document.createRange();
        r.setStart(fn, 0);
        r.setEnd(ln, ln.nodeType === 3 ? (ln.textContent || '').length : ln.childNodes.length);
        const s = window.getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
      });

    // Center all three.
    await selectAllThree();
    await page.evaluate(() => (window as any).__mbf.editor.setAlignment('center'));
    // Bold all three (re-select — the center op re-rendered the blocks).
    await selectAllThree();
    await page.evaluate(() => (window as any).__mbf.editor.format('bold'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__mbf;
      const read = (root: HTMLElement) =>
        (Array.from(root.querySelectorAll('p[data-anchor]')) as HTMLElement[])
          .filter((p) => /[ABC]{3}/.test(p.textContent || ''))
          .map((p) => {
            const span = p.querySelector('span');
            return {
              text: (p.textContent || '').trim(),
              align: getComputedStyle(p).textAlign,
              bold: span ? parseInt(getComputedStyle(span).fontWeight, 10) >= 600 : false,
            };
          });
      const live = read(container);

      // Round-trip.
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = (window as any).Docxodus.DocxEditor.open(c2, saved, (window as any).Docxodus, {});
      const reopened = read(c2);
      editor.close(); e2.close(); container.remove(); c2.remove();
      return { live, reopened };
    });

    expect(out.live.length).toBe(3);
    expect(out.live.every((b: any) => b.align === 'center')).toBe(true);
    expect(out.live.every((b: any) => b.bold)).toBe(true);
    // Survives save → reopen.
    expect(out.reopened.length).toBe(3);
    expect(out.reopened.every((b: any) => b.align === 'center')).toBe(true);
    expect(out.reopened.every((b: any) => b.bold)).toBe(true);
  });
});
