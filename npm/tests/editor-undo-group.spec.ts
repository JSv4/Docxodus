import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// A compound edit (multi-block format; later, cross-block delete) must be ONE undo unit.
test.describe('DocxEditor — atomic undo grouping', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // Build AAA / BBB / CCC in a blank doc, select all three, and return the editor + a reader.
  async function buildThree(page: Page, id: string) {
    await page.evaluate((cid) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div'); container.id = cid;
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any)[cid] = { editor, container };
      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      const r = document.createRange(); r.selectNodeContents(first);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    }, id);
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.keyboard.press('Enter');
    await page.keyboard.type('CCC');
    await page.evaluate((cid) => (window as any)[cid].editor.commitAllDirty(), id);
  }

  async function selectAllThree(page: Page, id: string) {
    await page.evaluate((cid) => {
      const { container } = (window as any)[cid];
      const blocks = Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
        .filter((p: any) => /[ABC]{3}/.test(p.textContent || '')) as HTMLElement[];
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
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    }, id);
  }

  test('a multi-block format is reversed by a single undo', async ({ page }) => {
    await buildThree(page, 'ug1');
    await selectAllThree(page, 'ug1');
    await page.evaluate(() => (window as any).ug1.editor.format('bold'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).ug1;
      const read = () =>
        (Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
          .filter((p) => /[ABC]{3}/.test(p.textContent || ''))
          .map((p) => { const sp = p.querySelector('span'); return sp ? parseInt(getComputedStyle(sp).fontWeight, 10) >= 600 : false; });
      const afterFormat = read();
      editor.undo(); // ONE undo
      const afterUndo = read();
      editor.close(); container.remove();
      return { afterFormat, afterUndo };
    });

    expect(out.afterFormat.length).toBe(3);
    expect(out.afterFormat.every((b: boolean) => b)).toBe(true);
    expect(out.afterUndo.length).toBe(3);                      // blocks intact
    expect(out.afterUndo.some((b: boolean) => b)).toBe(false); // one undo cleared ALL bold
  });
});
