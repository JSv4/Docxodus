import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// A REAL user gesture (mouse drag across blocks) must produce a multi-block selection that editor
// commands act on. This is the gap: per-block contenteditable made cross-block selection unreachable.
test.describe('DocxEditor — multi-block selection reachability', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('mouse-drag across three paragraphs bolds all of them', async ({ page }) => {
    // Build a blank doc with three paragraphs (AAA / BBB / CCC) via real typing + Enter.
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'mbr';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__mbr = { editor, container };
      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      const r = document.createRange(); r.selectNodeContents(first);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.keyboard.press('Enter');
    await page.keyboard.type('CCC');
    await page.evaluate(() => (window as any).__mbr.editor.commitAllDirty());

    // Real drag: press at the start of block 1, move to the end of block 3, release.
    const boxes = await page.evaluate(() => {
      const { container } = (window as any).__mbr;
      const bs = Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
        .filter((p: any) => /[ABC]{3}/.test(p.textContent || '')) as HTMLElement[];
      const f = bs[0].getBoundingClientRect(); const l = bs[bs.length - 1].getBoundingClientRect();
      return { fx: f.left + 2, fy: f.top + f.height / 2, lx: l.right - 2, ly: l.top + l.height / 2 };
    });
    await page.mouse.move(boxes.fx, boxes.fy);
    await page.mouse.down();
    await page.mouse.move((boxes.fx + boxes.lx) / 2, (boxes.fy + boxes.ly) / 2, { steps: 5 });
    await page.mouse.move(boxes.lx, boxes.ly, { steps: 5 });
    await page.mouse.up();

    // The selection now spans >1 block; bold applies to all three.
    await page.evaluate(() => (window as any).__mbr.editor.format('bold'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__mbr;
      const read = (root: HTMLElement) =>
        (Array.from(root.querySelectorAll('[data-anchor]')) as HTMLElement[])
          .filter((p) => /[ABC]{3}/.test(p.textContent || ''))
          .map((p) => {
            const sp = p.querySelector('span');
            return { bold: sp ? parseInt(getComputedStyle(sp).fontWeight, 10) >= 600 : false };
          });
      const live = read(container);
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div'); document.body.appendChild(c2);
      const e2 = (window as any).Docxodus.DocxEditor.open(c2, saved, (window as any).Docxodus, {});
      const reopened = read(c2);
      editor.close(); e2.close(); container.remove(); c2.remove();
      return { live, reopened };
    });
    expect(out.live.length).toBe(3);
    expect(out.live.every((b: any) => b.bold)).toBe(true);
    expect(out.reopened.length).toBe(3);
    expect(out.reopened.every((b: any) => b.bold)).toBe(true);
  });
});
