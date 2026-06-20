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

  test('shift+click selects across blocks and applies a paragraph op to all of them', async ({ page }) => {
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'mbr2';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__mbr2 = { editor, container };
      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      const r = document.createRange(); r.selectNodeContents(first);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await page.keyboard.type('one');
    await page.keyboard.press('Enter');
    await page.keyboard.type('two');
    await page.keyboard.press('Enter');
    await page.keyboard.type('three');
    await page.evaluate(() => (window as any).__mbr2.editor.commitAllDirty());

    // Shift+Click from block 1 to block 3 (a real cross-block extend gesture).
    const pts = await page.evaluate(() => {
      const { container } = (window as any).__mbr2;
      const bs = Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
        .filter((p: any) => /one|two|three/.test(p.textContent || '')) as HTMLElement[];
      const a = bs[0].getBoundingClientRect(); const c = bs[bs.length - 1].getBoundingClientRect();
      return { ax: a.left + 2, ay: a.top + a.height / 2, cx: c.right - 2, cy: c.top + c.height / 2 };
    });
    await page.mouse.click(pts.ax, pts.ay);
    await page.keyboard.down('Shift');
    await page.mouse.click(pts.cx, pts.cy);
    await page.keyboard.up('Shift');

    await page.evaluate(() => (window as any).__mbr2.editor.setAlignment('center'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__mbr2;
      const read = (root: HTMLElement) =>
        (Array.from(root.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
          .filter((p) => /one|two|three/.test(p.textContent || ''))
          .map((p) => getComputedStyle(p).textAlign);
      const live = read(container);
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div'); document.body.appendChild(c2);
      const e2 = (window as any).Docxodus.DocxEditor.open(c2, saved, (window as any).Docxodus, {});
      const reopened = read(c2);
      editor.close(); e2.close(); container.remove(); c2.remove();
      return { live, reopened };
    });
    // Shift+click reached a multi-block selection; alignment applied to all three (and round-trips).
    expect(out.live.length).toBe(3);
    expect(out.live.every((a: string) => a === 'center')).toBe(true);
    expect(out.reopened.every((a: string) => a === 'center')).toBe(true);
  });

  test('shift+ArrowDown extends the selection across blocks for a formatting command', async ({ page }) => {
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'mbr3';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__mbr3 = { editor, container };
      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      const r = document.createRange(); r.selectNodeContents(first);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await page.keyboard.type('alpha');
    await page.keyboard.press('Enter');
    await page.keyboard.type('beta');
    await page.evaluate(() => (window as any).__mbr3.editor.commitAllDirty());

    // Put the caret at the start of the first block, then Shift+ArrowDown + Shift+End to extend
    // the selection across into the second block.
    await page.evaluate(() => {
      const { container } = (window as any).__mbr3;
      const first = Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
        .find((p: any) => /alpha/.test(p.textContent || '')) as HTMLElement;
      const tn = document.createTreeWalker(first, NodeFilter.SHOW_TEXT).nextNode()!;
      const r = document.createRange(); r.setStart(tn, 0); r.collapse(true);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+End');

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__mbr3;
      const spans = window.getSelection()!.toString();
      editor.format('bold');
      const bolded = (Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
        .filter((p) => /alpha|beta/.test(p.textContent || ''))
        .map((p) => { const sp = p.querySelector('span'); return sp ? parseInt(getComputedStyle(sp).fontWeight, 10) >= 600 : false; });
      editor.close(); container.remove();
      return { crossBlock: spans.includes('alpha') && spans.includes('beta'), bolded };
    });
    // The shift+arrow selection genuinely spanned both blocks (single contenteditable root).
    expect(out.crossBlock).toBe(true);
    expect(out.bolded.length).toBe(2);
    expect(out.bolded.every((b: boolean) => b)).toBe(true);
  });
});
