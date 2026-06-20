import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Cross-block destructive editing beyond delete: typing over a multi-block selection replaces it;
// Enter over a multi-block selection collapses it then splits at the caret.
test.describe('DocxEditor — cross-block type-over and Enter', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  async function build(page: Page, id: string, lines: string[]) {
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
    for (let i = 0; i < lines.length; i++) {
      await page.keyboard.type(lines[i]);
      if (i < lines.length - 1) await page.keyboard.press('Enter');
    }
    await page.evaluate((cid) => (window as any)[cid].editor.commitAllDirty(), id);
  }

  async function selectAcross(page: Page, id: string, a: string, ao: number, b: string, bo: number) {
    await page.evaluate(({ cid, a, ao, b, bo }) => {
      const { container } = (window as any)[cid];
      const blocks = Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[];
      const textOf = (re: RegExp) => {
        const blk = blocks.find((p) => re.test(p.textContent || ''))!;
        return document.createTreeWalker(blk, NodeFilter.SHOW_TEXT).nextNode() as Text;
      };
      const an = textOf(new RegExp(a)); const bn = textOf(new RegExp(b));
      const r = document.createRange(); r.setStart(an, ao); r.setEnd(bn, bo);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    }, { cid: id, a, ao, b, bo });
  }

  const texts = (cid: string) =>
    (Array.from((window as any)[cid].container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
      .map((p) => (p.textContent || '').trim()).filter((t) => t.length > 0);

  test('typing over a multi-block selection replaces it with the typed text', async ({ page }) => {
    await build(page, 'ov1', ['Alpha', 'Bravo', 'Charlie']);
    await selectAcross(page, 'ov1', 'Alpha', 2, 'Charlie', 2); // "Al[…Ch]arlie"
    await page.keyboard.type('Z');

    const out = await page.evaluate((fn) => {
      const get = new Function('cid', `return (${fn})(cid)`) as (cid: string) => string[];
      const { editor, container } = (window as any).ov1;
      const live = get('ov1');
      const saved: Uint8Array = editor.save();
      const reopened = (window as any).Docxodus.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse((window as any).Docxodus.DocxSessionBridge.Project(reopened)).markdown as string;
      editor.close(); container.remove();
      return { live, md };
    }, texts.toString());

    expect(out.live).toEqual(['AlZarlie']); // "Al" + "Z" + "arlie"
    expect(out.md).toContain('AlZarlie');
  });

  test('Enter over a multi-block selection collapses it then splits at the caret', async ({ page }) => {
    await build(page, 'en1', ['Alpha', 'Bravo', 'Charlie']);
    await selectAcross(page, 'en1', 'Alpha', 2, 'Charlie', 2); // "Al[…Ch]arlie"
    await page.keyboard.press('Enter');

    const out = await page.evaluate((fn) => {
      const get = new Function('cid', `return (${fn})(cid)`) as (cid: string) => string[];
      const { editor, container } = (window as any).en1;
      const live = get('en1');
      const saved: Uint8Array = editor.save();
      const reopened = (window as any).Docxodus.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse((window as any).Docxodus.DocxSessionBridge.Project(reopened)).markdown as string;
      editor.close(); container.remove();
      return { live, md };
    }, texts.toString());

    expect(out.live).toEqual(['Al', 'arlie']); // selection deleted, then split at the join
  });

  test('one undo reverses a cross-block delete (round-trip to pre-edit)', async ({ page }) => {
    await build(page, 'un1', ['Alpha', 'Bravo', 'Charlie']);
    const mdBefore = await page.evaluate(() => {
      const { editor } = (window as any).un1;
      const saved: Uint8Array = editor.save();
      const r = (window as any).Docxodus.DocxSessionBridge.OpenSession(saved, '');
      return JSON.parse((window as any).Docxodus.DocxSessionBridge.Project(r)).markdown as string;
    });
    await selectAcross(page, 'un1', 'Alpha', 2, 'Charlie', 2);
    await page.keyboard.press('Delete');
    const mdAfterUndo = await page.evaluate(() => {
      const { editor, container } = (window as any).un1;
      editor.undo(); // ONE undo
      const saved: Uint8Array = editor.save();
      const r = (window as any).Docxodus.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse((window as any).Docxodus.DocxSessionBridge.Project(r)).markdown as string;
      editor.close(); container.remove();
      return md;
    });
    expect(mdAfterUndo).toBe(mdBefore);
  });
});
