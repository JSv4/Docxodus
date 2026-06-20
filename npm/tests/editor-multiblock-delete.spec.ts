import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Cross-block destructive editing: a Delete/Backspace over a multi-block selection trims the first
// block, removes whole middle blocks, and merges the trimmed last block into the first — as one
// atomic undo.
test.describe('DocxEditor — cross-block delete', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // Build the given lines in a blank doc (one paragraph each), keep {editor, container} on window[id].
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

  // Select from (textOf line `a`, offset `ao`) to (textOf line `b`, offset `bo`).
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

  test('Delete across three blocks trims first, removes middle, merges last', async ({ page }) => {
    await build(page, 'del1', ['Alpha', 'Bravo', 'Charlie']);
    await selectAcross(page, 'del1', 'Alpha', 2, 'Charlie', 2); // "Al[pha … Ch]arlie" → "Al" + "arlie"
    await page.keyboard.press('Delete');

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).del1;
      const texts = () =>
        (Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
          .map((p) => (p.textContent || '').trim()).filter((t) => t.length > 0);
      const live = texts();
      const saved: Uint8Array = editor.save();
      const reopened = (window as any).Docxodus.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse((window as any).Docxodus.DocxSessionBridge.Project(reopened)).markdown as string;
      editor.close(); container.remove();
      return { live, md };
    });

    expect(out.live).toEqual(['Alarlie']);        // "Al" + "arlie" (seamless join), Bravo gone
    expect(out.md).toContain('Alarlie');
    expect(out.md).not.toContain('Bravo');
  });

  test('Backspace deleting whole middle blocks keeps the untouched outer paragraphs', async ({ page }) => {
    await build(page, 'del2', ['Intro', 'XXXX', 'YYYY', 'Outro']);
    // Select from end of Intro to start of Outro (covers all of XXXX and YYYY, nothing of Intro/Outro).
    await selectAcross(page, 'del2', 'Intro', 5, 'Outro', 0);
    await page.keyboard.press('Backspace');

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).del2;
      const texts = (Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
        .map((p) => (p.textContent || '').trim()).filter((t) => t.length > 0);
      const saved: Uint8Array = editor.save();
      const reopened = (window as any).Docxodus.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse((window as any).Docxodus.DocxSessionBridge.Project(reopened)).markdown as string;
      editor.close(); container.remove();
      return { texts, md };
    });

    // Intro and Outro merge into one paragraph ("IntroOutro"); XXXX and YYYY are gone.
    expect(out.texts).toEqual(['IntroOutro']);
    expect(out.md).toContain('IntroOutro');
    expect(out.md).not.toContain('XXXX');
    expect(out.md).not.toContain('YYYY');
  });
});
