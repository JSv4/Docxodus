import { test, expect, Page } from '@playwright/test';

/**
 * Fixes for the P1/P2 issues found in the legal-contract smoke test:
 *
 *  - P1  — applying a font to a legal-numbered list item left its generated NUMBER MARKER in the
 *          document default font (the converter renders the marker, not the run), so a re-fonted
 *          clause "looked" un-fonted. The content runs were always fonted; only the marker lagged.
 *  - P2#6 — a paragraph-level op (alignment/indent) followed by an inline toggle (bold) on the same
 *          block silently no-opped, because the op's re-render only `focus()`ed the block instead of
 *          restoring a real caret, so the chained toggle read stale state.
 *  - P2#5 — list level could only be changed by Tab/Shift+Tab; a public `changeListLevel(delta)`
 *          lets the demo expose demote/promote buttons.
 */

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('Editor — list marker font, op chaining, and list-level control', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // P1
  test('font applied to a list item also restyles its number marker', async ({ page }) => {
    const out = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor: any = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const setCaret = (el: HTMLElement) => {
        el.focus();
        const r = document.createRange();
        r.selectNodeContents(el);
        r.collapse(false);
        const s = getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
        document.dispatchEvent(new Event('selectionchange'));
      };
      const body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.textContent = 'Section heading';
      body.dispatchEvent(new InputEvent('input', { bubbles: true }));
      setCaret(body);
      editor.toggleLegalNumbering();
      // collapsed caret in the (re-rendered) list item, then change its font + size
      setCaret(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.setFontFamily('Georgia');
      setCaret(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.setFontSize(20);
      const li = container.querySelector('p[data-anchor]') as HTMLElement;
      const marker = li.querySelector('[data-list-marker]') as HTMLElement;
      const content = Array.from(li.querySelectorAll('span')).find((s) =>
        /Section/.test(s.textContent || ''),
      ) as HTMLElement;
      const res = {
        contentFont: getComputedStyle(content).fontFamily,
        markerFont: getComputedStyle(marker).fontFamily,
        contentSize: getComputedStyle(content).fontSize,
        markerSize: getComputedStyle(marker).fontSize,
      };
      editor.close();
      container.remove();
      return res;
    });
    expect(out.contentFont).toContain('Georgia'); // content text is fonted (always was)
    expect(out.markerFont).toContain('Georgia'); // the number marker now follows the run font (fix)
    expect(out.markerSize).toBe(out.contentSize); // marker also follows the run size
  });

  // P2 #6
  test('inline format chains after a paragraph-level op (no dropped caret)', async ({ page }) => {
    const weights = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor: any = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const firstSpan = () =>
        getComputedStyle(container.querySelector('p[data-anchor] span') as HTMLElement).fontWeight;
      const selectWhole = (el: HTMLElement) => {
        el.focus();
        const r = document.createRange();
        r.selectNodeContents(el);
        const s = getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
        document.dispatchEvent(new Event('selectionchange'));
      };
      const setCaret = (el: HTMLElement) => {
        el.focus();
        const r = document.createRange();
        r.selectNodeContents(el);
        r.collapse(false);
        const s = getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
        document.dispatchEvent(new Event('selectionchange'));
      };
      const body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.textContent = 'Hello world';
      body.dispatchEvent(new InputEvent('input', { bubbles: true }));
      selectWhole(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.format('bold');
      const afterBold = firstSpan();
      // collapsed caret, paragraph op (re-render), then a chained inline toggle
      setCaret(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.setAlignment('center');
      editor.format('bold'); // should toggle bold OFF on the whole block
      const afterToggle = firstSpan();
      editor.close();
      container.remove();
      return { afterBold, afterToggle };
    });
    expect(weights.afterBold).toBe('700'); // block became bold
    expect(weights.afterToggle).toBe('400'); // chained toggle removed bold (fix)
  });

  // P2 #5
  test('changeListLevel(delta) promotes/demotes a list item', async ({ page }) => {
    const levels = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor: any = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const setCaret = (el: HTMLElement) => {
        el.focus();
        const r = document.createRange();
        r.selectNodeContents(el);
        r.collapse(false);
        const s = getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
        document.dispatchEvent(new Event('selectionchange'));
      };
      const ilvlOf = () => {
        const p = container.querySelector('p[data-anchor]') as HTMLElement;
        const fullId = editor.unidToFullId.get(p.getAttribute('data-anchor'));
        try {
          return JSON.parse(D.DocxSessionBridge.GetListMembership(editor.handle, fullId)).level;
        } catch {
          return null;
        }
      };
      const body = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      body.textContent = 'Clause';
      body.dispatchEvent(new InputEvent('input', { bubbles: true }));
      setCaret(body);
      editor.toggleLegalNumbering();
      const level0 = ilvlOf();
      setCaret(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.changeListLevel(1); // demote
      const level1 = ilvlOf();
      setCaret(container.querySelector('p[data-anchor]') as HTMLElement);
      editor.changeListLevel(-1); // promote back
      const back0 = ilvlOf();
      editor.close();
      container.remove();
      return { level0, level1, back0 };
    });
    expect(levels.level0).toBe(0);
    expect(levels.level1).toBe(1); // demoted one level (fix)
    expect(levels.back0).toBe(0); // promoted back
  });
});
