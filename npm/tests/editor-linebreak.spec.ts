import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue 1 — line-break fidelity. Shift+Enter must produce a REAL Word line break
// (w:br), not a literal '\n' in w:t (which Word renders as a space). After commit
// the block re-renders from the live session, where a w:br renders back as <br>;
// the round-tripped projection carries the canonical GFM hard break "  \n".
test.describe('DocxEditor — line-break fidelity (Shift+Enter → w:br)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Shift+Enter inserts a faithful line break that survives save/reopen', async ({ page }) => {
    // Set up a blank doc and focus its first paragraph (select the placeholder so the
    // first typed char replaces it). Keep the editor on window for the second step.
    await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'lb-container';
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});
      (window as any).__lb = { editor, container };
      const target = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      target.focus();
      const r = document.createRange();
      r.selectNodeContents(target);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });

    // REAL keyboard input: triggers the editor's Shift+Enter handler and native typing.
    await page.keyboard.type('AAA');
    await page.keyboard.press('Shift+Enter');
    await page.keyboard.type('BBB');

    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const { editor, container } = (window as any).__lb;
      const firstBlock = () =>
        container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;

      const target = firstBlock();
      const preCommitHTML = target.innerHTML;
      // Commit (blur). The block re-renders from the session.
      target.dispatchEvent(new Event('blur'));

      const committed = firstBlock();
      const brInCommitted = committed.querySelectorAll('br').length;
      const committedText = (committed.textContent || '').replace(/\s+/g, ' ').trim();

      // Save → reopen the bytes → project to markdown (the faithfulness oracle).
      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      // Re-open in a fresh editor: a w:br renders back to <br>.
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const reBlock = c2.querySelector('p[data-anchor]') as HTMLElement;
      const brAfterReopen = reBlock ? reBlock.querySelectorAll('br').length : 0;

      editor.close();
      e2.close();
      container.remove();
      c2.remove();

      return { brInCommitted, committedText, md, brAfterReopen, preCommitHTML };
    });

    // A real line break exists in the committed (re-rendered) block...
    expect(out.brInCommitted).toBeGreaterThanOrEqual(1);
    expect(out.committedText).toContain('AAA');
    expect(out.committedText).toContain('BBB');
    // ...the projected markdown carries the canonical GFM hard break "  \n"...
    expect(out.md).toContain('AAA  \nBBB');
    // ...and it survives a full save → reopen (w:br renders back as <br>).
    expect(out.brAfterReopen).toBeGreaterThanOrEqual(1);
  });

  test('exact 10pt rows keep exact browser geometry with unformatted break runs', async ({ page }) => {
    const result = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const bridge = D.DocxSessionBridge;
      const anchor = JSON.parse(bridge.FindByKind(editor.sessionHandle, 'p', 'body'))[0].id as string;
      const current = bridge.RawGetXml(editor.sessionHandle, anchor) as string;
      const rawOpenTag = current.match(/^<w:p\b[^>]*>/)?.[0];
      if (!rawOpenTag) throw new Error(`paragraph opener missing: ${current.slice(0, 120)}`);
      const openTag = rawOpenTag.endsWith('/>') ? `${rawOpenTag.slice(0, -2)}>` : rawOpenTag;

      const rows = Array.from({ length: 20 }, (_, i) => `ROW ${String(i).padStart(2, '0')}`);
      const textRun = (text: string) =>
        `<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>` +
        `<w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
      // Deliberately leave each break run unformatted. The document default is 11pt,
      // while the visible rows are 8pt and paragraph spacing is exactly 10pt.
      const content = rows.map((row, i) => textRun(row) + (i + 1 < rows.length
        ? '<w:r><w:br/></w:r>'
        : '')).join('');
      const replacement = `${openTag}<w:pPr><w:spacing w:line="200" w:lineRule="exact"/>` +
        `</w:pPr>${content}</w:p>`;
      const edit = JSON.parse(bridge.RawReplaceXml(editor.sessionHandle, anchor, replacement));
      if (!edit.success) throw new Error(JSON.stringify(edit.error));
      editor.refresh();

      const paragraph = container.querySelector('p[data-anchor]') as HTMLElement;
      const rect = paragraph.getBoundingClientRect();
      const breakParents = Array.from(paragraph.querySelectorAll('br')).map((br) => br.parentElement?.tagName);
      const computed = getComputedStyle(paragraph);
      const out = {
        height: rect.height,
        expectedHeight: rows.length * 10 * 96 / 72,
        lineHeight: computed.lineHeight,
        box: {
          paddingTop: computed.paddingTop,
          paddingBottom: computed.paddingBottom,
          borderTop: computed.borderTopWidth,
          borderBottom: computed.borderBottomWidth,
          minHeight: computed.minHeight,
        },
        breakCount: breakParents.length,
        breakParents,
      };
      editor.close();
      container.remove();
      return out;
    });

    expect(parseFloat(result.lineHeight)).toBeCloseTo(10 * 96 / 72, 2);
    expect(result.breakCount).toBe(19);
    expect(result.breakParents.every((tag) => tag === 'P')).toBe(true);
    expect(Math.abs(result.height - result.expectedHeight), JSON.stringify(result)).toBeLessThan(1);
  });
});
