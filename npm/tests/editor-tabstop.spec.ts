import { test, expect, Page } from '@playwright/test';

// Editor command for the real right tab-stop (complex-filing smoke-test #5): a single paragraph holds
// left text + a tab + right-aligned text (a filing masthead's "As filed… / Registration No." row),
// instead of faking it with a two-column table. editor.insertTab('right') must NOT split the
// paragraph and must round-trip (the OOXML correctness itself is covered by the C# DS240-242 tests).

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('DocxEditor — right tab stop', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('insertTab adds a right tab stop + tab run to the active paragraph (no split, round-trips)', async ({ page }) => {
    const r = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});

      const bodyParas = () =>
        Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
          .filter((b) => !(b as HTMLElement).closest('table')) as HTMLElement[];

      // Type the left text and commit it.
      const b0 = bodyParas()[0];
      b0.focus();
      b0.textContent = 'As filed';
      b0.dispatchEvent(new Event('blur'));

      // Caret at the end of the (re-rendered) paragraph's text, then insert a right tab.
      const blk = bodyParas()[0];
      blk.focus();
      const textNode = (document.createTreeWalker(blk, NodeFilter.SHOW_TEXT).nextNode() ?? blk) as Node;
      const sel = window.getSelection()!;
      const range = document.createRange();
      range.setStart(textNode, (textNode.textContent || '').length);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);

      const countBefore = bodyParas().length;
      editor.insertTab('right');
      const countAfter = bodyParas().length;

      const saved: Uint8Array = editor.save();
      editor.close();

      // Re-open and inspect the paragraph's raw XML (cross-layer check that the op landed).
      const h = D.DocxSessionBridge.OpenSession(saved, '');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      const m = (proj.markdown as string).match(/\{#(p:body:[0-9a-f]+)\}[^\n]*As filed/);
      const xml = m ? D.DocxSessionBridge.RawGetXml(h, m[1]) : '';
      D.DocxSessionBridge.CloseSession(h);

      const tabCount = (xml.match(/<w:tab[ />]/g) || []).length;
      return { countBefore, countAfter, hasRightStop: /w:val="right"/.test(xml), tabCount, keptText: /As filed/.test(xml) };
    });

    expect(r.countAfter).toBe(r.countBefore);   // one paragraph, not split into two
    expect(r.hasRightStop).toBe(true);          // a right tab STOP on the paragraph
    expect(r.tabCount).toBeGreaterThanOrEqual(2); // the stop element + the tab RUN
    expect(r.keptText).toBe(true);              // left text preserved
  });
});
