import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Tables are a selection boundary (v1): a multi-block selection that reaches into a table acts only
// on the body paragraphs, never on the table cells, and never corrupts the table grid.
test.describe('DocxEditor — table is a selection boundary', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('a selection spanning body paragraphs into a table only formats the body blocks', async ({ page }) => {
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div'); container.id = 'tb';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).tb = { editor, container };
      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      const r = document.createRange(); r.selectNodeContents(first);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.evaluate(() => (window as any).tb.editor.commitAllDirty());
    // Insert a 1×1 table with cell text CCC after BBB.
    await page.evaluate(() => (window as any).tb.editor.insertTable(1, 1, { cellContents: ['CCC'] }));

    const rowsBefore = await page.evaluate(() => (window as any).tb.container.querySelectorAll('table tr').length);

    // Build a range from AAA (body) into the CCC table cell, then bold.
    await page.evaluate(() => {
      const { container } = (window as any).tb;
      const aaa = (Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[])
        .find((p) => /AAA/.test(p.textContent || ''))!;
      const ccc = (Array.from(container.querySelectorAll('table [data-anchor]')) as HTMLElement[])
        .find((p) => /CCC/.test(p.textContent || '')) || container.querySelector('table td') as HTMLElement;
      const an = document.createTreeWalker(aaa, NodeFilter.SHOW_TEXT).nextNode()!;
      const cn = document.createTreeWalker(ccc, NodeFilter.SHOW_TEXT).nextNode() || ccc;
      const r = document.createRange();
      r.setStart(an, 0);
      r.setEnd(cn, cn.nodeType === 3 ? (cn.textContent || '').length : cn.childNodes.length);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    });
    await page.evaluate(() => (window as any).tb.editor.format('bold'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).tb;
      const boldOf = (re: RegExp) => {
        const p = (Array.from(container.querySelectorAll('[data-anchor]')) as HTMLElement[])
          .find((e) => re.test(e.textContent || ''));
        const sp = p?.querySelector('span');
        return sp ? parseInt(getComputedStyle(sp).fontWeight, 10) >= 600 : false;
      };
      const result = {
        aaaBold: boldOf(/^AAA/), bbbBold: boldOf(/^BBB/), cccBold: boldOf(/CCC/),
        rowsAfter: container.querySelectorAll('table tr').length,
        cellText: (container.querySelector('table td')?.textContent || '').trim(),
      };
      editor.close(); container.remove();
      return result;
    });

    expect(out.aaaBold).toBe(true);   // body paragraph formatted
    expect(out.bbbBold).toBe(true);   // body paragraph formatted
    expect(out.cccBold).toBe(false);  // table cell NOT formatted (boundary)
    expect(out.rowsAfter).toBe(rowsBefore); // table grid intact
    expect(out.cellText).toContain('CCC'); // cell content intact
  });
});
