import { test, expect, Page } from '@playwright/test';

// Round-6 smoke-test fixes (complex-filing drafting): inserted tables inherit the document font,
// insert-above for tables, single-block format selection chaining, collapsed-caret-in-cell format.

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

async function openBlank(page: Page) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'r6';
    document.body.appendChild(container);
    const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
    (window as any).__r6 = { editor, container, D };
    (container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement).focus();
  });
}

// Make the (single) active block Times via a whole-line selection, like a user would.
async function makeBlockTimes(page: Page, text: string) {
  await page.keyboard.type(text);
  await page.evaluate((t) => {
    const { editor, container } = (window as any).__r6;
    const blk = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
    blk.focus();
    const r = document.createRange(); r.selectNodeContents(blk);
    const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
    editor.setFontFamily('Times New Roman');
    return t;
  }, text);
}

test.describe('DocxEditor — round-6 table font + polish', () => {
  // The headline fix: a table inserted into a Times document has Times cells, not Calibri.
  test('inserted table inherits the document font on seeded cells', async ({ page }) => {
    await openBlank(page);
    await makeBlockTimes(page, 'Body');

    const fonts = await page.evaluate(() => {
      const { editor, container } = (window as any).__r6;
      (container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement).focus();
      editor.insertTable(2, 2, { borderless: true, cellContents: ['A', 'B', 'C', 'D'] });
      return (Array.from(container.querySelectorAll('table td span')) as HTMLElement[])
        .map((s) => getComputedStyle(s).fontFamily);
    });
    expect(fonts.length).toBeGreaterThan(0);
    expect(fonts.every((f) => /Times New Roman/.test(f))).toBe(true);
  });

  // The critical typed-later path: typing into an EMPTY cell of a font-inherited table produces
  // text in that font (and it survives save) — the editor's grid-picker flow inserts empty cells.
  test('typing into an empty cell inherits the table font and survives save', async ({ page }) => {
    await openBlank(page);
    await makeBlockTimes(page, 'Body');

    const out = await page.evaluate(() => {
      const { editor, container, D } = (window as any).__r6;
      (container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement).focus();
      editor.insertTable(1, 2, { borderless: true }); // empty cells
      const cell = container.querySelector('table td p[data-anchor]') as HTMLElement;
      cell.focus();
      const r = document.createRange(); r.selectNodeContents(cell);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
      document.execCommand('insertText', false, 'Texas');
      cell.dispatchEvent(new Event('blur'));

      const rendered = (Array.from(container.querySelectorAll('table td span')) as HTMLElement[])
        .some((el) => (el.textContent || '').includes('Texas') && /Times New Roman/.test(getComputedStyle(el).fontFamily));

      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div'); document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const reopened = (Array.from(c2.querySelectorAll('table td span')) as HTMLElement[])
        .some((el) => (el.textContent || '').includes('Texas') && /Times New Roman/.test(getComputedStyle(el).fontFamily));
      e2.close(); c2.remove();
      return { rendered, reopened };
    });
    expect(out.rendered).toBe(true);
    expect(out.reopened).toBe(true);
  });

  // Insert-above: a table can be placed BEFORE a non-empty block (position 'above').
  test('insertTable(position="above") places the table before a non-empty block', async ({ page }) => {
    await openBlank(page);
    await page.keyboard.type('Heading');
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__r6;
      const blk = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      blk.focus();
      editor.insertTable(1, 1, { borderless: true }, 'above');
      const table = container.querySelector('table');
      const heading = (Array.from(container.querySelectorAll('p[data-editable="1"]')) as HTMLElement[])
        .find((p) => (p.textContent || '').trim() === 'Heading');
      if (!table || !heading) return { hasTable: !!table, hasHeading: !!heading, tableBeforeHeading: false };
      // DOCUMENT_POSITION_FOLLOWING (4) => heading comes after the table in document order.
      const tableBeforeHeading = (table.compareDocumentPosition(heading) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0;
      return { hasTable: true, hasHeading: true, tableBeforeHeading };
    });
    expect(out.hasTable).toBe(true);
    expect(out.hasHeading).toBe(true);
    expect(out.tableBeforeHeading).toBe(true);
  });

  // Single-block selection chaining: a whole-block format must KEEP the selection (not collapse to
  // a caret) so the next command can chain — mirroring the multi-block path.
  test('single-block whole-block format keeps the selection for chaining', async ({ page }) => {
    await openBlank(page);
    await page.keyboard.type('Hello');

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__r6;
      const blk = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      blk.focus();
      const r = document.createRange(); r.selectNodeContents(blk);
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
      editor.setFontFamily('Times New Roman');
      // The selection must survive the re-render (this is the fix — stale collapses to a caret).
      const selAfterFont = window.getSelection()!;
      const collapsedAfterFont = selAfterFont.isCollapsed;
      // NO re-select — the chained command must still cover the line.
      editor.format('bold');
      const span = container.querySelector('p[data-anchor][data-editable="1"] span') as HTMLElement;
      const cs = span ? getComputedStyle(span) : null;
      return {
        collapsedAfterFont,
        times: cs ? /Times New Roman/.test(cs.fontFamily) : false,
        bold: cs ? parseInt(cs.fontWeight, 10) >= 600 : false,
      };
    });
    expect(out.collapsedAfterFont).toBe(false); // selection preserved (the chaining fix)
    expect(out.times).toBe(true);
    expect(out.bold).toBe(true);
  });

  // A collapsed caret inside a cell + a format command applies to the WHOLE cell (no fragile
  // line-selection needed in wrapped cells).
  test('collapsed caret in a cell + bold formats the whole cell', async ({ page }) => {
    await openBlank(page);
    const bold = await page.evaluate(() => {
      const { editor, container } = (window as any).__r6;
      (container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement).focus();
      editor.insertTable(1, 1, { borderless: true, cellContents: ['Texas'] });
      const cell = container.querySelector('table td p[data-anchor]') as HTMLElement;
      cell.focus();
      const r = document.createRange(); r.selectNodeContents(cell); r.collapse(true); // collapsed caret
      const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
      editor.format('bold');
      const span = container.querySelector('table td span') as HTMLElement;
      return span ? parseInt(getComputedStyle(span).fontWeight, 10) >= 600 : false;
    });
    expect(bold).toBe(true);
  });
});
