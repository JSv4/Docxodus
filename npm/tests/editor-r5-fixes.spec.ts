import { test, expect, Page } from '@playwright/test';

// Round-5 S-1 smoke-test fixes. Each test reproduces a bug found while drafting an S-1 in the
// live editor, WITHOUT the workarounds the older specs use (commit-first / re-select).

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

async function openBlank(page: Page) {
  await page.goto('/test-harness.html');
  await waitForDocxodus(page);
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'r5';
    document.body.appendChild(container);
    const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
    (window as any).__r5 = { editor, container, D };
    const first = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
    first.focus();
  });
}

function selectAllBlocks(page: Page) {
  return page.evaluate(() => {
    const { container } = (window as any).__r5;
    const blocks = Array.from(
      container.querySelectorAll('p[data-anchor][data-editable="1"]'),
    ) as HTMLElement[];
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
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
  });
}

const readBold = (page: Page) =>
  page.evaluate(() => {
    const { container } = (window as any).__r5;
    return (Array.from(container.querySelectorAll('p[data-anchor][data-editable="1"]')) as HTMLElement[])
      .map((p) => {
        const span = p.querySelector('span');
        return {
          text: (p.textContent || '').trim(),
          bold: span ? parseInt(getComputedStyle(span).fontWeight, 10) >= 600 : false,
          align: getComputedStyle(p).textAlign,
        };
      });
  });

test.describe('DocxEditor — round-5 S-1 fixes', () => {
  // Bug 1: multi-block inline format must NOT skip the last block when its text was just typed and
  // not yet committed (the natural "type lines → select all → Bold" gesture).
  test('multi-block Bold formats the just-typed (uncommitted) last block', async ({ page }) => {
    await openBlank(page);
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.keyboard.press('Enter');
    await page.keyboard.type('CCC'); // NOT committed — caret left in CCC

    await selectAllBlocks(page);
    await page.evaluate(() => (window as any).__r5.editor.format('bold'));

    const blocks = await readBold(page);
    expect(blocks.map((b) => b.text)).toEqual(['AAA', 'BBB', 'CCC']);
    expect(blocks.every((b) => b.bold)).toBe(true); // ALL three, including the last
  });

  // Bug 2: editor.save() must flush uncommitted (dirty) blocks so a programmatic save (no blur)
  // doesn't silently drop in-progress text.
  test('save() flushes uncommitted text typed just before saving', async ({ page }) => {
    await openBlank(page);
    await page.keyboard.type('Registration No. 333-'); // typed, NOT committed (no blur / caret-leave)

    const reopened = await page.evaluate(() => {
      const { editor, D } = (window as any).__r5;
      const saved: Uint8Array = editor.save(); // programmatic save while focus is still in the block
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const text = (c2.textContent || '').trim();
      e2.close(); c2.remove();
      return text;
    });
    expect(reopened).toContain('Registration No. 333-');
  });

  // UX-D: a multi-block format must keep the selection so commands can be chained (Bold then Center
  // without re-selecting).
  test('multi-block format keeps the selection so the next command also applies to all', async ({ page }) => {
    await openBlank(page);
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.keyboard.press('Enter');
    await page.keyboard.type('CCC');
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());

    await selectAllBlocks(page);
    await page.evaluate(() => (window as any).__r5.editor.format('bold'));
    // Selection must still cover all three blocks.
    const selCount = await page.evaluate(() => (window as any).__r5.editor.selectedBlocks().length);
    expect(selCount).toBe(3);
    // Chained op (no re-select) applies to all three.
    await page.evaluate(() => (window as any).__r5.editor.setAlignment('center'));

    const blocks = await readBold(page);
    expect(blocks.every((b) => b.bold)).toBe(true);
    expect(blocks.every((b) => b.align === 'center')).toBe(true);
  });

  // Bug 3 (verification): inserting a table row, then Undo via the editor's OWN undo path, must
  // remove the row and keep the cell text — table structural CRUD is undoable.
  test('table row insert is undone by editor.undo (cell text preserved)', async ({ page }) => {
    await openBlank(page);
    const result = await page.evaluate(() => {
      const { editor, container } = (window as any).__r5;
      const firstCellAnchor = () =>
        (container.querySelector('table td p[data-anchor]') as HTMLElement) || null;
      // Insert a 1x2 table on the blank paragraph, seed both cells.
      (container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement).focus();
      editor.insertTable(1, 2, { borderless: true });
      const cells = Array.from(container.querySelectorAll('table td p[data-anchor]')) as HTMLElement[];
      const setCell = (el: HTMLElement, text: string) => {
        el.focus();
        const r = document.createRange(); r.selectNodeContents(el);
        const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r);
        document.execCommand('insertText', false, text);
        el.dispatchEvent(new Event('blur'));
      };
      setCell(cells[0], 'LeftCell');
      setCell(cells[1], 'RightCell');
      // Insert a row below the first cell's row.
      (container.querySelector('table td p[data-anchor]') as HTMLElement).focus();
      editor.insertTableRow('below');
      const rowsAfterInsert = container.querySelectorAll('table tr').length;
      // Undo via the editor (NOT a raw bridge call).
      editor.undo();
      const rowsAfterUndo = container.querySelectorAll('table tr').length;
      const cellText = Array.from(container.querySelectorAll('table td')).map((td: any) => (td.textContent || '').trim());
      return { rowsAfterInsert, rowsAfterUndo, cellText };
    });
    expect(result.rowsAfterInsert).toBe(2);
    expect(result.rowsAfterUndo).toBe(1);
    expect(result.cellText.join('|')).toContain('LeftCell');
    expect(result.cellText.join('|')).toContain('RightCell');
  });

  // Bug 1b: the single-block format path (setFontFamily / setFontSize / format) must clamp the
  // selection span to the committed run. A line typed with real keystrokes keeps a placeholder
  // space in the DOM, so selecting the whole line yields a span that overshoots the trimmed run
  // the engine holds — the font/size silently no-ops. (The S-1 company name kept rendering Calibri.)
  test('setFontFamily applies to a whole line whose DOM carries a trailing placeholder space', async ({ page }) => {
    await openBlank(page);
    const out = await page.evaluate(() => {
      const { editor, container, D } = (window as any).__r5;
      const blk = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      blk.focus();
      const sel = window.getSelection()!;
      let r = document.createRange(); r.selectNodeContents(blk); r.collapse(true);
      sel.removeAllRanges(); sel.addRange(r);
      // contenteditable typing (execCommand, like the demo) leaves the empty paragraph's placeholder
      // space — DOM content is "<text> " (one longer than the committed, trimmed run).
      document.execCommand('insertText', false, 'Space Exploration Technologies Corp.');
      const blk2 = container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement;
      const domLen = blk2.textContent!.length;
      r = document.createRange(); r.selectNodeContents(blk2); // whole line incl. the placeholder
      sel.removeAllRanges(); sel.addRange(r);
      editor.setFontFamily('Times New Roman');
      const rendered = Array.from(container.querySelectorAll('span')).some((el: any) =>
        getComputedStyle(el).fontFamily.includes('Times New Roman'));
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div'); document.body.appendChild(c2);
      const e2 = D.DocxEditor.open(c2, saved, D, {});
      const reopened = Array.from(c2.querySelectorAll('span')).some((el: any) =>
        getComputedStyle(el).fontFamily.includes('Times New Roman'));
      e2.close(); c2.remove();
      return { rendered, reopened, domLen };
    });
    expect(out.domLen).toBeGreaterThan('Space Exploration Technologies Corp.'.length); // placeholder present
    expect(out.rendered).toBe(true);
    expect(out.reopened).toBe(true);
  });

  // Gap-G (verification): per-cell alignment via the ribbon (the S-1 filing header needs a
  // left cell and a right cell). setAlignment on a focused cell paragraph aligns just that cell.
  test('per-cell alignment: a left cell and a right cell in the same row', async ({ page }) => {
    await openBlank(page);
    const aligns = await page.evaluate(() => {
      const { editor, container } = (window as any).__r5;
      (container.querySelector('p[data-anchor][data-editable="1"]') as HTMLElement).focus();
      editor.insertTable(1, 2, { borderless: true });
      const cells = Array.from(container.querySelectorAll('table td p[data-anchor]')) as HTMLElement[];
      cells[0].focus(); editor.setAlignment('left');
      const refreshed = Array.from(container.querySelectorAll('table td p[data-anchor]')) as HTMLElement[];
      refreshed[1].focus(); editor.setAlignment('right');
      const final = Array.from(container.querySelectorAll('table td p[data-anchor]')) as HTMLElement[];
      return [getComputedStyle(final[0]).textAlign, getComputedStyle(final[1]).textAlign];
    });
    expect(aligns[0]).toBe('left');
    expect(aligns[1]).toBe('right');
  });
});
