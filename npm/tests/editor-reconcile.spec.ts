import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Structural ops reconcile the DOM incrementally (unit-sequence diff + batch block
// render) instead of remounting the whole document. These tests pin:
//   1. node identity — untouched blocks keep their DOM nodes across structural ops;
//   2. remount equivalence — the reconciled DOM matches what a full remount builds;
//   3. footnote chrome — markers/li values renumber correctly after note inserts;
//   4. container substitution — a row insert repaints the table (its unid is stable
//      but its content signature moved) without touching the rest of the document.
test.describe('DocxEditor — incremental structural reconcile', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  /** Fresh editor over a blank doc; types `names` as successive paragraphs. */
  async function buildParagraphs(page: Page, names: string[]) {
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'reconcile-host';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__rec = { editor, container };
      const first = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      first.focus();
    });
    for (let i = 0; i < names.length; i++) {
      if (i > 0) await page.keyboard.press('Enter');
      await page.keyboard.type(names[i]);
    }
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());
    // Full paint so every unit carries its content-signature stamp. Typing swaps
    // blocks with fresh (unstamped) nodes; unstamped units legitimately re-render on
    // the next reconcile, which would confuse the node-identity assertions below.
    await page.evaluate(() => (window as any).__rec.editor['remount']());
  }

  const focusBlock = (page: Page, index: number, atEnd = true) =>
    page.evaluate(({ index, atEnd }) => {
      const { editor } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      const el = blocks[index];
      el.focus();
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.selectNodeContents(el);
      r.collapse(!atEnd);
      sel.removeAllRanges();
      sel.addRange(r);
    }, { index, atEnd });

  /** Serialize the body edit root minus editor-added attributes (whose insertion
   *  order differs between mount-time stamping and reconcile-time construction). */
  const normalizedBody = (page: Page) =>
    page.evaluate(() => {
      const { editor } = (window as any).__rec;
      const clone = (editor['editRoot'] as HTMLElement).cloneNode(true) as HTMLElement;
      clone.querySelectorAll('style').forEach((s) => s.remove());
      for (const el of Array.from(clone.querySelectorAll('*'))) {
        el.removeAttribute('data-committed-text');
        el.removeAttribute('contenteditable');
        el.removeAttribute('data-render-sig');
        el.removeAttribute('data-note-anchor');
      }
      return clone.innerHTML;
    });

  test('insertTable reconciles: untouched blocks keep their DOM nodes', async ({ page }) => {
    await buildParagraphs(page, ['AAA', 'BBB', 'CCC', 'DDD']);
    await page.evaluate(() => {
      const { editor } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      (blocks[3] as any).__sentinel = 'survivor';
    });
    await focusBlock(page, 1);
    await page.evaluate(() => (window as any).__rec.editor.insertTable(2, 2));

    const r = await page.evaluate(() => {
      const { editor, container } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      return {
        survived: blocks.some((b: any) => b.__sentinel === 'survivor'),
        hasTable: !!container.querySelector('table[data-anchor]'),
        tableSigStamped: !!container.querySelector('table[data-render-sig]'),
        cellEditable: !!container.querySelector('table td p[contenteditable="true"]'),
      };
    });
    expect(r.survived).toBe(true); // a remount would have rebuilt every node
    expect(r.hasTable).toBe(true);
    expect(r.tableSigStamped).toBe(true);
    expect(r.cellEditable).toBe(true);
  });

  test('undo reconciles to a DOM equivalent to a full remount', async ({ page }) => {
    await buildParagraphs(page, ['AAA', 'BBB', 'CCC']);
    await focusBlock(page, 1);
    await page.evaluate(() => (window as any).__rec.editor.insertTable(2, 2));
    await page.evaluate(() => (window as any).__rec.editor.undo());
    const afterReconcile = await normalizedBody(page);
    expect(afterReconcile).not.toContain('<table'); // undo actually removed it
    await page.evaluate(() => (window as any).__rec.editor['remount']());
    const afterRemount = await normalizedBody(page);
    expect(afterReconcile).toBe(afterRemount);
  });

  test('redo restores the table incrementally', async ({ page }) => {
    await buildParagraphs(page, ['AAA', 'BBB']);
    await focusBlock(page, 0);
    await page.evaluate(() => (window as any).__rec.editor.insertTable(2, 3));
    await page.evaluate(() => (window as any).__rec.editor.undo());
    await page.evaluate(() => {
      const { editor } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      (blocks[0] as any).__sentinel = 'redo-survivor';
      editor.redo();
    });
    const r = await page.evaluate(() => {
      const { editor, container } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      return {
        survived: blocks.some((b: any) => b.__sentinel === 'redo-survivor'),
        cols: container.querySelectorAll('table tr:first-child td').length,
      };
    });
    expect(r.survived).toBe(true);
    expect(r.cols).toBe(3);
  });

  test('insertTableRow substitutes the table in place (sig change), sparing the rest', async ({ page }) => {
    await buildParagraphs(page, ['AAA', 'BBB']);
    await focusBlock(page, 0);
    await page.evaluate(() => (window as any).__rec.editor.insertTable(2, 2));
    const r = await page.evaluate(() => {
      const { editor, container } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      const outside = blocks.find((b: HTMLElement) => !b.closest('table'))!;
      (outside as any).__sentinel = 'row-survivor';
      const cellP = container.querySelector('table td p[contenteditable="true"]') as HTMLElement;
      cellP.focus();
      editor.insertTableRow('below');
      return {
        rows: container.querySelectorAll('table tr').length,
        survived: (editor['editableList']() as any[]).some((b) => b.__sentinel === 'row-survivor'),
        newCellsEditable:
          container.querySelectorAll('table td p[contenteditable="true"]').length >= 6,
      };
    });
    expect(r.rows).toBe(3);
    expect(r.survived).toBe(true);
    expect(r.newCellsEditable).toBe(true);
  });

  test('footnote inserts renumber marker chrome across the document', async ({ page }) => {
    await buildParagraphs(page, ['First paragraph.', 'Second paragraph.', 'Third paragraph.']);
    // Cite late → then earlier → then earliest, so ids shift and every insert
    // renumbers the chrome of the notes cited after it.
    await focusBlock(page, 2);
    await page.evaluate(() => (window as any).__rec.editor.insertFootnote('Note C.'));
    await focusBlock(page, 1);
    await page.evaluate(() => (window as any).__rec.editor.insertFootnote('Note B.'));
    await focusBlock(page, 0);
    await page.evaluate(() => (window as any).__rec.editor.insertFootnote('Note A.'));

    const r = await page.evaluate(() => {
      const { editor, container } = (window as any).__rec;
      const D = (window as any).Docxodus;
      const markers = Array.from(
        container.querySelectorAll('a.footnote-ref'),
      ).filter((a: any) => !a.closest('section.footnotes')) as HTMLElement[];
      const lis = Array.from(
        container.querySelectorAll('section.footnotes > ol > li'),
      ) as HTMLElement[];
      const notes = JSON.parse(D.DocxSessionBridge.ListNotes(editor['handle'], false));
      return {
        supTexts: markers.map((m) => m.querySelector('sup')?.textContent),
        hrefs: markers.map((m) => m.getAttribute('href')),
        liIds: lis.map((li) => li.id),
        liValues: lis.map((li) => li.getAttribute('value')),
        liTexts: lis.map((li) => (li.textContent ?? '').trim()),
        engineIds: notes.map((n: any) => n.id),
        backrefs: lis.map((li) => li.querySelector('a[class$="backref"]')?.getAttribute('href')),
      };
    });
    // Document-order markers display 1, 2, 3 …
    expect(r.supTexts).toEqual(['1', '2', '3']);
    expect(r.liValues).toEqual(['1', '2', '3']);
    // … and the notes section lists the notes in citation order.
    expect(r.liTexts[0]).toContain('Note A.');
    expect(r.liTexts[1]).toContain('Note B.');
    expect(r.liTexts[2]).toContain('Note C.');
    // Marker k targets note k (engine ids, ascending per the reference-order law).
    expect(r.hrefs).toEqual(r.engineIds.map((id: string) => `#fn-${id}`));
    expect(r.liIds).toEqual(r.engineIds.map((id: string) => `fn-${id}`));
    expect(r.backrefs).toEqual(r.engineIds.map((id: string) => `#fn-ref-${id}`));
  });

  test('the first footnote creates its section without remounting untouched blocks', async ({ page }) => {
    await buildParagraphs(page, ['Cite here.', 'Untouched survivor.']);
    await page.evaluate(() => {
      const { editor } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      (blocks[1] as any).__sentinel = 'first-note-survivor';
    });
    await focusBlock(page, 0);
    await page.evaluate(() => (window as any).__rec.editor.insertFootnote('First note.'));

    const r = await page.evaluate(() => {
      const { editor, container } = (window as any).__rec;
      const blocks = editor['editableList']() as HTMLElement[];
      const section = container.querySelector('section.footnotes');
      const li = section?.querySelector('ol > li') as HTMLElement | null;
      return {
        survived: blocks.some((b: any) => b.__sentinel === 'first-note-survivor'),
        fallback: editor.lastReconcileFallback as string | null,
        sections: container.querySelectorAll('section.footnotes').length,
        noteCount: section?.querySelectorAll(':scope > ol > li').length ?? 0,
        noteStamped: li?.hasAttribute('data-note-anchor') ?? false,
        noteEditable: !!li?.querySelector('[data-anchor][contenteditable="true"]'),
        noteText: li?.textContent ?? '',
      };
    });
    expect(r.survived).toBe(true);
    expect(r.fallback).toBeNull();
    expect(r.sections).toBe(1);
    expect(r.noteCount).toBe(1);
    expect(r.noteStamped).toBe(true);
    expect(r.noteEditable).toBe(true);
    expect(r.noteText).toContain('First note.');
  });

  test('reconciled edits survive save/reopen losslessly', async ({ page }) => {
    await buildParagraphs(page, ['AAA', 'BBB', 'CCC']);
    await focusBlock(page, 1);
    await page.evaluate(() => (window as any).__rec.editor.insertTable(2, 2));
    await focusBlock(page, 0);
    await page.evaluate(() => (window as any).__rec.editor.insertFootnote('Round-trip note.'));

    const r = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const { editor } = (window as any).__rec;
      const bytes = editor.save();
      const host = document.createElement('div');
      document.body.appendChild(host);
      const reopened = D.DocxEditor.open(host, bytes, D, {});
      const out = {
        tables: host.querySelectorAll('table[data-anchor]').length,
        noteLis: host.querySelectorAll('section.footnotes > ol > li').length,
        text: (host.textContent ?? '').replace(/\s+/g, ' '),
      };
      reopened.close();
      return out;
    });
    expect(r.tables).toBe(1);
    expect(r.noteLis).toBe(1);
    expect(r.text).toContain('AAA');
    expect(r.text).toContain('Round-trip note.');
  });
});
