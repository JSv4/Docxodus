import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(rel: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, rel)));
}

async function openInEditor(page: Page, bytes: Uint8Array) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.evaluate((arr: number[]) => {
    (window as any).__nvca = new Uint8Array(arr);
    (window as any).__demo.openDoc(new Uint8Array(arr));
  }, Array.from(bytes));
  await page.waitForFunction(
    () => document.querySelectorAll('#editor [data-anchor]').length > 0,
    { timeout: 90000 },
  );
}

// An editor must be able to RENDER what it can edit. Footnote bodies live in their own OOXML part,
// so the editor's render profile has to emit the notes section and citation markers, and the note
// paragraphs have to carry data-anchor stamps to be editable at all.
test.describe('DocxEditor footnote rendering + editing', () => {
  test('renders citation markers and an editable notes section', async ({ page }) => {
    // WC034 has footnotes; DD001 is the dense bookmark/xref/footnote fixture.
    await openInEditor(page, readTestFile('DD/DD001-DenseBookmarkXrefFootnote.docx'));

    const shape = await page.evaluate(() => {
      const ed = document.getElementById('editor')!;
      return {
        notesSection: !!ed.querySelector('section.footnotes'),
        markers: ed.querySelectorAll('a.footnote-ref').length,
        // every citation marker renders a display NUMBER, not an empty span
        markersNumbered: Array.from(ed.querySelectorAll('a.footnote-ref'))
          .every((a) => /\d/.test(a.textContent ?? '')),
        editableNoteParas: ed.querySelectorAll(
          'section.footnotes [data-anchor][contenteditable="true"]',
        ).length,
        // generated chrome must be inert so the caret can't enter it
        markersInert: Array.from(ed.querySelectorAll('a.footnote-ref'))
          .every((a) => a.getAttribute('contenteditable') === 'false'),
        backrefsInert: Array.from(ed.querySelectorAll('a[class$="-backref"]'))
          .every((a) => a.getAttribute('contenteditable') === 'false'),
      };
    });

    expect(shape.notesSection).toBe(true);
    expect(shape.markers).toBeGreaterThan(0);
    expect(shape.markersNumbered).toBe(true);
    expect(shape.editableNoteParas).toBeGreaterThan(0);
    expect(shape.markersInert).toBe(true);
    expect(shape.backrefsInert).toBe(true);
  });

  test('editing a footnote body persists to the saved document', async ({ page }) => {
    await openInEditor(page, readTestFile('DD/DD001-DenseBookmarkXrefFootnote.docx'));

    const result = await page.evaluate(async () => {
      const demo = (window as any).__demo;
      const bridge = demo.exports.DocxSessionBridge;
      const ed = demo.getEditor();
      const para = document.querySelector<HTMLElement>(
        'section.footnotes [data-anchor][contenteditable="true"]',
      )!;

      // Type at the end of the note's own text (never inside the backref).
      para.focus();
      const texts: Text[] = [];
      const walk = (n: Node) => {
        if (n.nodeType === 3) texts.push(n as Text);
        else n.childNodes.forEach(walk);
      };
      walk(para);
      const last = texts.filter((t) => !t.parentElement!.closest('a[class$="-backref"]')).pop()!;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(last, last.length);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      for (const ch of ' EDITED-IN-EDITOR.') document.execCommand('insertText', false, ch);

      para.blur();
      await new Promise((res) => setTimeout(res, 800));

      const saved = ed.save();
      const h2 = bridge.OpenSession(saved, '');
      try {
        // The projection escapes markdown punctuation, so compare on an unescaped copy.
        const md = JSON.parse(bridge.Project(h2)).markdown.replace(/\\/g, '');
        return { persisted: md.includes('EDITED-IN-EDITOR.'), savedBytes: saved.length };
      } finally {
        bridge.CloseSession(h2);
      }
    });

    expect(result.persisted).toBe(true);
    expect(result.savedBytes).toBeGreaterThan(0);
  });

  test('typing in a citing paragraph keeps the citation and never commits the display number', async ({ page }) => {
    await openInEditor(page, readTestFile('DD/DD001-DenseBookmarkXrefFootnote.docx'));

    const result = await page.evaluate(async () => {
      const demo = (window as any).__demo;
      const bridge = demo.exports.DocxSessionBridge;
      const ed = demo.getEditor();

      const before = JSON.parse(bridge.Project(ed.handle));
      const refsBefore = (before.markdown.match(/\[\^fn-/g) || []).length;

      const marker = document.querySelector<HTMLElement>('#editor a.footnote-ref')!;
      const displayNumber = marker.textContent!.trim();
      const para = marker.closest<HTMLElement>('[data-anchor][contenteditable="true"]')!;

      // Type at the START of the paragraph, so the typed run and the citation coexist.
      para.focus();
      const texts: Text[] = [];
      const walk = (n: Node) => {
        if (n.nodeType === 3) texts.push(n as Text);
        else n.childNodes.forEach(walk);
      };
      walk(para);
      const first = texts.filter((t) => !t.parentElement!.closest('a.footnote-ref'))[0];
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(first, 0);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      for (const ch of 'ZZTOP ') document.execCommand('insertText', false, ch);

      para.blur();
      await new Promise((res) => setTimeout(res, 800));

      const after = JSON.parse(bridge.Project(ed.handle));
      const md = after.markdown.replace(/\\/g, '');
      return {
        typedPersisted: md.includes('ZZTOP '),
        refsBefore,
        refsAfter: (after.markdown.match(/\[\^fn-/g) || []).length,
        // the marker's rendered number must not have been committed as literal text
        strayNumber: md.includes(`ZZTOP ${displayNumber}`),
      };
    });

    expect(result.typedPersisted).toBe(true);
    // The citation survived the edit: same number of note references as before.
    expect(result.refsAfter).toBe(result.refsBefore);
    expect(result.strayNumber).toBe(false);
  });
});
