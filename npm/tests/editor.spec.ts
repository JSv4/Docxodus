import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// End-to-end DocxEditor (Option B): render faithful pages → edit a block →
// ONLY that block re-renders from the live session → save is lossless for
// untouched content. The complete editor experience, in a real browser.
test.describe('DocxEditor — block editor end-to-end', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('renders, edits one block incrementally, and saves losslessly', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);

      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const editablePs = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]'))
          .filter((e) => norm(e.textContent || '').length > 5) as HTMLElement[];

      const before = editablePs();
      const blockCountBefore = container.querySelectorAll('[data-anchor]').length;

      // Target = first editable paragraph; witness = a different one (must survive).
      const target = before[0];
      const witness = before[before.length - 1];
      const targetUnidBefore = target.getAttribute('data-anchor');
      const witnessText = norm(witness.textContent || '');
      const anchorsBefore = Array.from(container.querySelectorAll('[data-anchor]'))
        .map((e) => e.getAttribute('data-anchor'));

      // Edit the target block and commit (blur fires the editor's commit path).
      const MARKER = 'EDITORLOOP99 brand new content';
      target.focus();
      target.textContent = MARKER;
      target.dispatchEvent(new Event('blur'));

      // After commit: the target block was re-rendered in place (its anchor changed,
      // the witness is untouched). Find the block now carrying the marker.
      const edited = editablePs().find((e) => norm(e.textContent || '').includes('EDITORLOOP99'));
      const editedText = edited ? norm(edited.textContent || '') : '(missing)';
      const editedAnchor = edited ? edited.getAttribute('data-anchor') : null;
      const witnessStillPresent = editablePs().some((e) => norm(e.textContent || '') === witnessText);
      const blockCountAfter = container.querySelectorAll('[data-anchor]').length;

      // How many block anchors changed? Only the edited one should differ.
      const anchorsAfter = Array.from(container.querySelectorAll('[data-anchor]'))
        .map((e) => e.getAttribute('data-anchor'));
      const changed = anchorsBefore.filter((a) => !anchorsAfter.includes(a)).length;

      // Save (lossless) and reopen to confirm persistence + untouched survival.
      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      // Compare on alphanumeric-normalized text — robust to markdown escaping/whitespace.
      const alnum = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      const mdAlnum = alnum(md);

      editor.close();
      container.remove();

      return {
        blockCountBefore,
        blockCountAfter,
        targetUnidBefore,
        editedAnchor,
        editedText,
        witnessText,
        witnessStillPresent,
        anchorsChanged: changed,
        savedLen: saved.length,
        savedHasEdit: mdAlnum.includes('editorloop99'),
        savedHasWitness: mdAlnum.includes(alnum(witnessText).slice(0, 30)),
        editableCount: before.length,
      };
    }, Array.from(bytes));

    // Rendered as many addressable, editable blocks.
    expect(out.editableCount).toBeGreaterThan(1);
    expect(out.blockCountBefore).toBeGreaterThan(1);
    // The edit is visible in the re-rendered block...
    expect(out.editedText).toContain('EDITORLOOP99');
    // ...and the block's anchor is STABLE within the live session (ReplaceText mutates
    // runs in place; the Unid attribute is not re-derived per edit) — so no anchor churn
    // and the editor needs no stable-key remap mid-session.
    expect(out.editedAnchor).toBe(out.targetUnidBefore);
    expect(out.anchorsChanged).toBe(0);
    expect(out.blockCountAfter).toBe(out.blockCountBefore);
    // An untouched block survived in the DOM and in the saved document (lossless).
    expect(out.witnessStillPresent).toBe(true);
    expect(out.savedLen).toBeGreaterThan(0);
    expect(out.savedHasEdit).toBe(true);
    expect(out.savedHasWitness).toBe(true);
  });

  // Paginated mode: blocks flow into real page boxes (margins/headers via the converter
  // + pagination.ts), and those page blocks remain editable with incremental re-render.
  test('paginated mode renders page boxes with editable blocks', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, { paginated: true });

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const pageContainer = container.querySelector('#pagination-container');
      const pageBoxes = container.querySelectorAll('.page-box').length;
      const editable = pageContainer
        ? (Array.from(
            pageContainer.querySelectorAll('p[data-anchor][contenteditable="true"]'),
          ) as HTMLElement[]).filter((e) => norm(e.textContent || '').length > 5)
        : [];

      let editedText = '(none)';
      const target = editable[0];
      if (target) {
        target.focus();
        target.textContent = 'PAGINATEDEDIT77 content';
        target.dispatchEvent(new Event('blur'));
        const edited = (Array.from(
          (pageContainer as HTMLElement).querySelectorAll('p[data-anchor]'),
        ) as HTMLElement[]).find((e) => norm(e.textContent || '').includes('PAGINATEDEDIT77'));
        editedText = edited ? norm(edited.textContent || '') : '(missing)';
      }

      editor.close();
      container.remove();
      return { pageBoxes, editableCount: editable.length, editedText };
    }, Array.from(bytes));

    expect(out.pageBoxes).toBeGreaterThan(0); // real page boxes rendered
    expect(out.editableCount).toBeGreaterThan(0); // blocks inside pages stay editable
    expect(out.editedText).toContain('PAGINATEDEDIT77'); // incremental edit inside a page
  });

  // M1: editing a block preserves inline formatting (bold/italic/link) instead of
  // flattening to plain text. The projector emits bold=** italic=* links=[..](..),
  // so a save/reopen round-trip should show those markers.
  test('M1: editing preserves inline formatting (bold/italic/link)', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const target = (Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[]).find((e) => norm(e.textContent || '').length > 5);

      if (target) {
        // Simulate a user edit that includes formatting. <b>/<i> carry UA bold/italic
        // computed styles; <a> carries an href — exactly what the serializer reads.
        target.focus();
        target.innerHTML =
          'Plain <b>BOLDWORD</b> and <i>ITALWORD</i> and ' +
          '<a href="https://example.com/x">LINKWORD</a> end.';
        target.dispatchEvent(new Event('blur'));
      }

      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const markdown = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      editor.close();
      container.remove();
      return { markdown };
    }, Array.from(bytes));

    // Formatting survived the edit + save (projector convention: ** bold, * italic).
    expect(out.markdown).toContain('**BOLDWORD**');
    expect(out.markdown).toContain('*ITALWORD*');
    expect(out.markdown).toContain('[LINKWORD](https://example.com/x)');
  });

  // M2: structural editing — Enter splits a paragraph, Backspace at start merges.
  // Split a block (+1), then merge the halves back (-1, text restored), then save.
  test('M2: split and merge blocks via keyboard', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const alnum = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      const editableEls = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];
      const count = () => editableEls().length;

      const firstText = (el: HTMLElement): Text | null =>
        document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode() as Text | null;
      const setCaret = (el: HTMLElement, offset: number) => {
        const tn = firstText(el);
        const sel = window.getSelection()!;
        const r = document.createRange();
        if (tn) r.setStart(tn, Math.min(offset, tn.length));
        else r.selectNodeContents(el);
        r.collapse(true);
        sel.removeAllRanges();
        sel.addRange(r);
      };
      const press = (el: HTMLElement, key: string) =>
        el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));

      // Pick a block with enough text to split in the middle.
      const target = editableEls().find((e) => norm(e.textContent || '').length > 12)!;
      const originalText = norm(target.textContent || '');
      const before = count();

      // SPLIT at offset 6.
      setCaret(target, 6);
      press(target, 'Enter');
      const afterSplit = count();

      // MERGE the second half back into the first (Backspace at its start).
      const halves = editableEls();
      const firstHalf = halves.find((e) => alnum(originalText).startsWith(alnum(e.textContent || '')) && alnum(e.textContent || '').length > 0) || halves[0];
      const secondHalf = firstHalf.nextElementSibling as HTMLElement;
      setCaret(secondHalf, 0);
      press(secondHalf, 'Backspace');
      const afterMerge = count();
      const mergedText = norm((firstHalf.isConnected ? firstHalf : editableEls()[0]).textContent || '');

      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      editor.close();
      container.remove();
      return {
        before, afterSplit, afterMerge,
        originalAlnum: alnum(originalText),
        mergedAlnum: alnum(mergedText),
        mdLen: md.length,
      };
    }, Array.from(bytes));

    expect(out.afterSplit).toBe(out.before + 1); // Enter split one block into two
    expect(out.afterMerge).toBe(out.before); // Backspace merged them back
    expect(out.mergedAlnum).toBe(out.originalAlnum); // text restored exactly
    expect(out.mdLen).toBeGreaterThan(0); // valid doc after structural edits
  });

  // M5: ribbon commands — apply bold to a selection (ApplyFormat), set a paragraph
  // style (Heading1), and undo. Routes through DocxSession, lossless on save.
  test('M5: format selection, set style, undo', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const editableEls = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];
      const h1count = () => container.querySelectorAll('h1[data-anchor]').length;
      const selectChars = (el: HTMLElement, start: number, len: number) => {
        const tn = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode() as Text | null;
        const sel = window.getSelection()!;
        const r = document.createRange();
        if (tn) { r.setStart(tn, Math.min(start, tn.length)); r.setEnd(tn, Math.min(start + len, tn.length)); }
        else r.selectNodeContents(el);
        sel.removeAllRanges();
        sel.addRange(r);
      };

      // BOLD the first word of a paragraph via the format command.
      const b1 = editableEls().find((e) => norm(e.textContent || '').length > 10)!;
      const word = norm(b1.textContent || '').split(' ')[0];
      b1.focus();
      selectChars(b1, 0, word.length);
      editor.format('bold');

      const saved1: Uint8Array = editor.save();
      const r1 = D.DocxSessionBridge.OpenSession(saved1, '');
      const md1 = JSON.parse(D.DocxSessionBridge.Project(r1)).markdown as string;
      D.DocxSessionBridge.CloseSession(r1);

      // SET STYLE Heading1 on a paragraph, then UNDO it.
      const h1Before = h1count();
      const b2 = editableEls().find((e) => norm(e.textContent || '').length > 10)!;
      b2.focus();
      editor.setParagraphStyle('Heading1');
      const h1AfterStyle = h1count();
      editor.undo();
      const h1AfterUndo = h1count();

      editor.close();
      container.remove();
      return {
        boldApplied: md1.includes('**' + word + '**'),
        h1Before, h1AfterStyle, h1AfterUndo,
      };
    }, Array.from(bytes));

    expect(out.boldApplied).toBe(true); // bold survived to the saved document
    expect(out.h1AfterStyle).toBe(out.h1Before + 1); // SetParagraphStyle made a heading
    expect(out.h1AfterUndo).toBe(out.h1Before); // undo reverted it
  });

  // M5b: extended controls — alignment, indent, superscript (paragraph + run props
  // via SetParagraphFormat / ApplyFormat vertAlign). Asserts the visible effect on the
  // re-rendered block (inline styles).
  test('M5b: alignment, indent, and superscript apply and render', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const editable = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];

      // ALIGNMENT: center a paragraph.
      const a = editable().find((e) => norm(e.textContent || '').length > 10)!;
      const aUnid = a.getAttribute('data-anchor');
      a.focus();
      editor.setAlignment('center');
      const centered = container.querySelector(`[data-anchor="${aUnid}"]`) as HTMLElement;
      const textAlign = centered ? getComputedStyle(centered).textAlign : '';

      // INDENT: increase a paragraph's left indent.
      const b = editable().find((e) => norm(e.textContent || '').length > 10)!;
      const bUnid = b.getAttribute('data-anchor');
      b.focus();
      editor.indent(720);
      const indented = container.querySelector(`[data-anchor="${bUnid}"]`) as HTMLElement;
      const marginLeft = indented ? parseFloat(getComputedStyle(indented).marginLeft) : 0;

      // SUPERSCRIPT: a word in a paragraph.
      const c = editable().find((e) => norm(e.textContent || '').length > 10)!;
      const word = norm(c.textContent || '').split(' ')[0];
      c.focus();
      const tn = document.createTreeWalker(c, NodeFilter.SHOW_TEXT).nextNode() as Text;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(tn, 0); r.setEnd(tn, word.length);
      sel.removeAllRanges(); sel.addRange(r);
      editor.format('superscript');
      const after = editable().find((e) => norm(e.textContent || '').startsWith(word));
      let superFound = false;
      if (after) {
        for (const el of after.querySelectorAll('*')) {
          if (getComputedStyle(el).verticalAlign === 'super' || el.tagName === 'SUP') { superFound = true; break; }
        }
      }

      editor.close();
      container.remove();
      return { textAlign, marginLeft, superFound };
    }, Array.from(bytes));

    expect(out.textAlign).toBe('center'); // alignment applied + rendered
    expect(out.marginLeft).toBeGreaterThan(0); // indent applied + rendered
    expect(out.superFound).toBe(true); // superscript applied + rendered
  });
});
