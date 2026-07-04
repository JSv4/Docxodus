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

// ===========================================================================
// DocxEditor — "write a business letter" smoke test.
//
// Produced by test-driving the DocxEditor demo (`npm run demo`) as a real user:
// open a .docx, type the letter, press Enter between paragraphs, and format with
// the ribbon (bold / align / bullets). All four tests pass; the last two are
// regression guards for two authoring bugs the smoke test found and that are now
// fixed in editor.ts:
//
//   • BUG 1 (headline) — pressing Enter to split the paragraph you are *currently
//     editing* threw an uncaught NotFoundError and dropped the split. splitAtCaret
//     called `el.replaceWith()` on the FOCUSED block; removing the focused node
//     fires a synchronous `blur`, which re-entered commitBlock, and the interleaved
//     double-replaceWith threw. The repo's existing M2 split test never hit this
//     because it sets a caret range WITHOUT focusing the block — it never exercised
//     the state a typing user is always in. Fix: a `replacing` re-entrancy guard
//     (commitBlock no-ops while a structural replace is in flight) + isConnected
//     checks around every replaceWith (commitBlock/splitAtCaret/mergeWithPrevious/
//     swapBlock).
//   • BUG 2 — Enter after typing the first line of a blank document did nothing
//     (the blank paragraph renders with a placeholder space; typing lands after it,
//     so the DOM caret offset exceeded the trimmed run-text length, SplitParagraph
//     returned OffsetOutOfRange, and splitAtCaret silently swallowed it). Fix:
//     trimmedSplitOffset maps the DOM caret offset into the trimmed run-text the
//     session actually commits.
// ===========================================================================

test.describe('DocxEditor — business letter authoring (smoke test)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // ---- Capability: edit text + bold, save losslessly --------------------
  test('edits a paragraph and bolds it; lossless save round-trips both', async ({ page }) => {
    const bytes = readTestFile('Blank-wml.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const editableEls = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];

      // Set the letterhead line. (Edit a NON-focused block + blur — the editor's
      // reliable commit path; this is how editor.spec.ts's M1 test edits a block.)
      const block = editableEls()[0];
      block.textContent = 'Docxodus, Inc.';
      block.dispatchEvent(new Event('blur'));

      // Bold the whole line via the ribbon's format command.
      const fresh = editableEls()[0];
      fresh.focus();
      const tn = document.createTreeWalker(fresh, NodeFilter.SHOW_TEXT).nextNode() as Text;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(tn, 0);
      r.setEnd(tn, tn.length);
      sel.removeAllRanges();
      sel.addRange(r);
      editor.format('bold');

      const boldRendered = !!editableEls()[0].querySelector('[style*="font-weight: bold"]');

      // Save → reopen → project: text + bold must survive losslessly.
      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      editor.close();
      container.remove();
      return {
        boldRendered,
        textSaved: md.includes('Docxodus, Inc.'),
        boldSaved: md.includes('**Docxodus, Inc.**'),
      };
    }, Array.from(bytes));

    expect(out.boldRendered).toBe(true); // bold rendered in the live editor
    expect(out.textSaved).toBe(true); // edited text persisted to the saved document
    expect(out.boldSaved).toBe(true); // bold persisted to the saved document
  });

  // ---- Capability: right-align + bullet list, persisted through save -----
  test('right-align and bullet formatting persist through save and reopen', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();

      const c1 = document.createElement('div');
      document.body.appendChild(c1);
      const editor = D.DocxEditor.open(c1, bin, D, {});
      const editable = (root: HTMLElement) =>
        Array.from(root.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];

      // Right-align one paragraph (the "date" line in a letter).
      const dateEl = editable(c1).find((e) => norm(e.textContent || '').length > 12)!;
      const dateText = norm(dateEl.textContent || '').slice(0, 30);
      dateEl.focus();
      editor.setAlignment('right');

      // Promote a different paragraph to a bullet (a "deliverables" item).
      const bulletEl = editable(c1)
        .filter((e) => norm(e.textContent || '') !== norm(dateEl.textContent || ''))
        .find((e) => norm(e.textContent || '').length > 12)!;
      const bulletText = norm(bulletEl.textContent || '').slice(0, 30);
      bulletEl.focus();
      editor.toggleList('bullet');

      // Save, reopen in a fresh editor, and confirm both survived the round-trip.
      const saved: Uint8Array = editor.save();
      editor.close();
      c1.remove();

      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const editor2 = D.DocxEditor.open(c2, saved, D, {});
      const dateAfter = editable(c2).find((e) => norm(e.textContent || '').startsWith(dateText.slice(0, 20)));
      const bulletAfter = editable(c2).find((e) => norm(e.textContent || '').includes(bulletText.slice(0, 20)));
      const dateAlignPersisted = dateAfter ? getComputedStyle(dateAfter).textAlign : '';
      const markerEl = bulletAfter ? bulletAfter.querySelector('[data-list-marker]') : null;
      const bulletPersisted = !!markerEl;
      const bulletIndented = bulletAfter ? parseFloat(getComputedStyle(bulletAfter).marginLeft) > 0 : false;
      const bulletGlyph = markerEl ? (markerEl.textContent || '') : '';

      editor2.close();
      c2.remove();
      return { dateAlignPersisted, bulletPersisted, bulletIndented, bulletGlyph };
    }, Array.from(bytes));

    expect(out.dateAlignPersisted).toBe('right'); // alignment survived save + reopen
    expect(out.bulletPersisted).toBe(true); // list marker survived save + reopen
    expect(out.bulletIndented).toBe(true); // list indent survived save + reopen
    expect(out.bulletGlyph).toContain('•'); // marker renders as Unicode "•" (mapped from Symbol U+F0B7)
    expect(out.bulletGlyph).not.toContain(String.fromCharCode(0xf0b7)); // not the blank Symbol private-use glyph
  });

  // ---- Capability (BUG 3 fix): apply a built-in style the doc hadn't defined ----
  // A blank document defines Heading1 but not Heading2/Heading3/Title, so the editor's style
  // dropdown used to silently no-op for them (SetParagraphStyle returned UnknownStyle). The engine
  // now find-or-creates the well-known built-ins, so applying Heading 2 promotes the paragraph to
  // an <h2> and the synthesized style persists through save + reopen.
  test('applies a built-in heading style the document had not defined', async ({ page }) => {
    const bytes = readTestFile('Blank-wml.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const editable = (root: HTMLElement) =>
        Array.from(root.querySelectorAll('[data-anchor][contenteditable="true"]')) as HTMLElement[];

      const c1 = document.createElement('div');
      document.body.appendChild(c1);
      const editor = D.DocxEditor.open(c1, bin, D, {});

      // Give the first paragraph text, then apply Heading 2 (absent from a blank document).
      const block = editable(c1)[0];
      block.textContent = 'Quarterly Report';
      block.dispatchEvent(new Event('blur'));
      const fresh = editable(c1)[0];
      fresh.focus();
      editor.setParagraphStyle('Heading2');
      const h2InEditor = c1.querySelectorAll('h2[data-anchor]').length;

      // Persist through save + reopen in a fresh editor.
      const saved: Uint8Array = editor.save();
      editor.close();
      c1.remove();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const editor2 = D.DocxEditor.open(c2, saved, D, {});
      const h2AfterReopen = c2.querySelectorAll('h2[data-anchor]').length;
      const headingText = (c2.querySelector('h2[data-anchor]') as HTMLElement | null)?.textContent ?? '';
      editor2.close();
      c2.remove();
      return { h2InEditor, h2AfterReopen, headingText };
    }, Array.from(bytes));

    expect(out.h2InEditor).toBe(1); // Heading 2 applied + rendered as <h2>
    expect(out.h2AfterReopen).toBe(1); // synthesized style persisted to the saved document
    expect(out.headingText).toContain('Quarterly Report');
  });

  // ---- Regression (BUG 1): Enter while editing a focused paragraph -------
  // A user types a line and presses Enter to start the next paragraph while the
  // block is focused. This used to throw an uncaught NotFoundError (focused-node
  // removal → re-entrant commitBlock → double replaceWith) and drop the split.
  test('Enter splits the paragraph being edited without throwing', async ({ page }) => {
    const pageErrors: string[] = [];
    page.on('pageerror', (e) => pageErrors.push(e.message));
    const bytes = readTestFile('Blank-wml.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const editableEls = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];

      const block = editableEls()[0];
      block.textContent = 'First line of the letter'; // clean text, no leading space
      block.focus(); // the user is editing this block — it has focus
      const tn = document.createTreeWalker(block, NodeFilter.SHOW_TEXT).nextNode() as Text;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(tn, tn.length);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);

      const before = editableEls().length;
      block.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
      const after = editableEls().length;

      editor.close();
      container.remove();
      return { before, after };
    }, Array.from(bytes));

    expect(out.after, 'Enter should create a second paragraph').toBe(out.before + 1);
    expect(pageErrors, pageErrors.join('\n')).toHaveLength(0);
  });

  // ---- Regression (BUG 2): Enter on the first line of a blank document ---
  // The blank paragraph renders with a placeholder space; typing lands after it,
  // so the DOM caret offset overshot the trimmed run-text and SplitParagraph
  // returned OffsetOutOfRange — Enter used to be a silent no-op on the first line.
  test('Enter after typing the first line of a blank document starts a new paragraph', async ({ page }) => {
    const bytes = readTestFile('Blank-wml.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const editableEls = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];

      // Typing into the placeholder paragraph leaves the rendered leading space.
      const block = editableEls()[0];
      block.textContent = ' First Paragraph';
      block.focus();
      const tn = document.createTreeWalker(block, NodeFilter.SHOW_TEXT).nextNode() as Text;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(tn, tn.length);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);

      const before = editableEls().length;
      block.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
      const after = editableEls().length;

      editor.close();
      container.remove();
      return { before, after };
    }, Array.from(bytes));

    expect(out.after, 'Enter should create a second paragraph').toBe(out.before + 1);
  });

  // ---- Regression (split-propagation cluster): Enter after a styled paragraph ----
  // Drafting from scratch, a user sets the letterhead/section to a built-in style and
  // presses Enter to start the body. `DocxSession.SplitParagraph` used to clone the whole
  // w:pPr into the new paragraph, so the new block inherited the heading's STYLE (every body
  // block became a Heading) and, once text was typed and the paragraph re-set to Normal, the
  // heading style's BOLD was baked into the run. The fix rebases an empty Enter-at-end split
  // of a non-list paragraph onto the style's w:next (Normal) with a clean pPr. This guards the
  // full editor flow (the C# unit tests DS046/DS047/DS049 guard the DocxSession layer).
  test('Enter after a heading starts a Normal, non-bold body paragraph (next-style)', async ({ page }) => {
    const bytes = readTestFile('Blank-wml.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const editable = (root: HTMLElement) =>
        Array.from(root.querySelectorAll('[data-anchor][contenteditable="true"]')) as HTMLElement[];

      const c = document.createElement('div');
      document.body.appendChild(c);
      const editor = D.DocxEditor.open(c, bin, D, {});

      // Block 0 → Heading 2 (a bold built-in style).
      const h = editable(c)[0];
      h.textContent = 'Principal Economic Terms';
      h.dispatchEvent(new Event('blur'));
      editable(c)[0].focus();
      editor.setParagraphStyle('Heading2');

      // Caret at the END of the heading, then Enter — the user starting the next line.
      const hBlock = c.querySelector('h2[data-anchor]') as HTMLElement;
      hBlock.focus();
      const tn = document.createTreeWalker(hBlock, NodeFilter.SHOW_TEXT).nextNode() as Text;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(tn, tn.length);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      hBlock.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));

      // Type body text into the freshly-created paragraph and commit it.
      const created = editable(c);
      const newBlock = created[created.length - 1];
      newBlock.focus();
      newBlock.textContent = 'The Investor shall purchase Series C Preferred Stock.';
      newBlock.dispatchEvent(new Event('blur'));

      // Re-locate after the commit re-render.
      const body = editable(c).find((b) => (b.textContent || '').includes('The Investor shall purchase'))!;
      const span = body.querySelector('span');
      editor.close();
      c.remove();
      return {
        headingStillH2: !!hBlock && hBlock.tagName === 'H2',
        newBlockTag: body.tagName, // A: should be P (next style Normal), not another H2
        newBlockWeight: span ? getComputedStyle(span).fontWeight : '400', // B: not bold
      };
    }, Array.from(bytes));

    expect(out.headingStillH2, 'the heading itself stays a heading').toBe(true);
    expect(out.newBlockTag, 'new paragraph after a heading is Normal (next style), not another heading').toBe('P');
    expect(out.newBlockWeight, 'typed body text is not bold (no baked-in heading bold)').not.toBe('700');
  });
});
