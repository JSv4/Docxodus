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

  // M1 (updated for full-fidelity editing): run formatting applied via the ribbon survives a later
  // TEXT edit to the same block. The old markdown-re-serialization path is gone — formatting is now
  // preserved structurally (the span-diff commit leaves untouched runs intact). Bold a word, then
  // append text elsewhere, and confirm the bold survived (and the doc still saves).
  test('M1: run formatting survives a later text edit', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
      const isBold = (el: HTMLElement) => {
        const fw = getComputedStyle(el).fontWeight; const n = parseInt(fw, 10);
        return fw === 'bold' || fw === 'bolder' || (!Number.isNaN(n) && n >= 600);
      };
      const firstTextNode = (el: Node): Text | null => {
        if (el.nodeType === 3) return el as Text;
        for (const c of Array.from(el.childNodes)) { const r = firstTextNode(c); if (r) return r; }
        return null;
      };

      const target = (Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[]).find((e) => norm(e.textContent || '').length > 12)!;
      const anchor = target.getAttribute('data-anchor')!;
      const head = norm(target.textContent || '').slice(0, 3);

      // Bold the first 5 content chars via the ribbon path (ApplyFormat → runs).
      const tn = firstTextNode(target)!;
      const raw = tn.textContent || '';
      const lead = raw.length - raw.replace(/^[‎‏]+/, '').length;
      target.focus();
      let sel = window.getSelection()!;
      let r = document.createRange();
      r.setStart(tn, lead); r.setEnd(tn, lead + 5);
      sel.removeAllRanges(); sel.addRange(r);
      editor.format('bold');

      // Append text at the end (a TEXT edit) and commit on blur.
      const blk = container.querySelector(`[data-anchor="${anchor}"]`) as HTMLElement;
      blk.focus();
      const r2 = document.createRange();
      r2.selectNodeContents(blk); r2.collapse(false);
      sel = window.getSelection()!; sel.removeAllRanges(); sel.addRange(r2);
      document.execCommand('insertText', false, ' TAILM1');
      blk.dispatchEvent(new Event('blur'));

      const after = container.querySelector(`[data-anchor="${anchor}"]`) as HTMLElement;
      const boldPreserved = (Array.from(after.querySelectorAll('span')) as HTMLElement[]).some(
        (s) => isBold(s) && norm(s.textContent || '').includes(head),
      );
      const savedLen = (editor.save() as Uint8Array).length;

      editor.close();
      container.remove();
      return { boldPreserved, hasTail: norm(after.textContent || '').includes('TAILM1'), savedLen };
    }, Array.from(bytes));

    expect(out.boldPreserved).toBe(true); // bold on the untouched word survived the text edit
    expect(out.hasTail).toBe(true);
    expect(out.savedLen).toBeGreaterThan(0);
  });

  // Text typed at the end of an underlined run inherits that run's formatting (ReplaceTextAtSpan
  // drops new text into the boundary run with its rPr intact), matching Word/contenteditable.
  test('text typed adjacent to a formatted run inherits its formatting', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
      const firstTextNode = (el: Node): Text | null => {
        if (el.nodeType === 3) return el as Text;
        for (const c of Array.from(el.childNodes)) { const r = firstTextNode(c); if (r) return r; }
        return null;
      };

      const target = (Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[]).find((e) => norm(e.textContent || '').length > 12)!;
      const anchor = target.getAttribute('data-anchor')!;

      // Underline the WHOLE paragraph, then type at the very end; the new text should be underlined.
      target.focus();
      const tn = firstTextNode(target)!;
      const raw = tn.textContent || '';
      const lead = raw.length - raw.replace(/^[‎‏]+/, '').length;
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.setStart(tn, lead); r.setEnd(target, target.childNodes.length);
      sel.removeAllRanges(); sel.addRange(r);
      editor.format('underline');

      const blk = container.querySelector(`[data-anchor="${anchor}"]`) as HTMLElement;
      blk.focus();
      const r2 = document.createRange();
      r2.selectNodeContents(blk); r2.collapse(false);
      sel.removeAllRanges(); sel.addRange(r2);
      document.execCommand('insertText', false, 'QQQ');
      blk.dispatchEvent(new Event('blur'));

      const after = container.querySelector(`[data-anchor="${anchor}"]`) as HTMLElement;
      const qSpan = (Array.from(after.querySelectorAll('span')) as HTMLElement[]).find(
        (s) => norm(s.textContent || '').includes('QQQ'),
      );
      const qUnderlined = !!qSpan && getComputedStyle(qSpan).textDecorationLine.includes('underline');

      editor.close();
      container.remove();
      return { hasQQQ: norm(after.textContent || '').includes('QQQ'), qUnderlined };
    }, Array.from(bytes));

    expect(out.hasQQQ).toBe(true);
    expect(out.qUnderlined).toBe(true); // typed text inherited the adjacent underlined run
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

  // Mlists: bullets / numbered lists — promote a plain paragraph to a list item
  // (synthesizing a numbering definition), confirm membership, remove it; plus the
  // editor toggleList wiring re-renders the block.
  test('Mlists: ApplyListFormat promotes a paragraph + editor toggle', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();

      // --- through the raw bridge (full WASM stack) ---
      const h: number = D.DocxSessionBridge.OpenSession(bin, '');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      const pAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:')) as string;
      const r = JSON.parse(D.DocxSessionBridge.ApplyListFormat(h, pAnchor, 'bullet'));
      const liId = r.modified[0].id as string;
      const lm = JSON.parse(D.DocxSessionBridge.GetListMembership(h, liId));
      const bridgeBullet = !!lm && typeof lm.format === 'string' && lm.format.toLowerCase().startsWith('bullet');
      const r2 = JSON.parse(D.DocxSessionBridge.ApplyListFormat(h, liId, 'none'));
      const lm2raw = D.DocxSessionBridge.GetListMembership(h, r2.modified[0].id);
      const removed = lm2raw === 'null' || JSON.parse(lm2raw) === null;
      D.DocxSessionBridge.CloseSession(h);

      // --- editor toggleList wiring + visible marker ---
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});
      const all = Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];
      // Pick a UNIQUE-text paragraph (HC031 repeats some), so we can re-find the exact block.
      const tgt = all.find((e) => {
        const t = norm(e.textContent || '').slice(0, 40);
        return t.length > 20 && all.filter((x) => norm(x.textContent || '').startsWith(t)).length === 1;
      })!;
      const key = norm(tgt.textContent || '').slice(0, 40);
      tgt.focus();
      editor.toggleList('bullet');
      const after = (Array.from(container.querySelectorAll('[data-anchor]')) as HTMLElement[])
        .find((e) => norm(e.textContent || '').includes(key));
      const marginLeft = after ? parseFloat(getComputedStyle(after).marginLeft) : 0;
      const hasMarker = after ? after.outerHTML.includes('\u2022') : false; // Unicode bullet glyph (mapped from Symbol U+F0B7)
      editor.close();
      container.remove();

      return { bridgeBullet, removed, liKind: r.modified[0].kind, marginLeft, hasMarker };
    }, Array.from(bytes));

    expect(out.liKind).toBe('li'); // plain paragraph promoted to a list item
    expect(out.bridgeBullet).toBe(true); // it's a bullet list
    expect(out.removed).toBe(true); // toggling to "none" removed membership
    expect(out.marginLeft).toBeGreaterThan(0); // list indent rendered
    expect(out.hasMarker).toBe(true); // bullet marker rendered as Unicode \u2022 in the editor
  });

  // Mlists2: numbered lists CONTINUE (1., 2.), and Enter on a list item adds a
  // continuing item — the two issues found in manual testing.
  test('Mlists2: numbered continuation + Enter adds a continuing item', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();

      // Numbering continuation is checked precisely via the bridge (anchor-addressed),
      // since HC031 repeats paragraphs and find-by-text is unreliable.
      const h: number = D.DocxSessionBridge.OpenSession(bin, '{"persistAnchorIds":true}');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      const pAnchors = Object.keys(proj.anchorIndex).filter((k) => k.startsWith('p:'));
      const liA = JSON.parse(D.DocxSessionBridge.ApplyListFormat(h, pAnchors[0], 'decimal')).modified[0].unid;
      const liB = JSON.parse(D.DocxSessionBridge.ApplyListFormat(h, pAnchors[1], 'decimal')).modified[0].unid;
      const full = D.DocumentConverter.ConvertDocxToHtmlComplete(
        D.DocxSessionBridge.Save(h), 'Document', 'docx-', false, '', -1, 'comment-', 0, 1.0, 'page-',
        false, 0, 'annot-', false, false, false, true, true, false, null, true);
      D.DocxSessionBridge.CloseSession(h);
      const doc = new DOMParser().parseFromString(full, 'text/html');
      const numFor = (unid: string) => {
        const el = doc.querySelector(`[data-anchor="${unid}"]`);
        const m = el ? norm(el.textContent || '').match(/^(\d+)\./) : null;
        return m ? parseInt(m[1], 10) : -1;
      };
      const nums = [numFor(liA), numFor(liB)].sort((a, b) => a - b);
      const numA = nums[0];
      const numB = nums[1];

      // Enter at end of a numbered item adds a continuing item (editor path).
      const edits: any[] = [];
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, { onEdit: (i: any) => edits.push(i) });
      const list = () => Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];
      const firstP = list().find((e) => norm(e.textContent || '').length > 15)!;
      firstP.focus();
      editor.toggleList('decimal'); // make it a numbered item
      const item = list().find((e) => /^\d+\./.test(norm(e.textContent || '')))!;
      const countBefore = list().length;
      const editsBefore = edits.length;
      item.focus();
      const tn = (() => {
        const w = document.createTreeWalker(item, NodeFilter.SHOW_TEXT, {
          acceptNode: (n) => (n.parentElement && (n.parentElement as HTMLElement).closest('[data-list-marker]')) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
        } as any);
        let n: Node | null, last: Text | null = null;
        while ((n = w.nextNode())) last = n as Text;
        return last;
      })();
      const sel = window.getSelection()!;
      const r = document.createRange();
      if (tn) { r.setStart(tn, tn.length); } else { r.selectNodeContents(item); }
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      item.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
      const countAfter = list().length;

      editor.close();
      container.remove();
      // Bridge probe: does SplitParagraph work on a list item at all?
      const hp: number = D.DocxSessionBridge.OpenSession(bin, '{"persistAnchorIds":true}');
      const pp = Object.keys(JSON.parse(D.DocxSessionBridge.Project(hp)).anchorIndex).find((k) => k.startsWith('p:')) as string;
      const liP = JSON.parse(D.DocxSessionBridge.ApplyListFormat(hp, pp, 'decimal')).modified[0].id;
      const splitMid = JSON.parse(D.DocxSessionBridge.SplitParagraph(hp, liP, 3));
      D.DocxSessionBridge.CloseSession(hp);

      const markerEls = Array.from(item.querySelectorAll('[data-list-marker]'));
      const markerText = markerEls.map((m) => m.textContent || '').join('');
      return {
        numA, numB, countBefore, countAfter,
        editsFired: edits.length - editsBefore,
        itemTag: item.tagName,
        markerSpanCount: markerEls.length,
        markerText: markerText.slice(0, 10),
        fullTextLen: (item.textContent || '').length,
        markerTextLen: markerText.length,
        bridgeSplitOk: splitMid.success === true,
        bridgeSplitErr: splitMid.error ? splitMid.error.code : null,
      };
    }, Array.from(bytes));

    // Diagnostics surfaced if it fails.
    expect(out.numA).toBeGreaterThan(0); // first numbered item is wired + numbered
    expect(out.numB).toBeGreaterThan(out.numA); // numbering CONTINUES (B > A)
    // Diagnostics print in the diff if this fails.
    expect(out).toEqual({ ...out, countAfter: out.countBefore + 1, editsFired: 1 });
  });

  // Mlists3: clicking from one numbered item to another must KEEP focus on the clicked item, and
  // typing into each (including freshly-created empty ones) must work. Regression: committing a
  // list item on blur re-rendered its DOM node mid-blur, which cancelled the browser's in-flight
  // focus transfer to the clicked bullet (focus fell to <body>, so typing went nowhere). Driven
  // with REAL mouse clicks + REAL keyboard — the only faithful repro of the focus-transfer bug.
  test('Mlists3: clicking between numbered items keeps focus; typing into each works', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    // Build a 3-item numbered list (item 1 has text; items 2 & 3 are empty) in a real container.
    const anchors = await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const container = document.createElement('div');
      container.id = 'mtest';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});
      (window as any).__m = { editor, container };
      const list = () => Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];
      const caretEnd = (el: HTMLElement) => {
        el.focus();
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
          acceptNode: (n: Node) => (n.parentElement && (n.parentElement as HTMLElement).closest('[data-list-marker]')) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
        } as any);
        let last: Node | null = null, n: Node | null;
        while ((n = w.nextNode())) last = n;
        const r = document.createRange();
        if (last) { r.setStart(last, (last as Text).length); r.collapse(true); } else { r.selectNodeContents(el); r.collapse(false); }
        const s = getSelection()!; s.removeAllRanges(); s.addRange(r);
      };
      const first = list().find((e) => norm(e.textContent || '').length > 20)!;
      caretEnd(first);
      editor.toggleList('decimal');
      const fire = (el: HTMLElement) => { caretEnd(el); el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); };
      const it = list().find((e) => /^\d+\./.test(norm(e.textContent || '')))!;
      fire(it); fire(document.activeElement as HTMLElement); // create items 2 and 3
      (document.activeElement as HTMLElement)?.blur?.();
      return list().filter((e) => /^\d+\./.test(norm(e.textContent || ''))).slice(0, 3).map((e) => e.getAttribute('data-anchor')!);
    }, Array.from(bytes));

    expect(anchors.length).toBe(3);

    // Real click item 1, append text.
    await page.locator(`#mtest [data-anchor="${anchors[0]}"]`).click();
    await page.keyboard.press('End');
    await page.keyboard.type(' X');

    // Click item 2 (commits item 1 on blur). Focus MUST land on item 2 — the bug dropped it to <body>.
    await page.locator(`#mtest [data-anchor="${anchors[1]}"]`).click();
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-anchor'))).toBe(anchors[1]);
    await page.keyboard.type('BB');

    // Click item 3 (commits item 2). Same focus requirement, and typing into the empty item works.
    await page.locator(`#mtest [data-anchor="${anchors[2]}"]`).click();
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-anchor'))).toBe(anchors[2]);
    await page.keyboard.type('CC');

    // Commit the last item; assert numbering stayed 1/2/3 and every item kept its typed text.
    const result = await page.evaluate((a: string[]) => {
      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const container = (window as any).__m.container as HTMLElement;
      (document.activeElement as HTMLElement)?.blur?.();
      const items = a.map((anc) => {
        const el = container.querySelector(`[data-anchor="${anc}"]`);
        const t = norm(el?.textContent || '');
        return { num: t.match(/^(\d+)\./)?.[1] ?? null, text: t };
      });
      const savedLen = (window as any).__m.editor.save()?.length ?? 0;
      return { items, savedLen };
    }, anchors);

    expect(result.items.map((i) => i.num)).toEqual(['1', '2', '3']); // numbering intact
    expect(result.items[0].text.endsWith('X')).toBe(true); // item 1 kept appended text
    expect(result.items[1].text).toBe('2. BB'); // typed into empty item 2
    expect(result.items[2].text).toBe('3. CC'); // typed into empty item 3
    expect(result.savedLen).toBeGreaterThan(0); // saves losslessly
  });

  // Mlists4: Tab / Shift+Tab nests / un-nests a numbered item (changes list LEVEL via
  // SetListLevel), so numbering nests (1, 2, [nested 1], 3) instead of staying flat. Regression:
  // the editor's indent only changed the paragraph margin, so "nested lists" rendered flat.
  test('Mlists4: Tab nests a numbered item (list level) and Shift+Tab un-nests it', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});
      const list = () => Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]')) as HTMLElement[];
      const isLi = (e: HTMLElement) => !!e.querySelector(':scope > [data-list-marker]');
      const caretEnd = (el: HTMLElement) => {
        el.focus();
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
          acceptNode: (n: Node) => (n.parentElement && (n.parentElement as HTMLElement).closest('[data-list-marker]')) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
        } as any);
        let last: Node | null = null, n: Node | null;
        while ((n = w.nextNode())) last = n;
        const r = document.createRange();
        if (last) { r.setStart(last, (last as Text).length); r.collapse(true); } else { r.selectNodeContents(el); r.collapse(false); }
        const s = getSelection()!; s.removeAllRanges(); s.addRange(r);
      };
      // Build a 4-item numbered list.
      caretEnd(list().find((e) => norm(e.textContent || '').length > 20)!);
      editor.toggleList('decimal');
      for (let i = 0; i < 3; i++) { const cur = document.activeElement as HTMLElement; caretEnd(cur); cur.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); }

      const liItems = () => list().filter(isLi);
      const snapshot = () => liItems().slice(0, 4).map((e) => ({ num: norm(e.textContent || '').match(/^(\d+)\./)?.[1] ?? null, indentPx: parseFloat(getComputedStyle(e).marginLeft) }));

      const flat = snapshot();
      // Tab on item 3 to nest it one level deeper.
      const it3 = liItems()[2];
      it3.focus(); editor.activeBlock = it3;
      it3.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }));
      const nested = snapshot();
      // Shift+Tab on the (still 3rd) item to outdent back to level 0.
      const it3b = liItems()[2];
      it3b.focus(); editor.activeBlock = it3b;
      it3b.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true }));
      const outdented = snapshot();
      const savedLen = editor.save()?.length ?? 0;
      editor.close(); container.remove();
      return { flat, nested, outdented, savedLen };
    }, Array.from(bytes));

    // Flat list: 1,2,3,4 all at the same indent.
    expect(out.flat.map((i) => i.num)).toEqual(['1', '2', '3', '4']);
    // After nesting item 3: it is indented deeper than item 2 and its number restarts at the
    // sub-level ("1"), while item 4 continues the outer level ("3").
    expect(out.nested[2].indentPx).toBeGreaterThan(out.nested[1].indentPx);
    expect(out.nested[2].num).toBe('1');
    expect(out.nested[3].num).toBe('3');
    // After Shift+Tab: back to a flat 1,2,3,4 at one indent.
    expect(out.outdented.map((i) => i.num)).toEqual(['1', '2', '3', '4']);
    expect(out.outdented[2].indentPx).toBe(out.outdented[1].indentPx);
    expect(out.savedLen).toBeGreaterThan(0);
  });

  // Regression: the HTML converter wraps directional run text in bidi marks (U+200E/U+200F)
  // that are NOT in the session's run text (Google-Docs-exported paragraphs hit this on every
  // block). The editor's caret-offset math must exclude those marks; otherwise a caret at
  // end-of-line maps past the session's text length and SplitParagraph rejects the offset, so
  // Enter is silently dropped. Here we simulate the converter output by injecting a leading and
  // trailing RLM into a block's DOM text (and matching its committed text, so the synthetic edit
  // isn't committed back), then press Enter at the very end.
  test('Enter at end-of-line works when block text carries bidi marks', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/[\u200E\u200F]/g, '').replace(/\s+/g, ' ').trim();
      const target = Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ).find((e) => norm((e as HTMLElement).textContent || '').length > 10) as HTMLElement;

      // Inject a leading + trailing RLM into the block's first text node, exactly as the
      // converter would for a directional paragraph, and align the editor's committed-text
      // bookkeeping so the marks are treated as already-present (not a user edit).
      const tn = (function f(n: Node): Text | null {
        if (n.nodeType === 3) return n as Text;
        for (const c of Array.from(n.childNodes)) { const r = f(c); if (r) return r; }
        return null;
      })(target)!;
      tn.textContent = '\u200F' + tn.textContent + '\u200F';
      (target as any).dataset.committedText = target.textContent;

      const blockCountBefore = container.querySelectorAll('[data-anchor]').length;

      // Caret at the very end, then a real Enter keydown.
      target.focus();
      const sel = window.getSelection()!;
      const range = document.createRange();
      range.selectNodeContents(target);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
      target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));

      const blockCountAfter = container.querySelectorAll('[data-anchor]').length;
      // The new (empty) paragraph should follow the one still holding the original text.
      const labelsBlock = Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ).find((e) => norm((e as HTMLElement).textContent || '').length > 10) as HTMLElement;
      const next = labelsBlock ? (labelsBlock.nextElementSibling as HTMLElement | null) : null;
      const newIsEmpty = !!next && norm(next.textContent || '') === '';

      editor.close();
      container.remove();
      return { blockCountBefore, blockCountAfter, newIsEmpty };
    }, Array.from(bytes));

    // A new paragraph was created (the Enter was NOT silently dropped) and it is empty.
    expect(out.blockCountAfter).toBe(out.blockCountBefore + 1);
    expect(out.newIsEmpty).toBe(true);
  });

  // Regression: editing a block must preserve run formatting the markdown subset can't express
  // (here, underline). The old commit path re-serialized to markdown (bold/italic/link only), so
  // underline on an untouched run was dropped. The diff-based commit edits only the changed text
  // span via ReplaceTextAtSpan, leaving every other run's rPr intact.
  test('editing a block preserves underline on an untouched run', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
      const firstTextNode = (el: Node): Text | null => {
        if (el.nodeType === 3) return el as Text;
        for (const c of Array.from(el.childNodes)) { const r = firstTextNode(c); if (r) return r; }
        return null;
      };

      const target = (Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[]).find((e) => norm(e.textContent || '').length > 12)!;
      const anchor = target.getAttribute('data-anchor')!;
      const firstWord = norm(target.textContent || '').slice(0, 4);

      // Underline the first 4 content chars via the ribbon path (ApplyFormat).
      const tn = firstTextNode(target)!;
      const raw = tn.textContent || '';
      const lead = raw.length - raw.replace(/^[‎‏]+/, '').length;
      target.focus();
      let sel = window.getSelection()!;
      let r = document.createRange();
      r.setStart(tn, lead); r.setEnd(tn, lead + 4);
      sel.removeAllRanges(); sel.addRange(r);
      editor.format('underline');

      // Type " ZZZ" at the very end of the (re-rendered) block, preserving its spans.
      const blk = container.querySelector(`[data-anchor="${anchor}"]`) as HTMLElement;
      blk.focus();
      const r2 = document.createRange();
      r2.selectNodeContents(blk); r2.collapse(false);
      sel = window.getSelection()!; sel.removeAllRanges(); sel.addRange(r2);
      document.execCommand('insertText', false, ' ZZZ');
      blk.dispatchEvent(new Event('blur'));

      // After commit, inspect the re-rendered block (rendered from the live session).
      const after = container.querySelector(`[data-anchor="${anchor}"]`) as HTMLElement;
      const underlinedSpan = (Array.from(after.querySelectorAll('span')) as HTMLElement[]).find(
        (s) => getComputedStyle(s).textDecorationLine.includes('underline') && norm(s.textContent || '').includes(firstWord),
      );
      const savedLen = (editor.save() as Uint8Array).length;

      editor.close();
      container.remove();
      return {
        firstWord,
        hasZZZ: norm(after.textContent || '').includes('ZZZ'),
        underlinePreserved: !!underlinedSpan,
        savedLen,
      };
    }, Array.from(bytes));

    expect(out.hasZZZ).toBe(true);          // the text edit landed
    expect(out.underlinePreserved).toBe(true); // underline on the untouched first word survived
    expect(out.savedLen).toBeGreaterThan(0);
  });
});
