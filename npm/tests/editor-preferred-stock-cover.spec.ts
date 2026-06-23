import { test, expect, Page } from '@playwright/test';

/**
 * GUI-driven reproduction of a venture-financing "preferred stock purchase agreement" cover page,
 * built end-to-end through the demo's real ribbon controls (toolbar button clicks, the font-family
 * dropdown, the legal-numbering button, the horizontal-rule button) and real, native keyboard input
 * (page.keyboard.type / page.keyboard.press). It is the fuller companion to
 * editor-contract-form.spec.ts and additionally exercises the surfaces a faithful reproduction of
 * such a page actually needs but the earlier spec did not cover:
 *
 *   - JUSTIFIED body paragraphs (w:jc="both") — the disclaimer, the note, the recital and the (a)/(b)
 *     sub-clauses are all justified, the hallmark of a typeset legal page;
 *   - a WHOLE-DOCUMENT serif font (Times New Roman) applied via Select-All + the font dropdown, and
 *     asserted to reach prose paragraphs AND legal-numbered list items;
 *   - a recital carrying FOUR inline-bold defined terms ("Agreement", "Company", "Purchaser",
 *     "Purchasers") plus an underlined cross-reference ("Exhibit A") in a single paragraph;
 *   - a bold + italic "Preliminary Note" heading and an italic note body;
 *   - a CENTERED BOLD document title;
 *   - multi-level legal numbering 1. / 1.1 / (a) / (b) with hanging indents, the captions bold +
 *     underlined, and EACH numbered item asserted to keep its OWN distinct text.
 *
 * All text is invented placeholder prose — never any copyrighted contract language.
 *
 * The suite saves the produced .docx and asserts the resulting OOXML (w:b / w:i / w:u / w:jc /
 * w:numPr+w:ilvl / w:pBdr / w:rFonts) so the editor's coverage is verified end to end, proves a
 * lossless save -> reopen round-trip, and records the three engine-level omissions a real agreement
 * page needs but the editor cannot author: footnotes, footer text and page numbers.
 *
 * Authoring note: text is entered with NATIVE keyboard input (page.keyboard.type), exactly the path a
 * human types through. An earlier helper that injected text with document.execCommand('insertText')
 * mis-routed characters when a list item was split with Enter (the new item's caret had not yet been
 * restored), collapsing every numbered clause's text into the last item. Native typing follows the
 * editor's restored caret and builds each item correctly; that contrast is locked in by the
 * "keeps each numbered item distinct" regression test below.
 */

// Invented placeholder text (structurally similar to a financing cover page, no copyrighted language).
const T = {
  intro1:
    'This sample document is the work product of a demonstration drafting group and is intended to serve as a starting point only; it should be tailored to meet your specific requirements and should not be construed as legal advice for any particular facts or circumstances.',
  intro2:
    'For those who will compare this against earlier sample versions, newly added footnotes are identified as New, and footnotes that contain substantive revisions since the prior version are flagged as Revised.',
  noteHead: 'Preliminary Note',
  noteBody:
    'The Stock Purchase Agreement sets forth the basic terms of the purchase and sale of the preferred stock to the investors, such as the purchase price, the closing date and the conditions to closing. The principal items of negotiation are the representations and warranties that the company must make and the closing conditions for the transaction.',
  title: 'SERIES [___] PREFERRED STOCK PURCHASE AGREEMENT',
  recital:
    'THIS SERIES [___] PREFERRED STOCK PURCHASE AGREEMENT (this "Agreement") is made as of [____], 20[__], by and among [____________], a Delaware corporation (the "Company"), and the investors listed on Exhibit A attached to this Agreement (each a "Purchaser" and together the "Purchasers").',
  transition: 'The parties hereby agree as follows:',
  clause1: 'Purchase and Sale of Preferred Stock.',
  clause11: 'Sale and Issuance of Preferred Stock.',
  clauseA:
    'The Company shall have adopted and filed with the Secretary of State of Delaware, on or before the initial closing, an Amended and Restated Certificate of Incorporation in the form attached to this Agreement (the "Restated Certificate").',
  clauseB:
    'Subject to the terms and conditions of this Agreement, each Purchaser agrees to purchase, and the Company agrees to sell and issue to each Purchaser, the number of shares of preferred stock set opposite that Purchaser name on the attached schedule, at the applicable per-share purchase price.',
};

const SERIF = 'Times New Roman';
const SEED = '#editor p[data-anchor][data-editable="1"]';

async function boot(page: Page) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
}

/** Place a collapsed caret at the end of the body block whose text contains `textMatch`. */
async function caretAtEndOf(page: Page, textMatch: string) {
  await page.evaluate((tm) => {
    const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find(
      (p) => (p.textContent || '').includes(tm),
    ) as HTMLElement | undefined;
    if (!blk) throw new Error('block not found for caret: ' + tm);
    blk.focus();
    const r = document.createRange();
    r.selectNodeContents(blk);
    r.collapse(false);
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    document.dispatchEvent(new Event('selectionchange'));
  }, textMatch);
}

/** Select the entire contents of the body block whose text contains `textMatch`. */
async function selectWholeBlock(page: Page, textMatch: string) {
  await page.evaluate((tm) => {
    const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find(
      (p) => (p.textContent || '').includes(tm),
    ) as HTMLElement | undefined;
    if (!blk) throw new Error('block not found: ' + tm);
    blk.focus();
    const r = document.createRange();
    r.selectNodeContents(blk);
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    document.dispatchEvent(new Event('selectionchange'));
  }, textMatch);
}

/** Apply an inline ribbon command (data-cmd) to the whole block containing `textMatch`. */
async function ribbonFormatBlock(page: Page, textMatch: string, cmd: string) {
  await selectWholeBlock(page, textMatch);
  await page.click(`#ribbon button[data-cmd="${cmd}"]`);
}

/** Set paragraph alignment on the whole block containing `textMatch`. */
async function alignBlock(page: Page, textMatch: string, align: 'left' | 'center' | 'right' | 'justify') {
  await selectWholeBlock(page, textMatch);
  await page.click(`#ribbon button[data-align="${align}"]`);
}

/** Select the inner word of a quoted/cross-referenced phrase inside the recital, then click a command. */
async function emphasizeSpan(page: Page, quoted: string, inner: string, cmd: string) {
  await page.evaluate(
    ({ quoted, inner }) => {
      const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find((p) =>
        /THIS SERIES/.test(p.textContent || ''),
      ) as HTMLElement;
      const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
      let target: Text | null = null;
      let idx = -1;
      let n: Node | null;
      while ((n = walker.nextNode())) {
        const i = (n.textContent || '').indexOf(quoted);
        if (i >= 0) {
          target = n as Text;
          idx = i;
          break;
        }
      }
      if (!target) throw new Error('span not found: ' + quoted);
      const start = idx + quoted.indexOf(inner);
      const r = document.createRange();
      r.setStart(target, start);
      r.setEnd(target, start + inner.length);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
      document.dispatchEvent(new Event('selectionchange'));
    },
    { quoted, inner },
  );
  await page.click(`#ribbon button[data-cmd="${cmd}"]`);
}

/** Save the document and project it; return per-block OOXML facts keyed off the anchor index. */
async function inspect(page: Page) {
  return page.evaluate(() => {
    const D = (window as any).__demo;
    const B = D.exports.DocxSessionBridge;
    const ed = D.getEditor();
    const saved: Uint8Array = ed.save();
    const h = B.OpenSession(saved, '');
    const proj = JSON.parse(B.Project(h));
    const anchors: string[] = [
      ...(proj.markdown as string).matchAll(/\{#((?:p|li|h\d):body:[0-9a-f]+)\}/g),
    ].map((m) => m[1]);
    const blocks = anchors.map((a) => {
      const xml: string = B.RawGetXml(h, a);
      const text = (xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
        .map((t) => t.replace(/<[^>]+>/g, ''))
        .join('');
      const ilvlM = xml.match(/<w:ilvl w:val="(\d+)"/);
      const numIdM = xml.match(/<w:numId w:val="(\d+)"/);
      const jcM = xml.match(/<w:jc w:val="(\w+)"/);
      const bolds = (xml.match(/<w:b\b[^>]*\/?>/g) || []).filter((m) => !/w:bCs/.test(m)).length;
      const italics = (xml.match(/<w:i\b[^>]*\/?>/g) || []).filter((m) => !/w:iCs|w:ind/.test(m)).length;
      const fonts = [...xml.matchAll(/w:ascii="([^"]+)"/g)].map((m) => m[1]);
      return {
        kind: a.split(':')[0],
        text,
        jc: jcM ? jcM[1] : null,
        bold: /<w:b\b[^>]*\/?>/.test(xml) && !/^<w:bCs/.test(xml),
        boldRuns: bolds,
        italic: italics > 0,
        underline: /<w:u\b/.test(xml),
        ilvl: ilvlM ? Number(ilvlM[1]) : null,
        numId: numIdM ? Number(numIdM[1]) : null,
        hasBorder: /<w:pBdr/.test(xml),
        allTimes: fonts.length > 0 && fonts.every((f) => f === 'Times New Roman'),
      };
    });
    B.CloseSession(h);
    return { blockCount: blocks.length, blocks };
  });
}

/** Build the whole cover page through the ribbon (real button clicks + native keyboard input). */
async function buildCoverPage(page: Page) {
  await boot(page);
  await page.click(SEED); // real click → caret in the empty seed paragraph

  // 1) Type the prose blocks as PLAIN text, one block per logical paragraph (native typing follows
  //    the editor's restored caret across each Enter split).
  await page.keyboard.type(T.intro1);
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.intro2);
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.noteHead);
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.noteBody);
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.title);
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.recital);
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.transition);

  // 2) The four legal-numbered clauses, authored with the NATURAL flow a human uses: type the
  //    clause, press Tab to demote it to the right outline level, then Enter for the next clause.
  //    A list-level change re-renders the item from the session and RESTORES the caret to its prior
  //    end-of-line position (the legal-numbering caret fix), so demoting an item that already holds
  //    freshly-typed text is loss-free and the following Enter splits at the end — the clauses stay
  //    distinct. (The "Tab after typing keeps the caret at the line end" test below guards this.)
  await page.keyboard.press('Enter');
  await page.keyboard.type(T.clause1);
  await page.click('#legalNum'); // caption -> "1." (level 0)
  await page.keyboard.press('End');
  await page.keyboard.press('Enter'); // new item (level 0)
  await page.keyboard.type(T.clause11);
  await page.keyboard.press('Tab'); // demote AFTER typing -> "1.1" (level 1); caret stays at the end
  await page.keyboard.press('End');
  await page.keyboard.press('Enter'); // new item (level 1)
  await page.keyboard.type(T.clauseA);
  await page.keyboard.press('Tab'); // demote AFTER typing -> "(a)" (level 2)
  await page.keyboard.press('End');
  await page.keyboard.press('Enter'); // new item (level 2)
  await page.keyboard.type(T.clauseB); // sibling -> "(b)" (level 2)

  // 3) Whole-document serif font via Select-All + the font dropdown. All text is in place, so the
  //    change must reach every prose paragraph AND every legal-numbered list item.
  await page.keyboard.press('Control+a');
  await page.selectOption('#fontfamily', SERIF);

  // 4) Emphasis — clean ADDs onto plain blocks (no toggle-off needed).
  await ribbonFormatBlock(page, T.intro1.slice(0, 24), 'bold');
  await ribbonFormatBlock(page, T.intro2.slice(0, 24), 'bold');
  await ribbonFormatBlock(page, T.noteHead, 'bold');
  await ribbonFormatBlock(page, T.noteHead, 'italic');
  await ribbonFormatBlock(page, T.noteBody.slice(0, 24), 'italic');

  // 5) JUSTIFY the body paragraphs (disclaimer, note body, recital, (a)/(b) sub-clauses).
  await alignBlock(page, T.intro1.slice(0, 24), 'justify');
  await alignBlock(page, T.intro2.slice(0, 24), 'justify');
  await alignBlock(page, T.noteBody.slice(0, 24), 'justify');
  await alignBlock(page, 'THIS SERIES', 'justify');
  await alignBlock(page, T.clauseA.slice(0, 24), 'justify');
  await alignBlock(page, T.clauseB.slice(0, 24), 'justify');

  // 6) Centered bold title.
  await alignBlock(page, T.title, 'center');
  await ribbonFormatBlock(page, T.title, 'bold');

  // 7) Inline-bold the four defined terms and underline the Exhibit cross-reference in the recital.
  await emphasizeSpan(page, '"Agreement"', 'Agreement', 'bold');
  await emphasizeSpan(page, '"Company"', 'Company', 'bold');
  await emphasizeSpan(page, '"Purchaser"', 'Purchaser', 'bold');
  await emphasizeSpan(page, '"Purchasers"', 'Purchasers', 'bold');
  await emphasizeSpan(page, 'Exhibit A', 'Exhibit A', 'underline');

  // 8) The two clause captions are bold + underlined (legal section captions).
  await ribbonFormatBlock(page, T.clause1, 'bold');
  await ribbonFormatBlock(page, T.clause1, 'underline');
  await ribbonFormatBlock(page, T.clause11, 'bold');
  await ribbonFormatBlock(page, T.clause11, 'underline');

  // 9) Horizontal separator rule below the bold header (between the disclaimer and the note).
  await caretAtEndOf(page, T.intro2.slice(0, 24));
  await page.selectOption('#rulepos', 'below');
  await page.click('#hr');
}

test.describe('Editor — preferred stock purchase agreement cover page (full GUI reproduction)', () => {
  test('reproduces justify, whole-document serif, multi-term recital, title and legal numbering', async ({
    page,
  }) => {
    await buildCoverPage(page);
    const { blockCount, blocks } = await inspect(page);

    // 7 prose paragraphs + 1 separator rule + 4 numbered clauses
    expect(blockCount).toBe(12);

    const find = (tm: string) => blocks.find((b) => b.text.includes(tm))!;

    // bold + justified disclaimer (both lines)
    for (const line of [T.intro1, T.intro2]) {
      const b = find(line.slice(0, 24));
      expect(b.bold).toBe(true);
      expect(b.jc).toBe('both');
    }

    // bold + italic note heading (left aligned); italic-only justified note body
    const head = find(T.noteHead);
    expect(head.bold).toBe(true);
    expect(head.italic).toBe(true);
    expect(head.jc).not.toBe('center');
    const body = find(T.noteBody.slice(0, 24));
    expect(body.italic).toBe(true);
    expect(body.bold).toBe(false);
    expect(body.jc).toBe('both');

    // centered bold title
    const title = find(T.title);
    expect(title.jc).toBe('center');
    expect(title.bold).toBe(true);

    // recital: justified, FOUR inline-bold defined terms, an underlined cross-reference, not centered
    const recital = find('THIS SERIES');
    expect(recital.jc).toBe('both');
    expect(recital.boldRuns).toBe(4);
    expect(recital.underline).toBe(true);

    // multi-level legal numbering 1. / 1.1 / (a) / (b) sharing one continuous list
    const numbered = blocks.filter((b) => b.ilvl !== null);
    expect(numbered.map((b) => b.ilvl)).toEqual([0, 1, 2, 2]);
    expect(new Set(numbered.map((b) => b.numId)).size).toBe(1);

    // each numbered item keeps its OWN distinct text (no caret-misrouting collapse into one item)
    expect(numbered[0].text).toBe(T.clause1);
    expect(numbered[1].text).toBe(T.clause11);
    expect(numbered[2].text).toBe(T.clauseA);
    expect(numbered[3].text).toBe(T.clauseB);

    // captions bold + underlined; (a)/(b) bodies justified
    expect(numbered[0].bold).toBe(true);
    expect(numbered[0].underline).toBe(true);
    expect(numbered[1].underline).toBe(true);
    expect(numbered[2].jc).toBe('both');
    expect(numbered[3].jc).toBe('both');

    // the separator rule survived as a bottom-bordered empty paragraph
    expect(blocks.some((b) => b.hasBorder && b.text.trim() === '')).toBe(true);

    // whole-document serif: every text-bearing block carries Times New Roman, incl. the list items
    const textBlocks = blocks.filter((b) => b.text.trim() !== '');
    expect(textBlocks.every((b) => b.allTimes)).toBe(true);
    expect(numbered.every((b) => b.allTimes)).toBe(true);
  });

  test('the produced document is lossless across save -> reopen', async ({ page }) => {
    await buildCoverPage(page);

    const reopened = await page.evaluate(() => {
      const D = (window as any).__demo;
      const ed = D.getEditor();
      const saved: Uint8Array = ed.save();
      const Editor = (window as any).DocxodusEditor.DocxEditor;
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = Editor.open(c2, saved, D.exports, {});
      const blocks = [...c2.querySelectorAll('[data-anchor]')];
      const out = {
        blockCount: blocks.length,
        center: blocks.some((b) => getComputedStyle(b as HTMLElement).textAlign === 'center'),
        justify: blocks.some((b) => getComputedStyle(b as HTMLElement).textAlign === 'justify'),
        serif: blocks.some((b) => {
          const span = b.querySelector('span') as HTMLElement | null;
          return !!span && /Times/.test(getComputedStyle(span).fontFamily);
        }),
        ruleBorder: [...c2.querySelectorAll('div')].some(
          (d) =>
            /solid/.test(getComputedStyle(d as HTMLElement).borderBottom) &&
            getComputedStyle(d as HTMLElement).borderBottomWidth !== '0px',
        ),
        numbered: blocks
          .map((b) => (b.textContent || '').trim())
          .filter((t) => /^(1\.|1\.1|\(a\)|\(b\))/.test(t)).length,
      };
      e2.close();
      c2.remove();
      return out;
    });

    expect(reopened.blockCount).toBe(12);
    expect(reopened.center).toBe(true);
    expect(reopened.justify).toBe(true);
    expect(reopened.serif).toBe(true);
    expect(reopened.ruleBorder).toBe(true);
    expect(reopened.numbered).toBe(4);
  });

  // Regression guard for the editor's Enter-in-list caret handling: typing four clauses and pressing
  // Enter/Tab between them must produce four DISTINCT numbered items, each with its own text — not a
  // single item that has swallowed the rest. (Driven with native typing, the path a human uses.)
  test('legal-numbered list built by typing keeps each item distinct (1./1.1/(a)/(b))', async ({
    page,
  }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type(T.transition);
    await page.keyboard.press('Enter');
    await page.keyboard.type(T.clause1);
    await page.click('#legalNum');
    await page.keyboard.press('End');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Tab'); // set level on the empty item before typing (loss-free)
    await page.keyboard.type(T.clause11);
    await page.keyboard.press('Enter');
    await page.keyboard.press('Tab');
    await page.keyboard.type(T.clauseA);
    await page.keyboard.press('Enter');
    await page.keyboard.type(T.clauseB);

    const { blocks } = await inspect(page);
    const numbered = blocks.filter((b) => b.ilvl !== null);
    expect(numbered.map((b) => b.ilvl)).toEqual([0, 1, 2, 2]);
    expect(numbered.map((b) => b.text)).toEqual([T.clause1, T.clause11, T.clauseA, T.clauseB]);
    // no item swallowed another's text
    expect(numbered[3].text).not.toContain(T.clause1);
  });

  // Regression guard for the legal-numbering caret fix. A list-level change (Tab, or the §→/§←
  // buttons) re-renders the item from the session and now RESTORES the caret to its prior
  // end-of-line position (it previously jumped to offset 0). That makes the natural "type the
  // clause, press Tab to indent it, press Enter, type the next clause" flow loss-free — a following
  // Enter splits at the end, so the clauses stay distinct instead of migrating into the last item.
  // The cover-page build above relies on exactly this flow.
  test('Tab after typing a list item keeps the caret at the line end', async ({
    page,
  }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type('First clause heading');
    await page.click('#legalNum');
    await page.keyboard.press('End');
    await page.keyboard.press('Enter');
    await page.keyboard.type('Second clause heading');
    await page.keyboard.press('Tab'); // demote AFTER typing — caret should stay at the end
    await page.keyboard.press('Enter'); // a continuation Enter must NOT migrate "Second clause heading"
    await page.keyboard.type('Third clause heading');
    const { blocks } = await inspect(page);
    const numbered = blocks.filter((b) => b.ilvl !== null).map((b) => b.text);
    expect(numbered).toEqual(['First clause heading', 'Second clause heading', 'Third clause heading']);
  });

  // ── History UX bug found while reproducing the cover page ────────────────────────────────────
  // The ribbon Undo/Redo BUTTONS always work — they call editor.undo()/redo() directly.
  test('the Redo button restores an undone rule', async ({ page }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type('A line of body text.');
    await page.selectOption('#rulepos', 'below');
    await page.click('#hr');
    const count = () => page.evaluate(() => document.querySelectorAll('#editor [data-anchor]').length);
    const withRule = await count();
    await page.click('#undo');
    expect(await count()).toBe(withRule - 1);
    await page.click('#redo');
    expect(await count()).toBe(withRule);
  });

  // Regression guard: keyboard undo/redo must keep working across CONSECUTIVE presses without the
  // user clicking back into a block. undo()/redo() re-render via remount, which now restores the
  // caret into a block (and sets activeBlock) so the root keydown handler can resolve a block for
  // the next shortcut; previously remount(-1, ...) left the caret on the bare contenteditable host
  // with activeBlock null, so a SECOND consecutive keyboard history op was silently dropped (the
  // event reached the host but keydownBlock returned null). The Undo/Redo buttons always worked
  // (they bypass keydownBlock); this guards the keyboard path.
  test('a second consecutive keyboard history op works without re-focusing a block', async ({
    page,
  }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type('A line of body text.');
    await page.selectOption('#rulepos', 'below');
    await page.click('#hr');
    const count = () => page.evaluate(() => document.querySelectorAll('#editor [data-anchor]').length);
    const withRule = await count();
    await page.locator('#editor [data-anchor]').first().click(); // caret in a real block
    await page.keyboard.press('Control+z'); // first keyboard undo removes the rule
    expect(await count()).toBe(withRule - 1);
    await page.keyboard.press('Control+Shift+z'); // second keyboard op WITHOUT re-clicking a block
    expect(await count()).toBe(withRule); // redo restores the rule (was dropped before the fix)
  });

  // ── Engine-level omissions: features a real agreement page needs but the editor cannot author. ──
  // The demo ribbon must not pretend to offer footnote / footer / page-number controls.
  test('the ribbon exposes no footnote, footer or page-number control', async ({ page }) => {
    await boot(page);
    const joined = await page.evaluate(() =>
      [...document.querySelectorAll('#ribbon button, #ribbon select, header button')]
        .map((el) => (el.getAttribute('title') || el.textContent || '').toLowerCase())
        .join(' | '),
    );
    expect(joined).not.toMatch(/footnote/);
    expect(joined).not.toMatch(/footer/);
    expect(joined).not.toMatch(/page number|page-number/);
  });

  // The contract's bottom-of-page footnote ("If only one closing is contemplated...") cannot be
  // authored: there is no footnote command in the ribbon and no DocxSession.AddFootnote; the editor's
  // full render hard-codes renderFootnotesAndEndnotes:false and markdown footnote refs are rejected.
  test.fixme('footnote creation (superscript marker + bottom-of-page text) — unsupported', async () => {
    // Re-enable when a footnote-authoring API ships and a ribbon control is wired to it.
  });

  // The footer line ("Last Updated October 2025") cannot be authored: DocxSession only reads existing
  // FooterPartUris and has no footer-part create path; the editor merely displays a pre-existing
  // footer in paginated mode.
  test.fixme('footer text creation ("Last Updated ...") — unsupported', async () => {
    // Re-enable when a footer-part create API ships.
  });

  // The centered page number is a PAGE field code living inside a (non-creatable) footer; there is no
  // field-insertion API anywhere in the editor or DocxSession.
  test.fixme('centered page-number field — unsupported', async () => {
    // Re-enable when PAGE-field insertion + footer creation ship.
  });
});
