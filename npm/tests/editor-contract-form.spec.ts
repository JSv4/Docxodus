import { test, expect, Page } from '@playwright/test';

/**
 * GUI-driven reproduction of a legal "preferred stock purchase agreement" cover page using ONLY the
 * demo's ribbon control surfaces (real toolbar button clicks). It exercises the formatting
 * vocabulary that such a contract's first page relies on:
 *
 *   - a bold multi-line disclaimer block,
 *   - a horizontal separator rule,
 *   - a bold + italic "Preliminary Note" heading and an italic note body,
 *   - a CENTERED BOLD document title,
 *   - a recital paragraph with an inline-bold defined term,
 *   - multi-level legal numbering 1. / 1.1 / (a) / (b) with hanging indents.
 *
 * All text is invented placeholder text — never any copyrighted contract language.
 *
 * The suite saves the produced .docx and asserts the resulting OOXML (w:b / w:i / w:u / w:jc /
 * w:numPr+w:ilvl / w:pBdr) so the editor's coverage is verified end to end, and it records the
 * three engine-level omissions a real agreement page needs but the editor cannot author
 * (footnotes, footer text, page numbers) plus one multi-block formatting gap.
 */

// Invented placeholder text (structurally similar, no copyrighted language).
const T = {
  intro1:
    'This sample agreement is provided for demonstration purposes only and should be tailored before use.',
  intro2:
    'Provisions introduced here are marked NEW and provisions revised here are marked REVISED in this sample.',
  noteHead: 'Preliminary Note',
  noteBody:
    'This note summarises the general structure of the demonstration agreement and the matters it does and does not cover.',
  title: 'DEMONSTRATION PREFERRED STOCK PURCHASE AGREEMENT',
  // recital is typed as one sentence; only the quoted defined term is bolded.
  recital:
    'THIS DEMONSTRATION AGREEMENT (this "Agreement") is entered into by the company and the purchasers identified on the attached schedule.',
  recitalTerm: 'Agreement',
  transition: 'The parties hereby agree as follows:',
  clause1: 'Purchase and Sale.',
  clause11: 'Sale and Issuance.',
  clauseA: 'The company shall file its charter documents on or before the first closing.',
  clauseB: 'Each purchaser shall purchase the number of shares set opposite its name on the schedule.',
};

async function boot(page: Page) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
}

/** Type text at the caret through the editor's own input path (execCommand insertText). */
async function typeAtCaret(page: Page, text: string) {
  await page.evaluate((t) => {
    document.execCommand('insertText', false, t);
  }, text);
}

/** Place a collapsed caret at the end of the first editable body block (the seed paragraph). */
async function focusSeed(page: Page) {
  await page.evaluate(() => {
    const p = document.querySelector('#editor p[data-anchor][data-editable="1"]') as HTMLElement;
    p.focus();
    const r = document.createRange();
    r.selectNodeContents(p);
    r.collapse(false);
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    document.dispatchEvent(new Event('selectionchange'));
  });
}

/** Select the entire contents of the body block whose text contains `textMatch`. */
async function selectWholeBlock(page: Page, textMatch: string) {
  await page.evaluate((tm) => {
    const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find(
      (p) => (p.textContent || '').includes(tm),
    ) as HTMLElement | undefined;
    if (!blk) throw new Error('block not found for: ' + tm);
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

/** Save the document and project it; return per-block OOXML facts keyed off the anchor index. */
async function inspect(page: Page) {
  return page.evaluate(() => {
    const D = (window as any).__demo;
    const B = D.exports.DocxSessionBridge;
    const ed = D.getEditor();
    const saved: Uint8Array = ed.save();
    const h = B.OpenSession(saved, '');
    const proj = JSON.parse(B.Project(h));
    const anchors: string[] = [...(proj.markdown as string).matchAll(/\{#((?:p|li|h\d):body:[0-9a-f]+)\}/g)].map(
      (m) => m[1],
    );
    const blocks = anchors.map((a) => {
      const xml: string = B.RawGetXml(h, a);
      const text = (xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
        .map((t) => t.replace(/<[^>]+>/g, ''))
        .join('');
      const ilvlM = xml.match(/<w:ilvl w:val="(\d+)"/);
      const numIdM = xml.match(/<w:numId w:val="(\d+)"/);
      // count emphasised runs (w:b, not w:bCs; w:i, not w:ind/w:iCs)
      const bolds = (xml.match(/<w:b\b[^>]*\/?>/g) || []).filter((m) => !/w:bCs/.test(m)).length;
      const italics = (xml.match(/<w:i\b[^>]*\/?>/g) || []).filter((m) => !/w:iCs|w:ind/.test(m)).length;
      return {
        anchorKind: a.split(':')[0],
        text,
        bold: /<w:b\b[^>]*\/?>/.test(xml) && !/^<w:bCs/.test(xml),
        boldRuns: bolds,
        italic: italics > 0,
        underline: /<w:u\b/.test(xml),
        center: /<w:jc w:val="center"/.test(xml),
        ilvl: ilvlM ? Number(ilvlM[1]) : null,
        numId: numIdM ? Number(numIdM[1]) : null,
        hasBorder: /<w:pBdr/.test(xml),
        timesFont: /w:ascii="Times New Roman"/.test(xml),
      };
    });
    B.CloseSession(h);
    return { blockCount: blocks.length, blocks };
  });
}

/** Build the whole contract cover page through the ribbon. */
async function buildContractForm(page: Page) {
  await boot(page);
  await focusSeed(page);

  // 1) Type the prose blocks as PLAIN text, one block per logical paragraph (Enter splits a block;
  //    starting from a plain seed, nothing inherits formatting — every later format is a clean ADD).
  await typeAtCaret(page, T.intro1);
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.intro2);
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.noteHead);
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.noteBody);
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.title);
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.recital);
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.transition);

  // 2) The four legal-numbered clauses: apply the legal scheme once, then Enter continues the list
  //    and Tab demotes the level (1. -> Tab -> 1.1 -> Tab -> (a); sibling (b) continues at level 2).
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.clause1);
  await page.click('#legalNum'); // -> "1." (level 0)
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.clause11);
  await page.keyboard.press('Tab'); // -> "1.1" (level 1)
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.clauseA);
  await page.keyboard.press('Tab'); // -> "(a)" (level 2)
  await page.keyboard.press('Enter');
  await typeAtCaret(page, T.clauseB); // -> "(b)" (level 2, sibling)

  // 3) Emphasis — all ADDs onto plain blocks (no toggle-off needed).
  await ribbonFormatBlock(page, T.intro1, 'bold');
  await ribbonFormatBlock(page, T.intro2, 'bold');
  await ribbonFormatBlock(page, T.noteHead, 'bold');
  await ribbonFormatBlock(page, T.noteHead, 'italic');
  await ribbonFormatBlock(page, T.noteBody, 'italic');

  // 4) Centered bold title (paragraph alignment + bold).
  await selectWholeBlock(page, T.title);
  await page.click('#ribbon button[data-align="center"]');
  await ribbonFormatBlock(page, T.title, 'bold');

  // 5) Inline-bold defined term: bold ONLY the quoted word "Agreement" in the recital.
  await page.evaluate((term) => {
    const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find((p) =>
      (p.textContent || '').includes('"' + term + '"'),
    ) as HTMLElement;
    const node = [...blk.querySelectorAll('span')]
      .map((s) => s.firstChild)
      .find((n) => n && n.textContent && n.textContent.includes('"' + term + '"')) as Text;
    const full = node.textContent!;
    const start = full.indexOf('"' + term + '"') + 1;
    const r = document.createRange();
    r.setStart(node, start);
    r.setEnd(node, start + term.length);
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    document.dispatchEvent(new Event('selectionchange'));
  }, T.recitalTerm);
  await page.click('#ribbon button[data-cmd="bold"]');

  // 6) Underline the two clause headings (bold + underline, as legal section captions usually are).
  await ribbonFormatBlock(page, T.clause1, 'bold');
  await ribbonFormatBlock(page, T.clause1, 'underline');
  await ribbonFormatBlock(page, T.clause11, 'bold');
  await ribbonFormatBlock(page, T.clause11, 'underline');

  // 7) Horizontal separator rule below the disclaimer (rulepos defaults to "below").
  await page.evaluate((tm) => {
    const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find((p) =>
      (p.textContent || '').includes(tm),
    ) as HTMLElement;
    blk.focus();
    const r = document.createRange();
    r.selectNodeContents(blk);
    r.collapse(false);
    const s = window.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    document.dispatchEvent(new Event('selectionchange'));
  }, T.intro2);
  await page.selectOption('#rulepos', 'below');
  await page.click('#hr');
}

test.describe('Editor — legal contract cover page (GUI reproduction)', () => {
  test('reproduces emphasis, centered title, separator rule and multi-level legal numbering', async ({
    page,
  }) => {
    await buildContractForm(page);
    const { blockCount, blocks } = await inspect(page);

    // structure: 7 prose paragraphs + 1 rule paragraph + 4 numbered clauses
    expect(blockCount).toBe(12);

    const find = (tm: string) => blocks.find((b) => b.text.includes(tm))!;

    // bold disclaimer (both lines)
    expect(find(T.intro1.slice(0, 20)).bold).toBe(true);
    expect(find(T.intro2.slice(0, 20)).bold).toBe(true);

    // bold + italic note heading; italic-only note body
    const head = find(T.noteHead);
    expect(head.bold).toBe(true);
    expect(head.italic).toBe(true);
    const body = find(T.noteBody.slice(0, 20));
    expect(body.italic).toBe(true);
    expect(body.bold).toBe(false);

    // centered bold title
    const title = find(T.title);
    expect(title.center).toBe(true);
    expect(title.bold).toBe(true);

    // recital: exactly one inline-bold run (the defined term), not the whole paragraph
    const recital = find('attached schedule');
    expect(recital.boldRuns).toBe(1);
    expect(recital.center).toBe(false);

    // multi-level legal numbering 1. / 1.1 / (a) / (b)
    const numbered = blocks.filter((b) => b.ilvl !== null);
    expect(numbered.map((b) => b.ilvl)).toEqual([0, 1, 2, 2]);
    // all four share one numbering instance (a single continuous list)
    const numIds = new Set(numbered.map((b) => b.numId));
    expect(numIds.size).toBe(1);
    // the two captions are bold + underlined
    expect(find(T.clause1).bold).toBe(true);
    expect(find(T.clause1).underline).toBe(true);

    // the separator rule survived as a bottom-bordered empty paragraph
    expect(blocks.some((b) => b.hasBorder && b.text.trim() === '')).toBe(true);
  });

  test('the produced document is lossless across save → reopen', async ({ page }) => {
    await buildContractForm(page);

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
    expect(reopened.ruleBorder).toBe(true);
    expect(reopened.numbered).toBe(4); // 1. / 1.1 / (a) / (b) all reopened
  });

  // ── Known multi-block formatting gap (live characterisation test) ─────────────────────────────
  // A document-wide font change (Ctrl+A then pick a font) reaches plain prose blocks but currently
  // SKIPS legal-numbered list items — the synthetic numbering marker throws off the committed-span
  // offset mapping in the multi-block inline path. This test pins the CURRENT behaviour; if the
  // engine is fixed to cover list items it will fail, which is the signal to update this assertion.
  test('multi-block font change reaches prose but currently skips legal-numbered list items', async ({
    page,
  }) => {
    await boot(page);
    await focusSeed(page);
    await typeAtCaret(page, 'Alpha plain prose line');
    await page.keyboard.press('Enter');
    await typeAtCaret(page, 'Beta numbered clause');
    await page.click('#legalNum'); // make the second block a legal-numbered list item

    await page.evaluate(() => {
      const root = document.querySelector('#editor') as HTMLElement;
      root.focus();
    });
    await page.keyboard.press('Control+a');
    await page.selectOption('#fontfamily', 'Georgia');

    const fonts = await page.evaluate(() =>
      [...document.querySelectorAll('#editor [data-anchor]')].map((b) => {
        const span = b.querySelector('span');
        return {
          text: (b.textContent || '').slice(0, 12),
          font: span ? getComputedStyle(span).fontFamily : '',
        };
      }),
    );

    const prose = fonts.find((f) => f.text.includes('Alpha'))!;
    const listItem = fonts.find((f) => /Beta/.test(f.text))!;
    expect(prose.font).toContain('Georgia'); // prose block picks up the new font
    expect(listItem.font).not.toContain('Georgia'); // list item is skipped (known gap)
  });

  // ── Engine-level omissions: features a real agreement page needs but the editor cannot author ──
  // The demo ribbon must not pretend to offer footnote / footer / page-number controls.
  test('the ribbon exposes no footnote, footer or page-number control', async ({ page }) => {
    await boot(page);
    const labels = await page.evaluate(() =>
      [...document.querySelectorAll('#ribbon button, #ribbon select, header button')].map(
        (el) => (el.getAttribute('title') || el.textContent || '').toLowerCase(),
      ),
    );
    const joined = labels.join(' | ');
    expect(joined).not.toMatch(/footnote/);
    expect(joined).not.toMatch(/footer/);
    expect(joined).not.toMatch(/page number|page-number/);
  });

  // Footnotes are output-only in the engine (DocxSession has no AddFootnote; the editor's full
  // render hard-codes renderFootnotesAndEndnotes:false; markdown footnote refs are rejected with
  // FootnoteRefNotSupported). The contract's "Initial Closing¹ … (footnote)" cannot be authored.
  test.fixme('footnote creation (superscript marker + bottom-of-page text) — unsupported', async () => {
    // No GUI path: there is no footnote command and no DocxSession.AddFootnote. Re-enable when a
    // footnote-authoring API ships and wire the editor + a ribbon control to it.
  });

  // Footers/headers have no create path in DocxSession (it only reads FooterPartUris); the editor
  // merely displays a pre-existing footer in paginated mode. "Last Updated October 2025" cannot be
  // authored through the editor.
  test.fixme('footer text creation ("Last Updated …") — unsupported', async () => {
    // Re-enable when a footer-part create API ships.
  });

  // Page numbers are PAGE field codes that live inside a (non-creatable) footer; there is no
  // field-insertion API anywhere in the editor or DocxSession.
  test.fixme('centered page-number field — unsupported', async () => {
    // Re-enable when PAGE-field insertion + footer creation ship.
  });
});
