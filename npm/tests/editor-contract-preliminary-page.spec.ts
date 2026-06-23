import { test, expect, Page } from '@playwright/test';

/**
 * GUI-driven smoke test of the "Preliminary Note" first page of a model preferred-stock purchase
 * agreement, focused on the surfaces a fresh reproduction pass surfaced that the companion
 * editor-preferred-stock-cover.spec.ts does NOT cover:
 *
 *   1. a SUPERSCRIPT footnote MARKER authored through the ribbon x² button — rendered as <sup> with
 *      w:vertAlign="superscript" in the OOXML — and the proof that it is only a DECORATIVE glyph,
 *      NOT a real linked w:footnoteReference (the editor / DocxSession cannot author footnotes);
 *   2. the MULTI-BLOCK Select-All gesture: bold + justify applied to two paragraphs in ONE action
 *      (Ctrl+A in the body selects every body block, then a single ribbon click formats them all);
 *   3. an underlined SUB-WORD inside an italic, justified note paragraph (italic + underline compose
 *      on the same run while the paragraph stays justified);
 *   4. the inline-bold TOGGLE semantics that govern a single defined term: bolding a word that is
 *      ALREADY bold REMOVES bold from it (exactly like Word). A paragraph split off a bold line
 *      inherits bold, so authoring one bold defined term needs the surrounding text non-bold first;
 *   5. the positive counterpart — four inline-bold defined terms + an underlined cross-reference
 *      author correctly in a NON-bold recital.
 *
 * All prose is invented placeholder text — never any copyrighted contract language. Text is entered
 * with NATIVE keyboard input (page.keyboard.type), the path a human types through; programmatic
 * Ranges are used only to place a caret/selection, mirroring the companion spec's helpers.
 */

// Invented placeholder text (structurally similar to a financing page, no copyrighted language).
const T = {
  intro1:
    'This sample agreement was prepared by a demonstration drafting group for general reference and is intended to serve as a starting framework only; it must be adapted to the facts of each transaction and should not be construed as legal advice.',
  intro2:
    'Reviewers comparing this draft against an earlier edition will find that newly added annotations are labelled New, while annotations carried over with material changes are labelled Revised.',
  noteBody:
    'The purchase agreement records the core commercial terms on which the preferred shares are sold to the incoming investors. It does not by itself describe the characteristics of the shares, which are set out in the charter, nor the wider arrangements among the holders.',
  recital:
    'THIS SERIES [___] PREFERRED STOCK PURCHASE AGREEMENT (this "Agreement") is entered into as of [____], 20[__] by and among [____________], a Delaware corporation (the "Company"), and the investors listed on Exhibit A attached to this Agreement (each a "Purchaser" and together the "Purchasers").',
  clauseA:
    'The Company shall adopt and file with the Secretary of State of Delaware, on or before the initial closing, an Amended and Restated Certificate of Incorporation in the form of Exhibit B attached to this Agreement.',
};

const SEED = '#editor p[data-anchor][data-editable="1"]';

async function boot(page: Page) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
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

/** Select an inner word inside the block containing `blockMatch`, then click an inline ribbon command. */
async function emphasizeWord(page: Page, blockMatch: string, phrase: string, inner: string, cmd: string) {
  await page.evaluate(
    ({ blockMatch, phrase, inner }) => {
      const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find(
        (p) => (p.textContent || '').includes(blockMatch),
      ) as HTMLElement;
      if (!blk) throw new Error('block not found: ' + blockMatch);
      const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
      let target: Text | null = null;
      let idx = -1;
      let n: Node | null;
      while ((n = walker.nextNode())) {
        const i = (n.textContent || '').indexOf(phrase);
        if (i >= 0) {
          target = n as Text;
          idx = i + (n.textContent || '').indexOf(inner, i) - i;
          break;
        }
      }
      if (!target) throw new Error('phrase not found: ' + phrase);
      const r = document.createRange();
      r.setStart(target, idx);
      r.setEnd(target, idx + inner.length);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
      document.dispatchEvent(new Event('selectionchange'));
    },
    { blockMatch, phrase, inner },
  );
  await page.click(`#ribbon button[data-cmd="${cmd}"]`);
}

/** Per-block OOXML facts (block-level flags), keyed off the projection anchor index. */
async function inspect(page: Page) {
  return page.evaluate(() => {
    const D = (window as any).__demo;
    const B = D.exports.DocxSessionBridge;
    const ed = D.getEditor();
    ed.commitAllDirty();
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
      const jcM = xml.match(/<w:jc w:val="(\w+)"/);
      const boldRuns = (xml.match(/<w:b\b[^>]*\/?>/g) || []).filter((m) => !/w:bCs/.test(m)).length;
      return {
        kind: a.split(':')[0],
        text,
        jc: jcM ? jcM[1] : null,
        boldRuns,
        bold: /<w:b\b[^>]*\/?>/.test(xml) && !/^<w:bCs/.test(xml),
        underline: /<w:u\b/.test(xml),
        superscript: /<w:vertAlign w:val="superscript"/.test(xml),
        footnoteRef: /w:footnoteReference/.test(xml),
      };
    });
    B.CloseSession(h);
    return { blockCount: blocks.length, blocks };
  });
}

/** Run-level facts for the single block containing `blockMatch`: text + b/i/u/vertAlign per run. */
async function inspectRuns(page: Page, blockMatch: string) {
  return page.evaluate((bm) => {
    const D = (window as any).__demo;
    const B = D.exports.DocxSessionBridge;
    const ed = D.getEditor();
    ed.commitAllDirty();
    const h = B.OpenSession(ed.save(), '');
    const anchors: string[] = [
      ...(JSON.parse(B.Project(h)).markdown as string).matchAll(/\{#((?:p|li|h\d):body:[0-9a-f]+)\}/g),
    ].map((m) => m[1]);
    let xml = '';
    for (const a of anchors) {
      const x = B.RawGetXml(h, a);
      // Match on the block's JOINED text — an inline format splits the phrase across <w:t> runs, so
      // the contiguous substring no longer exists in the raw markup.
      const joined = (x.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [])
        .map((s) => s.replace(/<[^>]+>/g, ''))
        .join('');
      if (joined.includes(bm)) {
        xml = x;
        break;
      }
    }
    B.CloseSession(h);
    const runs: { t: string; b: boolean; i: boolean; u: boolean; sup: boolean }[] = [];
    const runRe = /<w:r\b[\s\S]*?<\/w:r>/g;
    let m: RegExpExecArray | null;
    while ((m = runRe.exec(xml))) {
      const run = m[0];
      const t = (run.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/) || ['', ''])[1];
      if (!t) continue;
      runs.push({
        t,
        b: /<w:b\b(?![A-Za-z])[^>]*\/?>/.test(run) && !/<w:bCs/.test(run.replace(/<w:b\b(?![A-Za-z])[^>]*\/?>/, '')),
        i: /<w:i\b(?![A-Za-z])[^>]*\/?>/.test(run) && !/<w:iCs/.test(run.replace(/<w:i\b(?![A-Za-z])[^>]*\/?>/, '')),
        u: /<w:u\b/.test(run),
        sup: /<w:vertAlign w:val="superscript"/.test(run),
      });
    }
    const jc = (xml.match(/<w:jc w:val="(\w+)"/) || [, null])[1];
    const footnoteRef = /w:footnoteReference/.test(xml);
    return { runs, jc, footnoteRef };
  }, blockMatch);
}

test.describe('Editor — contract preliminary-note page (footnote-marker, multi-block, toggle)', () => {
  // ── 1. The page's footnote MARKER: a decorative superscript, not a real linked footnote. ──────
  test('a footnote marker is authorable as a superscript glyph but is NOT a real footnote ref', async ({
    page,
  }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type(T.clauseA);

    // Append the footnote marker digit and superscript it through the ribbon x² button.
    await page.keyboard.press('End');
    await page.keyboard.type('1');
    await page.evaluate(() => {
      const blk = [...document.querySelectorAll('#editor [data-anchor][data-editable="1"]')].find((p) =>
        /Company shall adopt/.test(p.textContent || ''),
      ) as HTMLElement;
      const walker = document.createTreeWalker(blk, NodeFilter.SHOW_TEXT);
      let last: Text | null = null;
      let n: Node | null;
      while ((n = walker.nextNode())) last = n as Text;
      const len = (last!.textContent || '').length;
      const r = document.createRange();
      r.setStart(last!, len - 1); // the trailing "1"
      r.setEnd(last!, len);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
      document.dispatchEvent(new Event('selectionchange'));
    });
    await page.click('#ribbon button[data-cmd="superscript"]');

    // The marker is rendered as a <sup> in the editor surface...
    const hasSup = await page.evaluate(() => !!document.querySelector('#editor sup'));
    expect(hasSup).toBe(true);

    // ...and carries w:vertAlign="superscript" in the OOXML, but is plain text — NOT a footnote ref.
    const { runs, footnoteRef } = await inspectRuns(page, 'Company shall adopt');
    expect(runs.some((r) => r.t === '1' && r.sup)).toBe(true);
    expect(footnoteRef).toBe(false);

    // And the editor never produced a footnotes part the marker could link into: no real footnote.
    const footnoteRefAnywhere = await page.evaluate(() => {
      const D = (window as any).__demo;
      const B = D.exports.DocxSessionBridge;
      const ed = D.getEditor();
      const h = B.OpenSession(ed.save(), '');
      const md: string = JSON.parse(B.Project(h)).markdown;
      const anchors = [...md.matchAll(/\{#((?:p|li|h\d|fn):[a-z]+:[0-9a-f]+)\}/g)].map((m) => m[1]);
      let anyRef = false;
      for (const a of anchors) if (/w:footnoteReference/.test(B.RawGetXml(h, a))) anyRef = true;
      B.CloseSession(h);
      // a real footnote would surface an fn:* anchor; a decorative superscript does not
      return { anyRef, hasFootnoteAnchor: anchors.some((a) => a.startsWith('fn:')) };
    });
    expect(footnoteRefAnywhere.anyRef).toBe(false);
    expect(footnoteRefAnywhere.hasFootnoteAnchor).toBe(false);
  });

  // ── 2. Multi-block Select-All: bold + justify two paragraphs in ONE gesture each. ─────────────
  test('Select-All applies bold and justify to the intro paragraphs in one gesture', async ({ page }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type(T.intro1);
    await page.keyboard.press('Enter');
    await page.keyboard.type(T.intro2);

    // Ctrl+A selects every body block; one ribbon Bold click bolds both. Re-select (the re-render
    // invalidates the prior selection), then one Justify click justifies both.
    await page.click(SEED);
    await page.keyboard.press('Control+a');
    await page.click('#ribbon button[data-cmd="bold"]');
    await page.keyboard.press('Control+a');
    await page.click('#ribbon button[data-align="justify"]');

    const { blocks } = await inspect(page);
    const intro = blocks.filter((b) => /sample agreement|Reviewers comparing/.test(b.text));
    expect(intro.length).toBe(2);
    for (const b of intro) {
      expect(b.bold).toBe(true);
      expect(b.jc).toBe('both');
    }
  });

  // ── 3. Underlined sub-word inside an italic, justified note paragraph. ────────────────────────
  test('underlining one word inside an italic justified note composes italic + underline', async ({
    page,
  }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type(T.noteBody);

    await selectWholeBlock(page, T.noteBody.slice(0, 24));
    await page.click('#ribbon button[data-cmd="italic"]');
    await selectWholeBlock(page, T.noteBody.slice(0, 24));
    await page.click('#ribbon button[data-align="justify"]');

    // Underline only the word "not" in "It does not by itself".
    await emphasizeWord(page, 'does not by itself', 'not by itself', 'not', 'underline');

    const { runs, jc } = await inspectRuns(page, 'does not by itself');
    expect(jc).toBe('both');
    const notRun = runs.find((r) => r.t.trim() === 'not');
    expect(notRun, 'the standalone "not" run exists').toBeTruthy();
    expect(notRun!.u).toBe(true); // underlined
    expect(notRun!.i).toBe(true); // still italic (composes)
    // the rest of the paragraph is italic but NOT underlined
    expect(runs.filter((r) => r.u).length).toBe(1);
    expect(runs.every((r) => r.i)).toBe(true);
  });

  // ── 4. Inline-bold TOGGLE semantics: bolding an already-bold word REMOVES bold from it. ───────
  // This is the behaviour that, applied to a paragraph that inherited bold from a preceding bold
  // line, makes "bolding a defined term" look like it bolded the complement — the toggle is simply
  // removing bold from the already-bold word. Word behaves identically. The authoring takeaway is in
  // test 5: to add a single bold term, the surrounding text must be non-bold first.
  test('bolding an already-bold word removes bold from it (toggle), not the rest', async ({ page }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type('alpha beta gamma');

    // Make the whole line bold...
    await selectWholeBlock(page, 'alpha beta gamma');
    await page.click('#ribbon button[data-cmd="bold"]');
    // ...then bold the already-bold "beta": it toggles OFF, leaving "alpha " and " gamma" bold.
    await emphasizeWord(page, 'alpha beta gamma', 'beta', 'beta', 'bold');

    const { runs } = await inspectRuns(page, 'alpha beta gamma');
    const beta = runs.find((r) => r.t.trim() === 'beta');
    expect(beta, 'the "beta" run is split out').toBeTruthy();
    expect(beta!.b).toBe(false); // toggled OFF
    // the surrounding words stay bold (this is the "complement bold" appearance, by design)
    expect(runs.filter((r) => r.b && r.t.trim() !== 'beta').every((r) => r.b)).toBe(true);
    expect(runs.some((r) => /alpha/.test(r.t) && r.b)).toBe(true);
    expect(runs.some((r) => /gamma/.test(r.t) && r.b)).toBe(true);
  });

  // ── 4b. Whole-block toggle direction: Bold must REMOVE bold from a fully-bold paragraph that is
  // selected as a WHOLE (triple-click / select-all → the selection's start container is the <p>
  // element, not a text node). Before the fix, selectionHasFormat read the block element's computed
  // weight (normal — the boldness lives on the inner run spans), so the toggle re-ADDED bold and a
  // fully-bold paragraph could never be un-bolded this way.
  test('Bold removes bold from a fully-bold paragraph selected as a whole block', async ({ page }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type('alpha beta gamma');

    // Make the whole paragraph bold (a clean ADD onto plain text).
    await selectWholeBlock(page, 'alpha beta gamma');
    await page.click('#ribbon button[data-cmd="bold"]');

    // Re-select the whole paragraph the SAME (element-container) way and click Bold — it must toggle OFF.
    await selectWholeBlock(page, 'alpha beta gamma');
    await page.click('#ribbon button[data-cmd="bold"]');

    const { blocks } = await inspect(page);
    const b = blocks.find((x) => x.text.includes('alpha beta gamma'))!;
    expect(b.bold).toBe(false);
    expect(b.boldRuns).toBe(0);
  });

  // ── 5. The positive path: four inline-bold defined terms + an underlined ref in a non-bold recital. ──
  test('four defined terms inline-bold and a cross-reference underlines in a non-bold recital', async ({
    page,
  }) => {
    await boot(page);
    await page.click(SEED);
    await page.keyboard.type(T.recital);
    await selectWholeBlock(page, 'THIS SERIES');
    await page.click('#ribbon button[data-align="justify"]');

    // Surrounding text is plain, so each Bold is a clean ADD onto a non-bold run.
    await emphasizeWord(page, 'THIS SERIES', '"Agreement"', 'Agreement', 'bold');
    await emphasizeWord(page, 'THIS SERIES', '"Company"', 'Company', 'bold');
    await emphasizeWord(page, 'THIS SERIES', '"Purchaser"', 'Purchaser', 'bold');
    await emphasizeWord(page, 'THIS SERIES', '"Purchasers"', 'Purchasers', 'bold');
    await emphasizeWord(page, 'THIS SERIES', 'Exhibit A', 'Exhibit A', 'underline');

    const { blocks } = await inspect(page);
    const recital = blocks.find((b) => b.text.includes('THIS SERIES'))!;
    expect(recital.jc).toBe('both');
    expect(recital.boldRuns).toBe(4); // exactly the four defined terms
    expect(recital.underline).toBe(true); // the Exhibit cross-reference
  });
});
