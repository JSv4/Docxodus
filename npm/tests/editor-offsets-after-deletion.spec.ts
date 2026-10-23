import { test, expect, Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// With tracked changes rendered inline, a paragraph shows its deleted text as a <del>. The
// session's run text does not hold that text, so the editor must leave it out of the offsets it
// sends the session; counting it put every position to the right of a deletion past where the
// session has it, and typing there was refused with offset_out_of_range or landed in the wrong
// place (issue #1278).

const REVISION = 'w:author="Counsel A" w:date="2026-05-01T09:00:00Z"';

function deletionFixture(): Uint8Array {
  const body = [
    // A deletion between live runs, followed by an insertion.
    '<w:p><w:r><w:t xml:space="preserve">The tenant shall </w:t></w:r>'
      + `<w:del w:id="1" ${REVISION}><w:r><w:delText xml:space="preserve">promptly </w:delText></w:r></w:del>`
      + '<w:r><w:t xml:space="preserve">repaint the fence</w:t></w:r>'
      + `<w:ins w:id="2" ${REVISION}><w:r><w:t xml:space="preserve"> each spring</w:t></w:r></w:ins>`
      + '<w:r><w:t>.</w:t></w:r></w:p>',
    // A move source inside a hyperlink, written as Word writes it (w:t, not w:delText). It is not a
    // direct child of the paragraph, so the paragraph stays editable. Its text is gone from the
    // paragraph a reader sees, so neither the session nor the editor counts it (issue #1282).
    '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="terms">'
      + `<w:moveFromRangeStart w:id="3" w:name="move1" ${REVISION}/>`
      + `<w:moveFrom w:id="4" ${REVISION}><w:r><w:t xml:space="preserve">the terms </w:t></w:r></w:moveFrom>`
      + '<w:moveFromRangeEnd w:id="3"/>'
      + '<w:r><w:t>below</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> now.</w:t></w:r></w:p>',
    `<w:p><w:moveToRangeStart w:id="5" w:name="move1" ${REVISION}/>`
      + `<w:moveTo w:id="6" ${REVISION}><w:r><w:t xml:space="preserve">the terms </w:t></w:r></w:moveTo>`
      + '<w:moveToRangeEnd w:id="5"/><w:r><w:t>Moved here.</w:t></w:r></w:p>',
    // A deleted paragraph mark, which the editor shows as a pilcrow.
    `<w:p><w:pPr><w:rPr><w:del w:id="7" ${REVISION}/></w:rPr></w:pPr><w:r><w:t>Joined to the next.</w:t></w:r></w:p>`,
    // Control: no revisions at all.
    '<w:p><w:r><w:t>Plain control paragraph.</w:t></w:r></w:p>',
    // A paragraph whose only text is deleted, inside a bookmark (a heading's TOC target), and the
    // same without the bookmark.
    '<w:p><w:bookmarkStart w:id="10" w:name="_Toc1"/>'
      + `<w:del w:id="11" ${REVISION}><w:r><w:delText>Gone heading</w:delText></w:r></w:del>`
      + '<w:bookmarkEnd w:id="10"/></w:p>',
    `<w:p><w:del w:id="12" ${REVISION}><w:r><w:delText>Gone too</w:delText></w:r></w:del></w:p>`,
    // A deletion at the end of a paragraph.
    '<w:p><w:r><w:t>Keep this</w:t></w:r>'
      + `<w:del w:id="13" ${REVISION}><w:r><w:delText xml:space="preserve"> tail</w:delText></w:r></w:del></w:p>`,
    // The same after a character outside the BMP, which is two UTF-16 code units.
    '<w:p><w:r><w:t>Smile\u{1F600}</w:t></w:r>'
      + `<w:del w:id="14" ${REVISION}><w:r><w:delText xml:space="preserve"> gone</w:delText></w:r></w:del></w:p>`,
    '<w:p><w:r><w:t>Wave</w:t></w:r>'
      + `<w:del w:id="15" ${REVISION}><w:r><w:delText xml:space="preserve"> gone</w:delText></w:r></w:del>`
      + '<w:r><w:t>\u{1F600}!</w:t></w:r></w:p>',
    // Deletions beside content that takes no room in the text: a line break, a tab.
    '<w:p><w:r><w:t>Line</w:t><w:br/></w:r>'
      + `<w:del w:id="16" ${REVISION}><w:r><w:delText>XX</w:delText></w:r></w:del><w:r><w:t>Tail4</w:t></w:r></w:p>`,
    '<w:p><w:r><w:t>Tab</w:t><w:tab/></w:r>'
      + `<w:del w:id="17" ${REVISION}><w:r><w:delText>XX</w:delText></w:r></w:del><w:r><w:t>Tail6</w:t></w:r></w:p>`,
    '<w:p><w:r><w:t>Ahead</w:t></w:r>'
      + `<w:del w:id="18" ${REVISION}><w:r><w:delText>XX</w:delText></w:r></w:del><w:r><w:tab/><w:t>Tail7</w:t></w:r></w:p>`,
    // The same nested move source written as Docxodus's own move marking writes it (w:delText).
    '<w:p><w:r><w:t xml:space="preserve">Read </w:t></w:r><w:hyperlink w:anchor="notes">'
      + `<w:moveFrom w:id="19" ${REVISION}><w:r><w:delText xml:space="preserve">the notes </w:delText></w:r></w:moveFrom>`
      + '<w:r><w:t>here</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> today.</w:t></w:r></w:p>',
  ].join('');
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

async function openTracked(page: Page, fixture = deletionFixture(), options: Record<string, unknown> = {}, tracked = true) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ bytes, options, tracked }) => {
    const w = window as any;
    const D = w.Docxodus;
    const container = document.createElement('div');
    document.body.appendChild(container);
    w.failures = [] as unknown[];
    w.editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, {
      ...(tracked ? { trackedChanges: 1 } : {}), // TrackedChangeMode.RenderInline, or the default
      onEditFailed: (info: unknown) => w.failures.push(info),
      ...options,
    });
    w.container = container;
    w.block = (text: string) => Array.from(container.querySelectorAll<HTMLElement>('p[data-anchor]'))
      .find((p) => (p.textContent ?? '').includes(text));
    // The paragraphs as a reader sees them once every tracked change is accepted.
    w.acceptedText = () => {
      const view = document.createElement('div');
      document.body.appendChild(view);
      const reopened = D.DocxEditor.open(view, w.editor.save(), D, {});
      reopened.acceptAllRevisions();
      const texts = Array.from(view.querySelectorAll('p[data-anchor]'))
        .map((p) => (p.textContent ?? '').replace(/\u00a0/g, ' '));
      reopened.close();
      view.remove();
      return texts;
    };
  }, { bytes: Array.from(fixture), options, tracked });
}

/** Put the caret in the block holding `text`, just after the first text node occurrence of `at`. */
async function caretAfter(page: Page, text: string, at: string) {
  await page.evaluate(({ text, at }) => {
    const w = window as any;
    const block: HTMLElement = w.block(text);
    block.focus();
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const index = (node.textContent ?? '').indexOf(at);
      if (index < 0) continue;
      const range = document.createRange();
      range.setStart(node, index + at.length);
      range.collapse(true);
      const selection = window.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      return;
    }
    throw new Error(`no text node holds "${at}"`);
  }, { text, at });
}

async function blur(page: Page) {
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
}

/** Put the caret at the start or end of the first text node in the block holding `text` that contains `node`. */
async function caretAtEdgeOf(page: Page, text: string, node: string, side: 'start' | 'end') {
  await page.evaluate(({ text, at, side }) => {
    const block: HTMLElement = (window as any).block(text);
    block.focus();
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (!(n.textContent ?? '').includes(at)) continue;
      const range = document.createRange();
      range.setStart(n, side === 'start' ? 0 : (n.textContent ?? '').length);
      range.collapse(true);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
      return;
    }
    throw new Error(`no text node holds "${at}"`);
  }, { text, at: node, side });
}

/** The deleted text, and deleted-mark pilcrows, the block holding `text` shows. */
async function deletedTextIn(page: Page, text: string): Promise<string[]> {
  return page.evaluate((t) => Array.from((window as any).block(t)
    ?.querySelectorAll('del, [class*="para-mark-del"]') ?? [])
    .map((d: any) => (d.textContent ?? '').replace(/\u00a0/g, ' ')), text);
}

test.describe('DocxEditor — offsets to the right of a rendered deletion (#1278)', () => {
  test('the committed baseline is the session text, without deleted text or a deleted mark', async ({ page }) => {
    await openTracked(page);
    const blocks = await page.evaluate(() => {
      const w = window as any;
      return ['The tenant', 'See', 'Joined', 'Plain control', 'Read'].map((t) => {
        const block: HTMLElement = w.block(t);
        return {
          editable: block.getAttribute('contenteditable'),
          deleted: Array.from(block.querySelectorAll('del, [class*="para-mark-del"]'))
            .map((d) => (d.textContent ?? '').replace(/\u00a0/g, ' ')),
          // The converter shows a space before a run boundary as an NBSP; the session holds a space.
          baseline: block.dataset.committedText?.replace(/\u00a0/g, ' '),
        };
      });
    });
    // The fixture renders as intended: editable paragraphs that show the deleted text.
    expect(blocks.map((b) => b.editable)).toEqual(['true', 'true', 'true', 'true', 'true']);
    expect(blocks.map((b) => b.deleted)).toEqual([['promptly '], ['the terms '], ['¶'], [], ['the notes ']]);
    expect(blocks.map((b) => b.baseline)).toEqual([
      'The tenant shall repaint the fence each spring.',
      'See below now.',
      'Joined to the next.',
      'Plain control paragraph.',
      'Read here today.',
    ]);
  });

  test('typing at the end of a paragraph with a deletion is recorded', async ({ page }) => {
    await openTracked(page);
    await page.evaluate(() => (window as any).block('The tenant').focus());
    await page.keyboard.press('End');
    await page.keyboard.type(' Z');
    await blur(page);
    // Again, in the block the commit re-rendered, whose baseline was measured afresh.
    await page.evaluate(() => (window as any).block('The tenant').focus());
    await page.keyboard.press('End');
    await page.keyboard.type('Z');
    await blur(page);
    const out = await page.evaluate(() => {
      const w = window as any;
      const shown = (w.block('The tenant')?.textContent ?? '').replace(/\u00a0/g, ' ');
      return { failures: w.failures, shown, accepted: w.acceptedText() };
    });
    expect(out.failures).toEqual([]);
    expect(out.shown).toBe('The tenant shall promptly repaint the fence each spring. ZZ');
    expect(out.accepted[0]).toBe('The tenant shall repaint the fence each spring. ZZ');
  });

  test('typing between a deletion and an insertion lands where the caret was', async ({ page }) => {
    await openTracked(page);
    await caretAfter(page, 'The tenant', 'repaint');
    await page.keyboard.type('ed');
    await blur(page);
    const out = await page.evaluate(() => ({
      failures: (window as any).failures,
      accepted: (window as any).acceptedText(),
    }));
    expect(out.failures).toEqual([]);
    expect(out.accepted[0]).toBe('The tenant shall repainted the fence each spring.');
  });

  // Word writes a move source's text as w:t, Docxodus's own move marking as w:delText; both render
  // as the same <del>, so the offsets must not depend on which (issue #1282).
  for (const [block, after, accepted] of [['See', 'below', 'See below! now.'], ['Read', 'here', 'Read here! today.']]) {
    test(`typing after a move source nested in a hyperlink ("${block}") lands where the caret was`, async ({ page }) => {
      await openTracked(page);
      await caretAfter(page, block, after);
      await page.keyboard.type('!');
      await blur(page);
      const out = await page.evaluate(() => ({
        failures: (window as any).failures,
        accepted: (window as any).acceptedText(),
      }));
      expect(out.failures).toEqual([]);
      expect(out.accepted).toContain(accepted);
    });
  }

  test('typing at the end of a paragraph whose mark is deleted is recorded', async ({ page }) => {
    await openTracked(page);
    await page.evaluate(() => (window as any).block('Joined').focus());
    await page.keyboard.press('End');
    await page.keyboard.type('!');
    await blur(page);
    const out = await page.evaluate(() => ({
      failures: (window as any).failures,
      accepted: (window as any).acceptedText(),
    }));
    expect(out.failures).toEqual([]);
    expect(out.accepted).toContain('Joined to the next.!Plain control paragraph.');
  });

  // The deletion is not editable, so the browser would remove it whole, although the session still
  // holds it. As in Word, the caret skips the deleted text and removes the live character beyond it.
  const besideDeletion = [
    { key: 'Backspace', block: 'The tenant', caret: { node: 'repaint', side: 'start' }, shown: 'promptly ',
      accepted: 'The tenant shallrepaint the fence each spring.' },
    { key: 'Delete', block: 'The tenant', caret: { node: 'shall', side: 'end' }, shown: 'promptly ',
      accepted: 'The tenant shall epaint the fence each spring.' },
    { key: 'Backspace', block: 'Keep thi', caret: 'End', shown: ' tail',
      accepted: 'Keep thi' },
    { key: 'Backspace', block: 'Smile', caret: 'End', shown: ' gone', accepted: 'Smile' },
    { key: 'Delete', block: 'Wave', caret: { node: 'Wave', side: 'end' }, shown: ' gone', accepted: 'Wave!' },
    // A move source nested in a link, in each form, is stepped over the same way (#1282).
    { key: 'Backspace', block: 'See', caret: { node: 'below', side: 'start' }, shown: 'the terms ',
      accepted: 'Seebelow now.' },
    { key: 'Delete', block: 'Read', caret: { node: 'Read', side: 'end' }, shown: 'the notes ',
      accepted: 'Read ere today.' },
    // A deleted paragraph mark: End puts the caret after the pilcrow.
    { key: 'Backspace', block: 'Joined', caret: 'End', shown: '¶',
      accepted: 'Joined to the nextPlain control paragraph.' },
  ] as const;
  for (const c of besideDeletion) {
    test(`${c.key} beside the deletion in "${c.block}" removes the live character beyond it`, async ({ page }) => {
      await openTracked(page);
      if (c.caret === 'End') {
        await page.evaluate((t) => (window as any).block(t).focus(), c.block);
        await page.keyboard.press('End');
      } else {
        await caretAtEdgeOf(page, c.block, c.caret.node, c.caret.side);
      }
      await page.keyboard.press(c.key);
      const shownBeforeCommit = await deletedTextIn(page, c.block);
      await blur(page);
      const out = await page.evaluate(() => ({
        failures: (window as any).failures,
        accepted: (window as any).acceptedText(),
      }));
      expect(shownBeforeCommit).toContain(c.shown);
      expect(await deletedTextIn(page, c.block)).toContain(c.shown);
      expect(out.failures).toEqual([]);
      expect(out.accepted).toContain(c.accepted);
    });
  }

  // Where the step would also cross something the text does not count, or there is no live
  // character beyond the deletion, the key does nothing rather than remove the wrong character.
  const refusedBeside = [
    { key: 'Backspace', block: 'Tail4', caret: { node: 'Tail4', side: 'start' }, why: 'a line break before the deletion' },
    { key: 'Backspace', block: 'Tail6', caret: { node: 'Tail6', side: 'start' }, why: 'a tab before the deletion' },
    { key: 'Delete', block: 'Tail7', caret: { node: 'Ahead', side: 'end' }, why: 'a tab after the deletion' },
    { key: 'Delete', block: 'Keep this', caret: { node: 'Keep this', side: 'end' }, why: 'the paragraph end after the deletion' },
  ] as const;
  for (const c of refusedBeside) {
    test(`${c.key} beside a deletion with ${c.why} changes nothing`, async ({ page }) => {
      await openTracked(page);
      const state = () => page.evaluate((t) => ({
        text: (window as any).block(t)?.textContent,
        revisions: (window as any).editor.listRevisions().map((r: any) => `${r.type}:${r.text}`),
      }), c.block);
      const before = await state();
      await caretAtEdgeOf(page, c.block, c.caret.node, c.caret.side);
      await page.keyboard.press(c.key);
      const shownBeforeCommit = (await state()).text;
      await blur(page);
      expect(shownBeforeCommit).toBe(before.text);
      expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
      expect(await state()).toEqual(before);
    });
  }

  test('Backspace after a deletion is recorded as a tracked deletion of the live character', async ({ page }) => {
    await openTracked(page);
    await page.evaluate(() => (window as any).block('Keep this').focus());
    await page.keyboard.press('End');
    await page.keyboard.press('Backspace');
    await blur(page);
    const revisions = await page.evaluate(() => (window as any).editor.listRevisions()
      .map((r: any) => `${r.type}:${r.text}`));
    expect(revisions).toContain('delete:s');
    expect(revisions).toContain('delete: tail');
  });

  // A paragraph whose only text is deleted has an empty baseline. Rewriting it whole is refused in
  // tracked mode when it holds a bookmark, so the typed text must go in as a span insertion.
  for (const [block, key] of [['Gone heading', 'Home'], ['Gone heading', 'End'], ['Gone too', 'Home'], ['Gone too', 'End']] as const) {
    test(`typing after ${key} in a paragraph whose only text ("${block}") is deleted is recorded`, async ({ page }) => {
      await openTracked(page);
      await page.evaluate((t) => (window as any).block(t).focus(), block);
      await page.keyboard.press(key);
      await page.keyboard.type('X');
      await blur(page);
      const out = await page.evaluate((t) => ({
        failures: (window as any).failures,
        shown: ((window as any).block(t)?.textContent ?? '').replace(/\u00a0/g, ' '),
        accepted: (window as any).acceptedText(),
      }), block);
      expect(out.failures).toEqual([]);
      expect(out.shown).toContain('X');
      expect(out.shown).toContain(block);
      expect(out.accepted).toContain('X');
    });
  }

  test('a line break typed into a paragraph whose only text is deleted is kept', async ({ page }) => {
    await openTracked(page);
    await page.evaluate(() => (window as any).block('Gone too').focus());
    await page.keyboard.press('Home');
    await page.keyboard.type('L1');
    await page.keyboard.press('Shift+Enter');
    await page.keyboard.type('L2');
    await blur(page);
    const out = await page.evaluate(() => {
      const block: HTMLElement = (window as any).block('Gone too');
      return { failures: (window as any).failures, breaks: block.querySelectorAll('br').length, text: block.textContent };
    });
    expect(out.failures).toEqual([]);
    expect(out.breaks).toBe(1);
    expect(out.text).toContain('L1');
    expect(out.text).toContain('L2');
  });

  test('control: typing to the left of a deletion is recorded', async ({ page }) => {
    await openTracked(page);
    await page.evaluate(() => (window as any).block('The tenant').focus());
    await page.keyboard.press('Home');
    await page.keyboard.type('Q');
    await blur(page);
    const out = await page.evaluate(() => ({
      failures: (window as any).failures,
      accepted: (window as any).acceptedText(),
    }));
    expect(out.failures).toEqual([]);
    expect(out.accepted[0]).toBe('QThe tenant shall repaint the fence each spring.');
  });
});

// The page view marks a PAGE field in the body. One deleted with its text — whole inside one
// w:del, or run by run as Word often writes it, even with differing dates — is deleted text too, so the offsets after it must
// not count its cached result (issue #1272).
function deletedPageFieldFixture(): Uint8Array {
  const field = (wrap: (run: string, index: number) => string) => [
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>',
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    '<w:r><w:delText>1</w:delText></w:r>',
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
  ].map(wrap).join('');
  const body = '<w:p><w:r><w:t xml:space="preserve">Sheet </w:t></w:r>'
    + `<w:del w:id="1" ${REVISION}>${field((run) => run)}</w:del>`
    + '<w:r><w:t xml:space="preserve"> follows</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t xml:space="preserve">Leaf </w:t></w:r>'
    + field((run) => `<w:del w:id="2" ${REVISION}>${run}</w:del>`)
    + '<w:r><w:t xml:space="preserve"> next</w:t></w:r></w:p>'
    // Run by run again, each deletion with its own id, the result's a day later.
    + '<w:p><w:r><w:t xml:space="preserve">Mark </w:t></w:r>'
    + field((run, i) => `<w:del w:id="${10 + i}" w:author="Counsel A" w:date="2026-05-0${i >= 3 ? 2 : 1}T09:00:00Z">${run}</w:del>`)
    + '<w:r><w:t xml:space="preserve"> last</w:t></w:r></w:p>';
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

test.describe('DocxEditor — offsets after a deleted PAGE field in the page view (#1272)', () => {
  for (const [block, tail] of [['Sheet', 'follows'], ['Leaf', 'next'], ['Mark', 'last']] as const) {
    test(`typing after the deleted field in "${block}" lands where the caret was`, async ({ page }) => {
      await openTracked(page, deletedPageFieldFixture(), { paginated: true });
      // The fixture renders as intended: the field is marked, inside the deletion.
      expect(await page.evaluate((b) => (window as any).block(b)
        .querySelector('del [data-field="PAGE"]')?.textContent, block)).toBe('1');
      await caretAfter(page, block, tail);
      await page.keyboard.type('!');
      await blur(page);
      const out = await page.evaluate(() => ({
        failures: (window as any).failures,
        accepted: (window as any).acceptedText(),
      }));
      expect(out.failures).toEqual([]);
      expect(out.accepted).toContain(`${block}  ${tail}!`);
    });
  }
});

test.describe('DocxEditor — typing inside a tracked move destination (#1319)', () => {
  // The engine refuses text strictly inside a w:moveTo (its resolvers disagree about text in a move
  // range, #1340), on the zero-length path and on the editor's replacement fallback alike. The
  // editor must report the refusal once and put the block back as the session holds it.
  for (const tracked of [true, false]) {
    test(`typing inside moved text is refused cleanly (${tracked ? 'tracked' : 'untracked'})`, async ({ page }) => {
      await openTracked(page, deletionFixture(), {}, tracked);
      const before = await page.evaluate(() => (window as any).acceptedText());
      await caretAfter(page, 'Moved here', 'the');
      await page.keyboard.type('Z');
      await blur(page);
      const out = await page.evaluate(() => ({
        failures: (window as any).failures.map((f: any) => f.code),
        shown: ((window as any).block('Moved here')?.textContent ?? '').replace(/ /g, ' '),
        accepted: (window as any).acceptedText(),
      }));
      expect(out.failures).toEqual(['unsupported_inline_boundary']);
      expect(out.shown).not.toContain('Z');
      expect(out.accepted).toEqual(before);
    });
  }
});

test.describe('DocxEditor — typing inside a tracked insertion (#1318)', () => {
  // The engine used to refuse an offset inside another author's insertion, and the editor's
  // fallback then struck the neighbouring character and retyped it with the keystroke. Word
  // splits the insertion around the typist's own instead.
  test('typing inside another author\'s insertion splits it without retyping a character', async ({ page }) => {
    await openTracked(page);
    const deletions = () => page.evaluate(() => (window as any).editor.listRevisions()
      .map((r: any) => `${r.type}:${r.text}`).filter((r: string) => r.startsWith('delete:')));
    const before = await deletions();
    await caretAfter(page, 'The tenant', 'each');
    await page.keyboard.type('Z');
    await blur(page);
    const out = await page.evaluate(() => ({
      failures: (window as any).failures,
      revisions: (window as any).editor.listRevisions().map((r: any) => `${r.type}:${r.text}`),
      accepted: (window as any).acceptedText(),
    }));
    expect(out.failures).toEqual([]);
    expect(await deletions()).toEqual(before);
    expect(out.revisions).toEqual(expect.arrayContaining(['insert: each', 'insert:Z', 'insert: spring']));
    expect(out.accepted[0]).toBe('The tenant shall repaint the fence eachZ spring.');
  });
});

test.describe('DocxEditor — typing beside a deletion stays on the caret side (#1290)', () => {
  // Deleted text takes no room in the session's offsets, so the offset alone names both sides of
  // it; the editor tells the session which side the text was typed on.
  test('typing right after a deletion is shown after it', async ({ page }) => {
    await openTracked(page);
    await caretAtEdgeOf(page, 'The tenant', 'repaint', 'start');
    await page.keyboard.type('Y');
    await blur(page);
    const out = await page.evaluate(() => ({
      failures: (window as any).failures,
      shown: ((window as any).block('The tenant')?.textContent ?? '').replace(/ /g, ' '),
      accepted: (window as any).acceptedText(),
    }));
    expect(out.failures).toEqual([]);
    expect(out.shown).toBe('The tenant shall promptly Yrepaint the fence each spring.');
    expect(out.accepted[0]).toBe('The tenant shall Yrepaint the fence each spring.');
  });

  for (const [key, shown] of [['Home', 'XGone heading'], ['End', 'Gone headingX']] as const) {
    test(`typing after ${key} in a deleted heading stays on that side, inside its bookmark`, async ({ page }) => {
      await openTracked(page);
      await page.evaluate(() => (window as any).block('Gone heading').focus());
      await page.keyboard.press(key);
      await page.keyboard.type('X');
      await blur(page);
      const out = await page.evaluate(() => {
        const w = window as any;
        const block: HTMLElement = w.block('Gone heading');
        const xml: string = w.Docxodus.DocxSessionBridge.RawGetXml(
          w.editor.handle, 'p:body:' + block.getAttribute('data-anchor'));
        return { failures: w.failures, shown: (block.textContent ?? '').replace(/ /g, ' '), xml };
      });
      expect(out.failures).toEqual([]);
      expect(out.shown).toBe(shown);
      const at = (needle: string) => out.xml.indexOf(needle);
      expect(at('<w:bookmarkStart')).toBeLessThan(at('<w:ins'));
      expect(at('</w:ins>')).toBeLessThan(at('<w:bookmarkEnd'));
      expect(key === 'Home' ? at('</w:ins>') < at('<w:del ') : at('</w:del>') < at('<w:ins')).toBe(true);
    });
  }
});
