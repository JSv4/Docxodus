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
    // direct child of the paragraph, so the paragraph stays editable, and the session reads its
    // text, so the editor must keep counting it.
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

async function openTracked(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((bytes) => {
    const w = window as any;
    const D = w.Docxodus;
    const container = document.createElement('div');
    document.body.appendChild(container);
    w.failures = [] as unknown[];
    w.editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, {
      trackedChanges: 1, // TrackedChangeMode.RenderInline
      onEditFailed: (info: unknown) => w.failures.push(info),
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
  }, Array.from(deletionFixture()));
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
      return ['The tenant', 'See', 'Joined', 'Plain control'].map((t) => {
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
    expect(blocks.map((b) => b.editable)).toEqual(['true', 'true', 'true', 'true']);
    expect(blocks.map((b) => b.deleted)).toEqual([['promptly '], ['the terms '], ['¶'], []]);
    expect(blocks.map((b) => b.baseline)).toEqual([
      'The tenant shall repaint the fence each spring.',
      'See the terms below now.',
      'Joined to the next.',
      'Plain control paragraph.',
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

  test('control: typing after a nested move source, whose text the session reads, is recorded', async ({ page }) => {
    await openTracked(page);
    await caretAfter(page, 'See', 'below');
    await page.keyboard.type('!');
    await blur(page);
    const out = await page.evaluate(() => ({
      failures: (window as any).failures,
      accepted: (window as any).acceptedText(),
    }));
    expect(out.failures).toEqual([]);
    expect(out.accepted).toContain('See below! now.');
  });

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
