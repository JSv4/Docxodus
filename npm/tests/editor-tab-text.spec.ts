import { test, expect, Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import type { DocxSession } from '../src/session.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// A w:tab is one character, "\t", of the session's text (issue #1387): grep finds it, spans count
// it, and the editor, which renders it as an empty box, must count it as one too or every offset
// after a tab lands one short.

function tabFixture(): Uint8Array {
  const body = [
    '<w:p><w:r><w:t>1.2</w:t><w:tab/><w:t>Alpha</w:t></w:r></w:p>',
    '<w:p><w:r><w:tab/><w:t>Lead</w:t></w:r></w:p>',
    '<w:p><w:r><w:t>Name:</w:t><w:tab/><w:t>[_____]</w:t><w:tab/><w:t>Date</w:t></w:r></w:p>',
    '<w:p><w:r><w:t>Plain control text</w:t></w:r></w:p>',
  ].join('');
  const relationship = (id: string, kind: string, target: string) =>
    `<Relationship Id="${id}" Type="${R_NS}/${kind}" Target="${target}"/>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>') },
    { name: '_rels/.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + relationship('rId1', 'officeDocument', 'word/document.xml') + '</Relationships>') },
    { name: 'word/_rels/document.xml.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + relationship('rId1', 'styles', 'styles.xml') + relationship('rId2', 'settings', 'settings.xml') + '</Relationships>') },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"/>`) },
    { name: 'word/settings.xml', data: xml(`<w:settings xmlns:w="${W_NS}"/>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>${body}`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
      + '</w:body></w:document>') },
  ]);
}

async function ready(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** Each paragraph's run content in a saved package: text as itself, a tab as `|`. */
const SHAPES = `(bytes) => {
  const session = window.Docxodus.openTypedSession(bytes);
  try {
    return Object.entries(session.project().anchorIndex).filter(([, a]) => a.kind === 'p').map(([id]) => {
      const doc = new DOMParser().parseFromString(session.raw.getXml(id), 'application/xml');
      return Array.from(doc.getElementsByTagNameNS('${W_NS}', 'r')).flatMap((r) => Array.from(r.children))
        .map((c) => c.localName === 't' ? c.textContent : c.localName === 'tab' ? '|' : '').join('');
    });
  } finally { session.close(); }
}`;

test.describe('a tab is one character of the session text (#1387)', () => {
  test('grep, a replacement across a tab, and coalescing beside one', async ({ page }) => {
    await ready(page);
    const out = await page.evaluate(({ bytes, shapes }) => {
      const D = (window as any).Docxodus;
      const shapesOf = eval(shapes) as (b: Uint8Array) => string[];
      const session = D.openTypedSession(new Uint8Array(bytes)) as DocxSession;
      try {
        const tabs = session.grep('\t').map((m) => [m.span.start, m.span.length]);
        const across = session.grep('2\tAlpha');
        const replaced = session.replaceMatch(across[0], '2\tBeta').success;
        const fill = session.fillPlaceholders(() => '', { coalesceWhitespaceAroundEmptyFill: true });
        return { tabs, across: across.length, replaced, filled: fill.filled, shapes: shapesOf(session.save()) };
      } finally { session.close(); }
    }, { bytes: Array.from(tabFixture()), shapes: SHAPES });
    expect(out.tabs).toEqual([[3, 1], [0, 1], [5, 1], [13, 1]]);
    expect(out.across).toBe(1);
    expect(out.replaced).toBe(true);
    expect(out.filled).toBe(1);
    // One tab, before "Beta"; both tabs either side of the emptied placeholder stay.
    expect(out.shapes).toEqual(['1.2|Beta', '|Lead', 'Name:||Date', 'Plain control text']);
  });

  async function openEditor(page: Page) {
    await ready(page);
    await page.evaluate((bytes) => {
      const w = window as any;
      const container = document.createElement('div');
      document.body.appendChild(container);
      w.failures = [] as unknown[];
      w.editor = w.Docxodus.DocxEditor.open(container, new Uint8Array(bytes), w.Docxodus, {
        onEditFailed: (info: unknown) => w.failures.push(info),
      });
      w.block = (text: string) => Array.from(container.querySelectorAll<HTMLElement>('p[data-anchor]'))
        .find((p) => (p.textContent ?? '').includes(text));
    }, Array.from(tabFixture()));
  }

  /** Put the caret in the block holding `text`, `at` characters into its text node holding `node`. */
  async function caretIn(page: Page, text: string, node: string, at: number) {
    await page.evaluate(({ text, node, at }) => {
      const block: HTMLElement = (window as any).block(text);
      block.focus();
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if ((n.textContent ?? '') !== node) continue;
        const range = document.createRange();
        range.setStart(n, at);
        range.collapse(true);
        window.getSelection()!.removeAllRanges();
        window.getSelection()!.addRange(range);
        return;
      }
      throw new Error(`no text node is "${node}"`);
    }, { text, node, at });
  }

  async function saved(page: Page): Promise<string[]> {
    await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
    return page.evaluate((shapes) => (eval(shapes) as (b: Uint8Array) => string[])(
      (window as any).editor.save()), SHAPES);
  }

  test('the committed baseline holds the tab', async ({ page }) => {
    await openEditor(page);
    const baselines = await page.evaluate(() => ['Alpha', 'Lead', 'Plain'].map((t) =>
      (window as any).block(t).dataset.committedText));
    expect(baselines).toEqual(['1.2\tAlpha', '\tLead', 'Plain control text']);
  });

  test('typing just after a tab lands after it, and just before one before it', async ({ page }) => {
    await openEditor(page);
    await caretIn(page, 'Alpha', 'Alpha', 0);
    await page.keyboard.type('X');
    await caretIn(page, 'Alpha', '1.2', 3);
    await page.keyboard.type('Y');
    expect((await saved(page))[0]).toBe('1.2Y|XAlpha');
    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
  });

  test('typing and Enter in a paragraph that starts with a tab', async ({ page }) => {
    await openEditor(page);
    await caretIn(page, 'Lead', 'Lead', 2);
    await page.keyboard.type('Z');
    await page.keyboard.press('Enter');
    const shapes = await saved(page);
    expect(shapes.slice(1, 3)).toEqual(['|LeZ', 'ad']);
    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
  });

  for (const [key, node, at] of [['Backspace', 'Alpha', 0], ['Delete', '1.2', 3]] as const) {
    test(`${key} beside a tab removes the tab from the document`, async ({ page }) => {
      await openEditor(page);
      await caretIn(page, 'Alpha', node, at);
      await page.keyboard.press(key);
      await page.keyboard.type('Q');
      expect((await saved(page))[0]).toBe('1.2QAlpha');
      expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
    });
  }

  // The editor counts a rendered tab box as one character; that holds only while the converter
  // renders exactly one box per run tab, and none for numbering's own tab (inside the marker).
  for (const name of ['HC024-Tabs-01.docx', 'HC025-Tabs-02.docx', 'HC026-Tabs-03.docx', 'HC027-Tabs-04.docx', 'HW010-Tab01.docx']) {
    test(`every block's baseline holds as many tabs as the session text: ${name}`, async ({ page }) => {
      await ready(page);
      const bytes = Array.from(readFileSync(new URL(`../../TestFiles/${name}`, import.meta.url)));
      const out = await page.evaluate((bytes) => {
        const w = window as any;
        const container = document.createElement('div');
        document.body.appendChild(container);
        const editor = w.Docxodus.DocxEditor.open(container, new Uint8Array(bytes), w.Docxodus, {});
        const tabs = (s: string) => s.split('\t').length - 1;
        let checked = 0;
        let withTabs = 0;
        const mismatches: string[] = [];
        for (const block of Array.from(container.querySelectorAll<HTMLElement>('[data-anchor]'))) {
          const committed = block.dataset.committedText;
          const id = editor.unidToFullId.get(block.getAttribute('data-anchor'));
          if (committed === undefined || !id) continue;
          const info = JSON.parse(editor.exports.DocxSessionBridge.GetAnchorInfo(editor.handle, id));
          const session = info.visibleText ?? '';
          checked++;
          if (tabs(session) > 0) withTabs++;
          if (tabs(committed) !== tabs(session)) mismatches.push(`${id}: ${JSON.stringify(committed)} vs ${JSON.stringify(session)}`);
        }
        return { checked, withTabs, mismatches };
      }, bytes);
      expect(out.checked).toBeGreaterThan(0);
      expect(out.withTabs).toBeGreaterThan(0);
      expect(out.mismatches).toEqual([]);
    });
  }

  test('control: typing in a paragraph without a tab', async ({ page }) => {
    await openEditor(page);
    await caretIn(page, 'Plain', 'Plain control text', 5);
    await page.keyboard.type('!');
    expect((await saved(page))[3]).toBe('Plain! control text');
  });
});
