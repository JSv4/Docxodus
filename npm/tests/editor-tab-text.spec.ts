import { test, expect, Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import type { DocxSession } from '../src/session.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// A w:tab is one character, "\t", of the session's text (issue #1387): grep finds it, spans count
// it, and the editor, which renders it as an empty box, must count it as one too or every offset
// after a tab lands one short.

function tabFixture(bodyOverride?: string): Uint8Array {
  const body = bodyOverride ?? [
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

  async function openEditor(page: Page, tracked = false, body?: string, options: Record<string, unknown> = {}) {
    await ready(page);
    await page.evaluate(({ bytes, tracked, options }) => {
      const w = window as any;
      const container = document.createElement('div');
      document.body.appendChild(container);
      w.failures = [] as unknown[];
      w.editor = w.Docxodus.DocxEditor.open(container, new Uint8Array(bytes), w.Docxodus, {
        ...options,
        ...(tracked ? { trackedChanges: 1 } : {}), // TrackedChangeMode.RenderInline
        onEditFailed: (info: unknown) => w.failures.push(info),
      });
      w.block = (text: string) => Array.from(w.editor.root.querySelectorAll('p[data-anchor]') as NodeListOf<HTMLElement>)
        .find((p) => !p.closest('#pagination-staging') && (p.textContent ?? '').includes(text));
    }, { bytes: Array.from(tabFixture(body)), tracked, options });
  }

  /** Put the caret in the block holding `text`, `at` characters into its text node holding `node`. */
  async function caretIn(page: Page, text: string, node: string, at: number) {
    await page.evaluate(({ text, node, at }) => {
      const block: HTMLElement = (window as any).block(text);
      block.focus();
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const start = (n.textContent ?? '').indexOf(node);
        if (start < 0) continue;
        const range = document.createRange();
        range.setStart(n, start + at);
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

  for (const [key, node, at, expected] of [
    ['ArrowRight', '1.2', 3, '1.2|XAlpha'],
    ['ArrowLeft', 'Alpha', 0, '1.2X|Alpha'],
  ] as const) {
    test(`${key} crosses a tab in one step`, async ({ page }) => {
      await openEditor(page);
      await caretIn(page, 'Alpha', node, at);
      await page.keyboard.press(key);
      await page.keyboard.type('X');
      expect((await saved(page))[0]).toBe(expected);
    });
  }

  for (const [name, body, node, at, key, expected] of [
    ['adjacent tabs', '<w:p><w:r><w:t>A</w:t><w:tab/><w:tab/><w:t>Beta</w:t></w:r></w:p>', 'A', 1, 'ArrowRight', 'A|X|Beta'],
    ['decorated tab', '<w:p><w:r><w:rPr><w:u w:val="single"/><w:strike/></w:rPr><w:t>A</w:t><w:tab/><w:t>Beta</w:t></w:r></w:p>', 'A', 1, 'ArrowRight', 'A|XBeta'],
    ['trailing tab', '<w:p><w:r><w:t>Tail</w:t><w:tab/></w:r></w:p>', 'Tail', 4, 'ArrowRight', 'Tail|X'],
    ['emoji after tab', '<w:p><w:r><w:t>A</w:t><w:tab/><w:t>😀Beta</w:t></w:r></w:p>', '😀Beta', 0, 'ArrowRight', 'A|😀XBeta'],
    ['combining character before tab', '<w:p><w:r><w:t>é</w:t><w:tab/><w:t>Beta</w:t></w:r></w:p>', 'é', 2, 'ArrowLeft', 'Xé|Beta'],
    ['RTL tab', '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>אב</w:t><w:tab/><w:t>גד</w:t></w:r></w:p>', 'אב', 2, 'ArrowLeft', 'אב|Xגד'],
  ] as const) {
    test(`one arrow preserves native character movement: ${name}`, async ({ page }) => {
      await openEditor(page, false, body);
      await caretIn(page, node, node, at);
      await page.keyboard.press(key);
      await page.keyboard.type('X');
      expect((await saved(page))[0]).toBe(expected);
    });
  }

  test('ArrowRight crosses a tab-only paragraph without losing typed text', async ({ page }) => {
    await openEditor(page, false, '<w:p><w:r><w:tab/></w:r></w:p>');
    await page.evaluate(() => {
      const tab = document.querySelector('[data-anchor][contenteditable="true"] [data-docx-tab]')!;
      tab.closest<HTMLElement>('[contenteditable="true"]')!.focus();
      window.getSelection()!.collapse(tab.parentNode, Array.prototype.indexOf.call(tab.parentNode!.childNodes, tab));
    });
    await page.keyboard.press('ArrowRight');
    await page.keyboard.type('X');
    expect((await saved(page))[0]).toBe('|X');
  });

  test('Shift+ArrowRight selects the tab in one step', async ({ page }) => {
    await openEditor(page);
    await caretIn(page, 'Alpha', '1.2', 3);
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.type('X');
    expect((await saved(page))[0]).toBe('1.2XAlpha');
  });

  for (const mode of ['flow', 'paginated', 'adopted']) test(`native copy includes a selected tab, mode=${mode}`, async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openEditor(page);
    await page.evaluate(mode => {
      const editor = (window as any).editor;
      editor.setPaginated(mode === 'paginated');
      if (mode === 'adopted') {
        const host = document.createElement('div');
        document.body.append(host);
        editor.adoptContainer(host);
      }
    }, mode);
    await page.evaluate(() => {
      const block = (window as any).block('Alpha') as HTMLElement;
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    });
    await page.keyboard.press('Control+c');
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('1.2\tAlpha');
  });

  for (const text of ['Tail', '']) test(`repeated arrows beyond a trailing tab add no break: ${text}`, async ({ page }) => {
    await openEditor(page, false, `<w:p><w:r><w:t>${text}</w:t><w:tab/></w:r></w:p>`);
    await page.evaluate(() => {
      const block = document.querySelector<HTMLElement>('p[data-anchor][contenteditable="true"]')!;
      const tab = block.querySelector('[data-docx-tab]')!;
      block.focus();
      window.getSelection()!.collapse(tab.parentNode, Array.prototype.indexOf.call(tab.parentNode!.childNodes, tab));
    });
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.type('X');
    expect((await saved(page))[0]).toBe(`${text}|X`);
  });

  test('copy after crossing a terminal tab excludes the caret placeholder', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openEditor(page, false, '<w:p><w:r><w:t>Tail</w:t><w:tab/></w:r></w:p>');
    await caretIn(page, 'Tail', 'Tail', 4);
    await page.keyboard.press('ArrowRight');
    await page.evaluate(() => {
      const range = document.createRange();
      range.selectNodeContents((window as any).block('Tail'));
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    });
    await page.keyboard.press('Control+c');
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('Tail\t');
  });

  test('class-styled hidden runs remain hidden in copied HTML', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openEditor(page, false, '<w:p><w:r><w:t>1.2</w:t><w:tab/></w:r>'
      + '<w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t></w:r><w:r><w:t>Cedar</w:t></w:r></w:p>',
      { fabricateClasses: true });
    await page.evaluate(() => {
      const block = (window as any).block('Cedar') as HTMLElement;
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    });
    await page.keyboard.press('Control+c');
    const copied = await page.evaluate(async () => {
      const items = await navigator.clipboard.read();
      const html = await (await items[0].getType('text/html')).text();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const hidden = Array.from(doc.querySelectorAll<HTMLElement>('*')).find(el => el.textContent === 'Hidden');
      return { text: await navigator.clipboard.readText(), hidden: !!hidden?.closest('[style*="display: none"]') };
    });
    expect(copied).toEqual({ text: '1.2\tCedar', hidden: true });
  });

  for (const linked of [false, true]) test(`rich copy retains omitted ancestor context, linked=${linked}`, async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    const runs = '<w:r><w:t>אב</w:t><w:tab/><w:t>12-CD</w:t></w:r>';
    await openEditor(page, false, '<w:p><w:pPr><w:bidi/></w:pPr>'
      + (linked ? `<w:hyperlink w:anchor="CedarTarget">${runs}</w:hyperlink>` : runs) + '</w:p>');
    await page.evaluate(linked => {
      const block = document.querySelector<HTMLElement>('p[data-anchor][contenteditable="true"]')!;
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(linked ? block.querySelector('a[href]')! : block);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    }, linked);
    await page.keyboard.press('Control+c');
    const copied = await page.evaluate(async () => {
      const items = await navigator.clipboard.read();
      const html = await (await items[0].getType('text/html')).text();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      return { direction: doc.querySelector('div')?.style.direction, href: doc.querySelector('a[href]')?.getAttribute('href') };
    });
    expect(copied.direction).toBe('rtl');
    // Clipboard HTML may absolutize a fragment target against its own document base.
    if (linked) expect(copied.href?.split('#').at(-1)).toBe('CedarTarget');
    else expect(copied.href).toBeUndefined();
  });

  test('copying a partial span includes exactly its tab and text', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openEditor(page);
    await page.evaluate(() => {
      const block = (window as any).block('Alpha') as HTMLElement;
      block.focus();
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      const nodes: Text[] = [];
      for (let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n as Text);
      const range = document.createRange();
      range.setStart(nodes.find(n => n.data === '1.2')!, 2);
      range.setEnd(nodes.find(n => n.data === 'Alpha')!, 2);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    });
    await page.keyboard.press('Control+c');
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('2\tAl');
  });

  for (const [name, body, expected] of [
    ['tab only', '<w:p><w:r><w:tab/></w:r></w:p>', '\t'],
    ['class-styled text', '<w:p><w:r><w:rPr><w:b/><w:i/><w:spacing w:val="40"/></w:rPr><w:t>Bold</w:t><w:tab/><w:t>Italic</w:t></w:r></w:p>', 'Bold\tItalic'],
    ['rich text', '<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold</w:t><w:tab/><w:t>Italic</w:t></w:r></w:p>', 'Bold\tItalic'],
    ['plain paragraph control', '<w:p><w:r><w:t>One Two</w:t></w:r></w:p><w:p><w:r><w:t>Three</w:t></w:r></w:p>', 'One Two\n\nThree'],
    ['two paragraphs', '<w:p><w:r><w:t>One</w:t><w:tab/><w:t>Two</w:t></w:r></w:p><w:p><w:r><w:t>Three</w:t></w:r></w:p>', 'One\tTwo\n\nThree'],
  ] as const) {
    test(`copy retains plain and rich content: ${name}`, async ({ page, context }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      await openEditor(page, false, body, { fabricateClasses: name === 'class-styled text' });
      const before = await page.evaluate(() => {
        const root = (window as any).editor.root as HTMLElement;
        const blocks = root.querySelectorAll('p[data-anchor][contenteditable="true"]');
        (blocks[0] as HTMLElement).focus();
        const range = document.createRange();
        range.setStart(blocks[0], 0);
        range.setEnd(blocks[blocks.length - 1], blocks[blocks.length - 1].childNodes.length);
        window.getSelection()!.removeAllRanges();
        window.getSelection()!.addRange(range);
        return root.innerHTML;
      });
      await page.keyboard.press('Control+c');
      const copied = await page.evaluate(async () => {
        const items = await navigator.clipboard.read();
        const html = await (await items[0].getType('text/html')).text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        return { text: await navigator.clipboard.readText(), tab: doc.querySelector('[data-docx-tab]')?.textContent,
          bold: !!doc.querySelector('[style*="font-weight: bold"], [style*="font-weight: 700"]'), italic: !!doc.querySelector('[style*="font-style: italic"]'),
          spacing: Math.max(0, ...Array.from(doc.querySelectorAll<HTMLElement>('*')).map(el => parseFloat(el.style.letterSpacing) || 0)),
          source: ((window as any).editor.root as HTMLElement).innerHTML };
      });
      expect(copied.text).toBe(expected);
      expect(copied.tab).toBe(name === 'plain paragraph control' ? undefined : '\t');
      expect(copied.source).toBe(before);
      if (name === 'rich text' || name === 'class-styled text') expect([copied.bold, copied.italic]).toEqual([true, true]);
      if (name === 'class-styled text') expect(copied.spacing).toBeCloseTo(8 / 3, 4);
    });
  }

  for (const key of ['Backspace', 'Delete']) {
    test(`${key} stays conservative across a deletion beside a tab`, async ({ page }) => {
      const deletion = '<w:del w:id="1" w:author="Counsel" w:date="2026-05-01T09:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>';
      const left = '<w:r><w:t>A</w:t>' + (key === 'Backspace' ? '<w:tab/>' : '') + '</w:r>';
      const right = '<w:r>' + (key === 'Delete' ? '<w:tab/>' : '') + '<w:t>Beta</w:t></w:r>';
      await openEditor(page, true, `<w:p>${left}${deletion}${right}</w:p>`);
      await caretIn(page, 'Beta', key === 'Backspace' ? 'Beta' : 'A', key === 'Backspace' ? 0 : 1);
      const before = await saved(page);
      await caretIn(page, 'Beta', key === 'Backspace' ? 'Beta' : 'A', key === 'Backspace' ? 0 : 1);
      await page.keyboard.press(key);
      expect(await saved(page)).toEqual(before);
      expect(await page.evaluate(() => (window as any).editor.listRevisions().map((r: any) => r.text))).toEqual(['gone']);
    });
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

  // Chromium's own word deletion beside a tab box removes the word on the other side of the caret
  // (Ctrl+Delete before a tab deleted "1.2"); a tab is the whitespace between words, so it goes.
  for (const [key, node, at] of [['Control+Delete', '1.2', 3], ['Control+Backspace', 'Alpha', 0]] as const) {
    for (const tracked of [false, true]) {
      test(`${key} beside a tab removes only the tab (tracked: ${tracked})`, async ({ page }) => {
        await openEditor(page, tracked);
        await caretIn(page, 'Alpha', node, at);
        await page.keyboard.press(key);
        const shapes = await saved(page);
        const revisions = await page.evaluate(() => (window as any).editor.listRevisions()
          .map((r: any) => `${r.type}:${JSON.stringify(r.text)}`));
        // A tracked deletion keeps the tab, inside its w:del, until it is accepted.
        expect(shapes[0]).toBe(tracked ? '1.2|Alpha' : '1.2Alpha');
        expect(revisions).toEqual(tracked ? ['delete:"\\t"'] : []);
        expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
      });
    }
  }

  test('typing over a selection that spans a tab replaces the tab with it', async ({ page }) => {
    await openEditor(page);
    await page.evaluate(() => {
      const block: HTMLElement = (window as any).block('Alpha');
      block.focus();
      const texts: Text[] = [];
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) texts.push(n as Text);
      const range = document.createRange();
      range.setStart(texts.find((t) => t.data === '1.2')!, 2);
      range.setEnd(texts.find((t) => t.data === 'Alpha')!, 2);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
    });
    await page.keyboard.type('Q');
    expect((await saved(page))[0]).toBe('1.Qpha');
    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
  });

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
