import { test, expect, Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Text in a symbol font renders as its Unicode equivalent (#1707), but the session's text is the
// private-use character the run stores. The editor counts and commits the stored character, so an
// edit beside a mapped glyph lands at the session's offset and leaves the glyph as it was. The
// pencil, U+F021 in Wingdings, maps to U+1F589, two UTF-16 units for one stored unit.

const WINGDINGS = '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr>';

function fixture(glyph: string): Uint8Array {
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
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>`
      + `<w:p><w:r><w:t xml:space="preserve">Pen </w:t></w:r><w:r>${WINGDINGS}<w:t>${glyph}</w:t></w:r>`
      + '<w:r><w:t xml:space="preserve"> Box</w:t></w:r></w:p>'
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
      + '</w:body></w:document>') },
  ]);
}

async function openEditor(page: Page, glyph: string) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate((bytes) => {
    const w = window as any;
    const container = document.createElement('div');
    document.body.appendChild(container);
    w.failures = [] as unknown[];
    w.editor = w.Docxodus.DocxEditor.open(container, new Uint8Array(bytes), w.Docxodus, {
      onEditFailed: (info: unknown) => w.failures.push(info),
    });
  }, Array.from(fixture(glyph)));
}

/** The saved paragraph's run text, one entry per run. */
async function savedRuns(page: Page): Promise<string[]> {
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  return page.evaluate((ns) => {
    const session = (window as any).Docxodus.openTypedSession((window as any).editor.save());
    try {
      const [id] = Object.entries(session.project().anchorIndex).find(([, a]: any) => a.kind === 'p')!;
      const doc = new DOMParser().parseFromString(session.raw.getXml(id), 'application/xml');
      return Array.from(doc.getElementsByTagNameNS(ns, 'r')).map((r) => r.textContent ?? '');
    } finally { session.close(); }
  }, W_NS);
}

for (const [glyph, shown] of [['\uF021', '\u{1F589}'], ['\uF0FE', '☑']]) {
  test(`typing after a mapped symbol glyph keeps the stored character: U+${glyph.codePointAt(0)!.toString(16).toUpperCase()}`, async ({ page }) => {
    await openEditor(page, glyph);
    const block = page.locator('p[data-anchor][contenteditable="true"]').filter({ hasText: 'Box' }).first();
    await expect(block.locator('[data-docx-symbol-text]')).toHaveText(shown);
    await page.evaluate(() => {
      const block = Array.from(document.querySelectorAll<HTMLElement>('p[data-anchor][contenteditable="true"]'))
        .find((p) => (p.textContent ?? '').includes('Box'))!;
      block.focus();
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      let last: Text | null = null;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) if ((n.textContent ?? '').includes('Box')) last = n as Text;
      window.getSelection()!.collapse(last!, last!.data.length);
    });
    await page.keyboard.type('X');
    expect(await savedRuns(page)).toEqual(['Pen ', glyph, ' BoxX']);
    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
  });
}
