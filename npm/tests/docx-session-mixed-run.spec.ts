import { test, expect } from '@playwright/test';
import type { DocxSession } from '../src/session.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

function mixedRunDocx(marker: 'tab' | 'br'): Uint8Array {
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
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr>`
      + `<w:t>ab</w:t><w:${marker}/><w:t>cd</w:t></w:r></w:p></w:body></w:document>`) },
  ]);
}

for (const marker of ['tab', 'br'] as const) for (const tracked of [false, true]) {
  test(`mixed run ${marker}, tracking ${tracked}: surgical edits retain marker order`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const results = await page.evaluate(({ bytes, tracked }) => {
      const open = (input: Uint8Array, record = false) => (window as any).Docxodus.openTypedSession(input,
        JSON.stringify({ trackedChanges: record ? 'render_inline' : 'accept', revisionAuthor: 'Editor' })) as DocxSession;
      const read = (session: DocxSession) => {
        const anchor = Object.entries(session.project().anchorIndex).find(([, a]) => a.kind === 'p')![0];
        const document = new DOMParser().parseFromString(session.raw.getXml(anchor), 'application/xml');
        return Array.from(document.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'r'))
          .flatMap(run => Array.from(run.children))
          .map(child => child.localName === 't' ? child.textContent : ['tab', 'br'].includes(child.localName) ? '|' : '').join('');
      };
      return [[1, 1], [2, 1], [1, 2]].map(([start, length]) => {
        const session = open(new Uint8Array(bytes), tracked);
        try {
          const match = session.grep('abcd'.slice(start, start + length))[0];
          const edited = session.replaceMatch(match, 'X');
          const saved = session.save();
          const accepted = open(saved), rejected = open(saved);
          try {
            const accept = accepted.acceptAllRevisions(), reject = rejected.rejectAllRevisions();
            const undo = session.undo();
            return { edited: edited.success, accept: accept.success, reject: reject.success,
              accepted: read(accepted), rejected: read(rejected), undo, restored: read(session) };
          } finally { accepted.close(); rejected.close(); }
        } finally { session.close(); }
      });
    }, { bytes: Array.from(mixedRunDocx(marker)), tracked });
    for (const [index, result] of results.entries()) {
      const expected = ['aX|cd', 'ab|Xd', 'aX|d'][index];
      expect(result.edited).toBe(true);
      expect(result.accept).toBe(true);
      expect(result.reject).toBe(true);
      expect(result.accepted).toBe(expected);
      expect(result.rejected).toBe(tracked ? 'ab|cd' : expected);
      expect(result.undo).toBe(true);
      expect(result.restored).toBe('ab|cd');
    }
  });
}
