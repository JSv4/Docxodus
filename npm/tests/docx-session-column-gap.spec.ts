import { test, expect } from '@playwright/test';
import type { DocxSession } from '../src/session.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

function columnGapDocx(leading: boolean): Uint8Array {
  const cell = (text: string) => '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
    + `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const row = (number: number) => `<w:tr>${cell('A' + number)}${cell('B' + number)}</w:tr>`;
  const body = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    + row(1) + `<w:tr><w:trPr><w:${leading ? 'gridBefore' : 'gridAfter'} w:val="1"/></w:trPr>`
    + cell(leading ? 'B2' : 'A2') + '</w:tr>' + row(3) + '</w:tbl><w:p/>';
  const rel = (id: string, kind: string, target: string) => `<Relationship Id="${id}" Type="${R_NS}/${kind}" Target="${target}"/>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>') },
    { name: '_rels/.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + rel('rId1', 'officeDocument', 'word/document.xml') + '</Relationships>') },
    { name: 'word/_rels/document.xml.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + rel('rId1', 'styles', 'styles.xml') + rel('rId2', 'settings', 'settings.xml') + '</Relationships>') },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"/>`) },
    { name: 'word/settings.xml', data: xml(`<w:settings xmlns:w="${W_NS}"/>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`) },
  ]);
}

for (const leading of [false, true]) for (const tracked of [false, true]) {
  test(`column insertion over ${leading ? 'leading' : 'trailing'} omissions, tracking ${tracked}`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const result = await page.evaluate(({ bytes, leading, tracked }) => {
      const open = (input: Uint8Array, record = false) => (window as any).Docxodus.openTypedSession(input,
        JSON.stringify({ trackedChanges: record ? 'render_inline' : 'accept', revisionAuthor: 'Editor' })) as DocxSession;
      const geometry = (session: DocxSession) => {
        const anchor = Object.entries(session.project().anchorIndex).find(([, a]) => a.kind === 'tbl')![0];
        const table = new DOMParser().parseFromString(session.raw.getXml(anchor), 'application/xml').documentElement;
        const child = (element: Element | undefined, name: string) => Array.from(element?.children ?? []).find(e => e.localName === name);
        const number = (element: Element | undefined) => Number(element?.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val') ?? '0');
        const rows = Array.from(table.children).filter(e => e.localName === 'tr');
        return { columns: Array.from(child(table, 'tblGrid')!.children).filter(e => e.localName === 'gridCol').length,
          coverage: rows.map(row => number(child(child(row, 'trPr'), 'gridBefore')) + number(child(child(row, 'trPr'), 'gridAfter'))
            + Array.from(row.children).filter(e => e.localName === 'tc').reduce((sum, cell) => sum + (number(child(child(cell, 'tcPr'), 'gridSpan')) || 1), 0)) };
      };
      const session = open(new Uint8Array(bytes), tracked);
      try {
        const target = Object.entries(session.project().anchorIndex).find(([, a]) => a.kind === 'tc' && a.textPreview.startsWith(leading ? 'A1' : 'B1'))![0];
        const edit = session.insertTableColumn(target, leading ? 'before' : 'after');
        const revisions = session.listRevisions();
        const saved = session.save();
        const resolutions = [false, true].flatMap(accept => [false, true].map(selective => {
          const review = open(saved);
          try {
            const result = selective && tracked
              ? (accept ? review.acceptRevision(review.listRevisions()[0].id) : review.rejectRevision(review.listRevisions()[0].id))
              : (accept ? review.acceptAllRevisions() : review.rejectAllRevisions());
            return { accept, success: result.success, remaining: review.listRevisions().length, ...geometry(review) };
          } finally { review.close(); }
        }));
        return { edited: edit.success, count: revisions.length, statuses: revisions.map(r => r.resolutionStatus), resolutions };
      } finally { session.close(); }
    }, { bytes: Array.from(columnGapDocx(leading)), leading, tracked });
    expect(result.edited).toBe(true);
    expect(result.count).toBe(tracked ? 1 : 0);
    expect(result.statuses.every(status => status === 'supported')).toBe(true);
    for (const resolved of result.resolutions) {
      expect(resolved.success).toBe(true);
      expect(resolved.remaining).toBe(0);
      expect(resolved.columns).toBe(resolved.accept || !tracked ? 3 : 2);
      expect(resolved.coverage).toEqual(Array(3).fill(resolved.columns));
    }
  });
}
