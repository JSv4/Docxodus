import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

function fixture(prelude: number, mode: string): Uint8Array {
  const lines = (prefix: string, count: number) => Array.from({ length: count }, (_, i) =>
    `<w:r><w:t>${prefix}${String(i + 1).padStart(2, '0')} receipt.</w:t></w:r>`).join('<w:r><w:br/></w:r>');
  const props = mode === 'inherited' ? '<w:pStyle w:val="Inventory"/>'
    : `${mode === 'keep' ? '<w:keepLines/>' : ''}<w:widowControl w:val="${mode === 'off' ? 0 : 1}"/>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      ${['document', 'styles', 'settings'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p}${p === 'document' ? '.main' : ''}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(p => `<Relationship Id="${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Inventory"><w:name w:val="Inventory"/><w:pPr><w:widowControl/></w:pPr></w:style></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p>${lines('Prelude', prelude)}</w:p><w:p><w:pPr>${props}</w:pPr>${lines('Basil', 9)}</w:p><w:sectPr><w:pgSz w:w="6120" w:h="8000"/><w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="720"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

for (const [mode, prelude, counts] of [
  ['direct', 10, [6, 3]], ['inherited', 10, [6, 3]], ['avoid-widow', 8, [7, 2]],
  ['off', 8, [8, 1]], ['keep', 10, [0, 9]], ['avoid-orphan', 15, [0, 9]],
] as const) {
  test(`PDF widow control ${mode} keeps the expected lines on each page`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(fixture(prelude, mode), { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const pages: string[][] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
        pages.push(text.match(/Basil\d+/g) ?? []);
      }
      expect(pages.map(p => p.length)).toEqual(counts);
      expect(pages.flat()).toEqual(Array.from({ length: 9 }, (_, i) => `Basil${String(i + 1).padStart(2, '0')}`));
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
    } finally { await task.destroy(); }
  });
}
