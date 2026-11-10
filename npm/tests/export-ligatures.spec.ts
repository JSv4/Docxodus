import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

// An original test face (fixtures/fonts/generate-ligature-font.py): 7.5pt letters, 3pt spaces
// and one standard ligature, F F, which is 6pt narrower than the pair at 12pt. Every advance is a
// whole pixel. Fourteen of these words fit a 468pt line unligated (459pt); seventeen fit ligated (456pt).
const fonts = fileURLToPath(new URL('./fixtures/fonts/ligature/', import.meta.url));
const words = Array.from('ABCDEGHIJKLMNOPQRSTUVWXYZ', letter => `${letter}FF${letter}`);

function ligatureDocx(ligatures: string | null): Uint8Array {
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const props = ligatures ? `<w14:ligatures w14:val="${ligatures}"/>` : '';
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="${wml}.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="${wml}.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Docxodus Ligature Test" w:hAnsi="Docxodus Ligature Test"/><w:sz w:val="24"/></w:rPr>
      </w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${words.join(' ')}</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      </w:sectPr></w:body></w:document>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

// Word 2024 forms no ligatures without w14:ligatures and forms them with it (#1562).
for (const [ligatures, count] of [[null, 14], ['none', 14], ['standard', 17], ['all', 17]] as const) {
  test(`w14:ligatures ${ligatures ?? 'absent'} puts ${count} words on the first PDF line`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(ligatureDocx(ligatures), { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', fontDirectories: [fonts], strictFonts: true,
      timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const items = content.items.flatMap(item => 'str' in item && item.str.trim()
        ? [{ text: item.str, x: item.transform[4], y: Math.round(item.transform[5]) }] : []);
      const top = Math.max(...items.map(item => item.y));
      // A ligature glyph can extract as its own item; compare the line's letters.
      const first = items.filter(item => item.y === top).sort((a, b) => a.x - b.x).map(item => item.text).join('');
      expect(first.replace(/\s/g, '')).toBe(words.slice(0, count).join(''));
    } finally { await task.destroy(); }
  });
}
