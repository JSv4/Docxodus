import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

/**
 * Issue #1571 in PDF export: one-line double-spaced 12pt paragraphs, with a text area of 23 full
 * pitches plus `spareTwips` / 20 pt. Word 2024 prints 24 lines at a 15pt remainder, whose text height
 * (13.80pt) fits, and 23 at 13pt (docs/ooxml_corner_cases.md, #1571).
 */
function doubleSpacedDocx(bottomTwips: number): Uint8Array {
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const paragraphs = Array.from({ length: 51 }, (_, index) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
    + `w:line="480" w:lineRule="auto"/></w:pPr><w:r><w:t>Line ${index + 1} of the harbor log.</w:t></w:r></w:p>`).join('');
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
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="24"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="${bottomTwips}" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr></w:body></w:document>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

for (const [spare, bottom, lines] of [[15, 1404, 24], [13, 1444, 23]] as const) {
  test(`a page's last double-spaced line with ${spare}pt to spare prints ${lines} lines`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(doubleSpacedDocx(bottom), { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const rows = new Set(content.items.flatMap(item => 'str' in item && item.str.trim()
        ? [Math.round(item.transform[5])] : []));
      const text = content.items.map(item => 'str' in item ? item.str : '').join(' ').replace(/\s+/g, ' ');
      expect(rows.size).toBe(lines);
      expect(text).toContain(`Line ${lines} of`);
      expect(text).not.toContain(`Line ${lines + 1} of`);
    } finally { await task.destroy(); }
  });
}
