import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/**
 * 120 one-line paragraphs of `size`pt text with `spacing` in each. `tracked` makes the last one a tracked
 * insertion; `'del'` follows every paragraph's text with a tracked deletion. `word` is the run properties
 * of a word that follows every paragraph's label, and `bottom` the page's bottom margin in twips.
 */
export function lineGridDocx(pitch: number, snap = true, spacing = '', tracked: boolean | 'del' = false, size = 12,
  word = '', bottom = 1440): Uint8Array {
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const parts = {
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
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="${size * 2}"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>` +
      Array.from({ length: 120 }, (_, i) => {
        const run = `<w:r><w:t>Cedar${String(i + 1).padStart(3, '0')}</w:t></w:r>`
          + (word ? `<w:r><w:rPr>${word}</w:rPr><w:t xml:space="preserve"> seed</w:t></w:r>` : '');
        const deleted = tracked === 'del' ? `<w:del w:id="${i + 2}" w:author="Editor"><w:r>
          <w:delText xml:space="preserve"> old</w:delText></w:r></w:del>` : '';
        return `<w:p><w:pPr>${snap ? '' : '<w:snapToGrid w:val="0"/>'}${spacing}</w:pPr>
          ${tracked === true && i === 119 ? `<w:ins w:id="1" w:author="Editor">${run}</w:ins>` : run}${deleted}</w:p>`;
      }).join('') +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="${bottom}" w:left="1440"/>
      ${pitch ? `<w:docGrid w:type="lines" w:linePitch="${pitch * 20}"/>` : ''}</w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
