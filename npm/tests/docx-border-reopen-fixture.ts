import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Kestrel inventory cards with separate paragraph borders. */
export function generateBorderReopenDocx(style: 'single' | 'dashed', control = false): Uint8Array {
  const cards = Array.from({ length: 18 }, (_, index) => {
    const sides = ['top', 'left', 'bottom', 'right'].map(side =>
      `<w:${side} w:val="${style}" w:sz="${control ? 12 : 10}" w:space="${control ? 3 : 1}"
        w:color="${index % 2 ? '594A79' : '386E60'}"/>`).join('');
    return `<w:p><w:pPr><w:pBdr>${sides}</w:pBdr><w:spacing w:before="0" w:after="0"/></w:pPr>
      <w:r><w:t>Kestrel card ${index + 1}: specimen logged.</w:t></w:r></w:p>`;
  }).join('');
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship
      Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship
      Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${cards}
      <w:p><w:r><w:t>Inventory complete.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
