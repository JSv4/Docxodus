import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Cedar register, authored as literal XML without a source DOCX. */
export function generateIndentationUnitDocx(value: string): Uint8Array {
  const words = 'Cedar archive entries retain their order across the paragraph line while the register is inspected and checked. '.repeat(3).trim();
  const paragraph = (marker: string, attributes: string, border = false, bidi = false) => `<w:p><w:pPr>
    ${border ? '<w:pBdr><w:bottom w:val="single" w:sz="8" w:color="246B7D"/></w:pBdr>' : ''}
    ${bidi ? '<w:bidi/>' : ''}<w:spacing w:before="0" w:after="120"/>
    <w:ind ${attributes}/></w:pPr><w:r><w:t>${marker} ${words}</w:t></w:r></w:p>`;
  const body = paragraph('EdgeMarker', `w:left="${value}" w:right="${value}"`)
    + paragraph('LogicalMarker', `w:start="${value}" w:end="${value}" w:left="1440" w:right="1440"`, false, true)
    + paragraph('FirstMarker', `w:left="${value}" w:firstLine="${value}"`)
    + paragraph('BorderMarker', `w:left="${value}" w:hanging="${value}"`, true);
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
