import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export interface BorderSizeOptions {
  size?: string;
  style?: 'single' | 'double' | 'dashed' | 'nil' | 'none';
  cell?: boolean;
}

/** Original Cypress card, authored directly without an input DOCX or template. */
export function generateBorderSizeDocx(options: BorderSizeOptions = {}): Uint8Array {
  const borders = ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:val="${options.style ?? 'single'}"
    w:color="246B7D" w:space="4"${options.size === undefined ? '' : ` w:sz="${options.size}"`}/>`).join('');
  const text = 'BorderMarker Cypress archive checked.';
  const p = `<w:p><w:pPr>${options.cell ? '' : `<w:pBdr>${borders}</w:pBdr>`}<w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const body = options.cell ? `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/><w:tcBorders>${borders}</w:tcBorders></w:tcPr>${p}</w:tc></w:tr></w:tbl>` : p;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
