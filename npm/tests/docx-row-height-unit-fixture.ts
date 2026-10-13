import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Alder ledger, authored as literal XML without a source DOCX. */
export function generateRowHeightUnitDocx(value: string | null, rows = 3, rule = 'atLeast'): Uint8Array {
  const paragraph = (text: string) => `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
  const body = `<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="6" w:color="444444"/>`).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>${Array.from({ length: rows }, (_, i) => `<w:tr><w:trPr><w:trHeight ${value === null ? '' : `w:val="${value}"`} w:hRule="${rule}"/></w:trPr>${cell(`Alder${String(i + 1).padStart(2, '0')} Original height record.`)}${cell(`Value${String(i + 1).padStart(2, '0')} Complete second cell.`)}</w:tr>`).join('')}</w:tbl>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
