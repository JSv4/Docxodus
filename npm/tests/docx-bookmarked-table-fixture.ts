import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** Independent inventory table carrying a bookmark directly inside the table. */
export function bookmarkedTableDocx(): Uint8Array {
  const rows = Array.from({ length: 61 }, (_, row) => `<w:tr>${['A', 'B'].map(column => `<w:tc>
    <w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr><w:p><w:r>
    <w:t>Birch${String(row + 1).padStart(2, '0')}${column} confirmed</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('');
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': rels,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:tbl>
      <w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid>
      <w:bookmarkStart w:id="21" w:name="BirchLedger"/>${rows}<w:bookmarkEnd w:id="21"/>
      </w:tbl><w:p><w:r><w:t>Birch ledger complete.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
