import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Wren inventory with independently selectable decimal measure spellings. */
export function fractionalMeasuresDocx(options: {
  tab?: string; rowHeight?: string; pageWidth?: string; pageHeight?: string;
} = {}): Uint8Array {
  const { tab = '1827.875', rowHeight = '687.25', pageWidth = '12240.5', pageHeight = '15840.5' } = options;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}">
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
      <w:style w:type="paragraph" w:styleId="MeasureTrack"><w:name w:val="Measure Track"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="${tab}"/></w:tabs></w:pPr></w:style>
    </w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="MeasureTrack"/></w:pPr><w:r><w:t>Wren</w:t><w:tab/><w:t>field inventory.</w:t></w:r></w:p>
      <w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr><w:trPr><w:trHeight w:val="${rowHeight}" w:hRule="atLeast"/></w:trPr>
          <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Trail sample 07.</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>All items retained.</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:p><w:r><w:t>Field inspection complete.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
    </w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
