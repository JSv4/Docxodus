import { storedZip, xml } from './docx-zip.js';

/** Two original tables whose widths, borders and cell margins must remain independent. */
export function adjacentTablesDocx(separated = false): Uint8Array {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const paragraph = (text: string) => '<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>';
  const table = (text: string, second: boolean) => {
    const width = second ? 6000 : 4800;
    const borders = second ? '<w:tblBorders>' + ['top', 'left', 'bottom', 'right']
      .map(side => '<w:' + side + ' w:val="single" w:sz="8" w:color="246B7D"/>').join('') + '</w:tblBorders>' : '';
    return '<w:tbl><w:tblPr><w:tblW w:w="' + width + '" w:type="dxa"/>' + borders
      + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="' + (second ? 720 : 0)
      + '" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="' + width + '"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="'
      + width + '" w:type="dxa"/></w:tcPr>' + paragraph(text) + '</w:tc></w:tr></w:tbl>';
  };
  const body = paragraph('HarborEntry') + table('FirstDock', false) + (separated ? paragraph('') : '')
    + table('SecondDock', true) + paragraph('HarborExit');
  const parts = [
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="' + r + '/officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="' + r + '/styles" Target="styles.xml"/>'
      + '<Relationship Id="rId2" Type="' + r + '/settings" Target="settings.xml"/></Relationships>'],
    ['word/styles.xml', '<w:styles xmlns:w="' + w + '"><w:docDefaults><w:rPrDefault><w:rPr>'
      + '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
      + '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/>'
      + '</w:pPr></w:pPrDefault></w:docDefaults></w:styles>'],
    ['word/settings.xml', '<w:settings xmlns:w="' + w + '"/>'],
    ['word/document.xml', '<w:document xmlns:w="' + w + '"><w:body>' + body
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
      + '</w:sectPr></w:body></w:document>'],
  ];
  return storedZip(parts.map(([name, content]) => ({ name, data: xml(content) })));
}
