import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export function longBodyCellDocx(shape: 'paragraph' | 'nested-table', count = 73): Uint8Array {
  const lines = Array.from({ length: count }, (_, index) => `Umber${String(index + 1).padStart(3, '0')} shipment logged.`);
  const paragraph = (runs: string) => `<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
    <w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>${runs}</w:p>`;
  const run = (text: string) => `<w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>
    <w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
  const cell = (content: string, width: number) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>
    <w:vAlign w:val="top"/></w:tcPr>${content}</w:tc>`;
  const table = (rows: string, widths: number[]) => `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a,b) => a+b,0)}" w:type="dxa"/>
    <w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="20" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/>
      <w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tblCellMar></w:tblPr>
    <w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
  const content = shape === 'paragraph'
    ? paragraph(lines.map(run).join('<w:r><w:br/></w:r>'))
    : table(lines.map(line => `<w:tr>${cell(paragraph(run(line)), 5800)}</w:tr>`).join(''), [5800]) + paragraph('');
  const body = table(`<w:tr>${cell(content, 6000)}${cell(paragraph(run('Receipt confirmed.')), 3360)}</w:tr>`, [6000,3360])
    + paragraph(run('Inventory complete.'));
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440"
        w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}
