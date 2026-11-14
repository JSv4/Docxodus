import { storedZip, W_NS, xml } from './docx-zip.js';

export interface GridTableRow {
  /** Each cell's `w:tcW` in twips and its `w:gridSpan` (default 1). */
  cells: Array<[twips: number, span?: number]>;
  /** `w:trPr` children, such as `w:gridAfter`. */
  props?: string;
}

/**
 * Issue #1750: one bordered fixed-layout table on A4 with 1,418-twip margins, as in #1733's
 * reproducer. The grid, the rows and the preferred width are the variables. Row `n`'s cells read
 * A1, A2, … for the first row, B1, … for the second, so a test can find each cell's text.
 */
export function gridTableDocx(grid: number[], rows: GridTableRow[], tableWidth = grid.reduce((a, b) => a + b, 0)): Uint8Array {
  const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map(side => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
  const body = rows.map(({ cells, props }, row) => '<w:tr>' + (props ? `<w:trPr>${props}</w:trPr>` : '')
    + cells.map(([twips, span = 1], column) => `<w:tc><w:tcPr><w:tcW w:w="${twips}" w:type="dxa"/>`
      + (span > 1 ? `<w:gridSpan w:val="${span}"/>` : '') + '</w:tcPr>'
      + `<w:p><w:r><w:t>${String.fromCharCode(65 + row)}${column + 1}</w:t></w:r></w:p></w:tc>`).join('') + '</w:tr>').join('');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="${tableWidth}" w:type="dxa"/>${borders}<w:tblLayout w:type="fixed"/></w:tblPr>`
    + `<w:tblGrid>${grid.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${body}</w:tbl>`;
  const font = 'Liberation Serif';
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`) },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`) },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>`
      + `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Inventory sheet</w:t></w:r></w:p>`
      + `${table}<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>`
      + '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>') },
  ]);
}
