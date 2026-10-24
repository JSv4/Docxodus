import { storedZip, xml } from './docx-zip.js';

/**
 * A paragraph, one or two one-row tables and a closing paragraph, for the paragraph-to-table
 * boundary (issue #1244). Every line is an exact 16pt line with no paragraph spacing, so the gap
 * between two baselines is the sum of known heights and nothing else.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface TableBoundaryOptions {
  /** The opening paragraph's `w:spacing/@w:after` in twips; `null` leaves the attribute out. */
  afterTwips: string | null;
  /**
   * Follow the first table directly with a second one. Their distinct boundaries must survive the
   * export's accept-revisions pass (#1341), so the second table's own position is measured.
   */
  adjacent?: boolean;
  borders?: boolean;
}

export function tableBoundaryDocx({ afterTwips, adjacent = false, borders = false }: TableBoundaryOptions): Uint8Array {
  const paragraph = (text: string, props = '') => `<w:p><w:pPr>${props}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const edge = borders
    ? '<w:tblBorders>' + ['top', 'left', 'bottom', 'right'].map(s => `<w:${s} w:val="single" w:sz="8" w:color="000000"/>`).join('') + '</w:tblBorders>'
    : '';
  const table = (text: string) => `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/>${edge}<w:tblLayout w:type="fixed"/>`
    + '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="360" w:hRule="atLeast"/></w:trPr>'
    + `<w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc></w:tr></w:tbl>`;
  const before = paragraph('BoundaryBefore', afterTwips === null ? '' : `<w:spacing w:after="${afterTwips}"/>`);
  const second = adjacent
    ? table('BoundarySecond')
    : '';
  const body = before + table('BoundaryCell') + second + paragraph('BoundaryAfter');
  const font = 'Liberation Serif';
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`) },
    { name: '_rels/.rels', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`) },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
      + `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>`
      + '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault>'
      + '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>') },
  ]);
}
