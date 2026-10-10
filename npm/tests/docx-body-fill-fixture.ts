import { storedZip, xml } from './docx-zip.js';

/**
 * Documents that fill a Letter page's 648pt body band (1in margins) with arithmetic precision, for
 * the paginated body-overflow contract (issue #848). Every paragraph is one line at EXACT line
 * spacing with no space before or after, so each occupies exactly `lineTwips` of height.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface BodyFillOptions {
  /** One-line paragraphs; the optional table sits after the first. */
  paragraphs: number;
  /** Font size in half-points. */
  fontHalfPoints: number;
  /** Exact line height in twips (240 = 12pt). */
  lineTwips: number;
  /** Use minimum spacing to exercise glyph overflow outside the exact-spacing contract. */
  lineRule?: 'exact' | 'atLeast';
  /**
   * Put a one-row, one-cell table after the first paragraph. The converter gives a table that
   * follows a paragraph with no space after it a 7.5pt top margin of its own (Word's implicit
   * paragraph/table separation), which collapses through the table's wrapper `div`.
   */
  tableAfterFirstParagraph?: boolean;
  /** Space before the first paragraph, in twips; shifts where the page boundary falls. */
  firstSpaceBeforeTwips?: number;
  /**
   * Put a one-row table with an EXACT row height of this many twips after the first paragraph —
   * an indivisible block, taller than the 648pt band when above 12960.
   */
  exactRowTableTwips?: number;
}

export function bodyFillDocx(options: BodyFillOptions): Uint8Array {
  const pPr = (before = 0) =>
    `<w:pPr><w:spacing w:before="${before}" w:after="0" w:line="${options.lineTwips}" w:lineRule="${options.lineRule ?? 'exact'}"/></w:pPr>`;
  const rPr = `<w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="${options.fontHalfPoints}"/></w:rPr>`;
  const paragraph = (text: string, before = 0) =>
    `<w:p>${pPr(before)}<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const table = options.tableAfterFirstParagraph
    ? '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>' +
      '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${paragraph('Table cell')}</w:tc></w:tr></w:tbl>`
    : '';
  const tallTable = options.exactRowTableTwips
    ? '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
      `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:hRule="exact" w:val="${options.exactRowTableTwips}"/></w:trPr>` +
      `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${paragraph('Tall row')}</w:tc></w:tr></w:tbl>`
    : '';
  const lines = Array.from({ length: options.paragraphs },
    (_, index) => paragraph(`Line ${index + 1} of the body, typography and spacing.`,
      index === 0 ? options.firstSpaceBeforeTwips ?? 0 : 0));
  const body = lines.slice(0, 1).join('') + table + tallTable + lines.slice(1).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${body}
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
  </w:sectPr>
</w:body></w:document>`;
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}
