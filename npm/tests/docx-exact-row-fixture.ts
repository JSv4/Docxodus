import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * Exact-height table rows (`w:trHeight w:hRule="exact"`), authored as literal XML: Word draws only what
 * fits a positive row height and clips the rest; an exact zero uses the content height.
 */
export interface ExactRowOptions {
  /** Body rows below the header. Enough of them run the table onto later pages. */
  bodyRows?: number;
  /** The header row's `w:hRule`; omitted, the table has no header row. */
  headerRule?: 'exact' | 'atLeast' | 'auto';
  /**
   * Paragraphs in the first body row's first cell, held to one line's height by an exact row, so the
   * second is cut and the third and later are hidden whole.
   */
  clippedParagraphs?: number;
  /**
   * Top and bottom paragraph borders (red, `w:sz="8"`) on every paragraph of the clipped cell. Word
   * draws a run of paragraphs with the same borders as one bordered group: its top border lies inside
   * the row, its bottom border below the last paragraph, where the row hides it.
   */
  borders?: boolean;
  /**
   * A box border (green, `w:sz="8"`, `w:space="10"`) on every side of the clipped cell's paragraphs.
   * Word draws the left and right sides outside the text, in the cell margin.
   */
  box?: boolean;
  /** The first body row's exact height in twips; 300 holds one 11pt line. */
  firstRowTwips?: number;
  /** Table borders (`w:tblBorders`), independent of the row height constraint. */
  tableBorders?: boolean;
  /** A comment (id 7) on the header row's first paragraph. */
  headerComment?: boolean;
}

const BORDERS = '<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/>'
  + '<w:bottom w:val="single" w:sz="8" w:space="1" w:color="FF0000"/></w:pBdr>';
const BOX = `<w:pBdr>${['top', 'left', 'bottom', 'right']
  .map(side => `<w:${side} w:val="single" w:sz="8" w:space="10" w:color="00AA00"/>`).join('')}</w:pBdr>`;
const paragraph = (text: string, pBdr = '', comment = false) =>
  `<w:p><w:pPr>${pBdr}`
  + `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`
  + (comment ? '<w:commentRangeStart w:id="7"/>' : '')
  + `<w:r><w:t>${text}</w:t></w:r>`
  + (comment ? '<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>' : '')
  + '</w:p>';
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const row = (height: string, rule: string, cells: string, header = false) =>
  `<w:tr><w:trPr>${header ? '<w:tblHeader/>' : ''}<w:trHeight w:val="${height}" w:hRule="${rule}"/></w:trPr>${cells}</w:tr>`;
const number = (index: number) => String(index).padStart(2, '0');

export const CLIPPED_WORDS = ['Alder', 'Birch', 'Cedar', 'Damson', 'Elder'];

export function generateExactRowDocx(options: ExactRowOptions = {}): Uint8Array {
  const { bodyRows = 3, headerRule, clippedParagraphs = 1, borders = false, box = false, headerComment = false,
    firstRowTwips = 300, tableBorders = true } = options;
  const pBdr = box ? BOX : borders ? BORDERS : '';
  const rows: string[] = [];
  if (headerRule)
    rows.push(row('400', headerRule, cell(paragraph('Heading Ledger', '', headerComment))
      + cell(paragraph('Heading Value')), true));
  for (let index = 1; index <= bodyRows; index++) {
    const first = index === 1
      ? CLIPPED_WORDS.slice(0, clippedParagraphs).map(word => paragraph(`${word} clipped cell line.`, pBdr)).join('')
      : paragraph(`Row${number(index)} ledger entry.`);
    // The first body row is one line high by default, so its later paragraphs fall outside it; other
    // rows have room for their single line.
    rows.push(row(index === 1 ? String(firstRowTwips) : '400', 'exact', cell(first) + cell(paragraph(`Value${number(index)} entry.`))));
  }
  const borderXml = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map(side => `<w:${side} w:val="single" w:sz="4" w:color="444444"/>`).join('');
  const body = `<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/>${tableBorders ? `<w:tblBorders>${borderXml}</w:tblBorders>` : ''}`
    + `<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>`
    + `${rows.join('')}</w:tbl>${paragraph('Closing paragraph after the table.')}`;
  const comments = headerComment
    ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
    : '';
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${comments}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/>${headerComment ? `<Relationship Id="comments" Type="${R_NS}/comments" Target="comments.xml"/>` : ''}</Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  if (headerComment) {
    parts['word/comments.xml'] = `<w:comments xmlns:w="${W_NS}"><w:comment w:id="7" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="R">`
      + `<w:p><w:r><w:t>Header ledger note.</w:t></w:r></w:p></w:comment></w:comments>`;
  }
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
