import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export interface VerticalMergeOptions {
  merged?: boolean;
  mergeStart?: number;
  mergeSpan?: number;
  rowCount?: number;
  colSpan?: 1 | 2;
  rightMerge?: { start: number; span: number };
  rowHeightRule?: 'atLeast' | 'exact';
}

/** Original two-column register; the control differs only in vertical merge metadata. */
export function generateVerticalMergeDocx(options: VerticalMergeOptions = {}): Uint8Array {
  const { merged = true, mergeStart = 1, mergeSpan = 2, rowCount = 18, colSpan = 1 } = options;
  const columns = colSpan + 1 + (options.rightMerge ? 1 : 0);
  const columnWidth = 9360 / columns;
  const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const rows = Array.from({ length: rowCount }, (_, index) => {
    const row = index + 1;
    const continues = row > mergeStart && row < mergeStart + mergeSpan;
    const merge = !merged ? '' : row === mergeStart ? '<w:vMerge w:val="restart"/>' : continues ? '<w:vMerge/>' : '';
    const rightContinues = options.rightMerge && row > options.rightMerge.start && row < options.rightMerge.start + options.rightMerge.span;
    const rightMerge = !merged || !options.rightMerge ? '' : row === options.rightMerge.start
      ? '<w:vMerge w:val="restart"/>' : rightContinues ? '<w:vMerge/>' : '';
    const cell = (content: string, width: number, properties = '') => `<w:tc><w:tcPr>
      <w:tcW w:w="${width}" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;
    return `<w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="${options.rowHeightRule ?? 'atLeast'}"/><w:cantSplit/></w:trPr>`
      + cell(continues ? '<w:p/>' : paragraph(`Left${String(row).padStart(2, '0')} Cedar tray.`), columnWidth * colSpan,
        (colSpan > 1 ? `<w:gridSpan w:val="${colSpan}"/>` : '') + merge)
      + cell(rightContinues ? '<w:p/>' : paragraph(`Right${String(row).padStart(2, '0')} Archive register.`), columnWidth, rightMerge)
      + (options.rightMerge ? cell(paragraph(`Row${String(row).padStart(2, '0')}`), columnWidth) : '') + '</w:tr>';
  }).join('');
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:tbl><w:tblPr>
      <w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>`
      + Array.from({ length: columns }, () => `<w:gridCol w:w="${columnWidth}"/>`).join('')
      + `</w:tblGrid>${rows}</w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
