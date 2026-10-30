import { fittingImageRun, pageFittingDocx } from './docx-page-fitting-fixture.js';

/** Original four-column ledger with overlapping vertical and horizontal merges. */
export function oversizedMergeDocx(rowCount = 46): Uint8Array {
  const paragraph = (text: string, image = '') => `<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
    <w:spacing w:after="0" w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:t>${text}</w:t></w:r>${image}</w:p>`;
  const cell = (width: number, content: string, properties = '') => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>
    <w:vAlign w:val="top"/>${properties}</w:tcPr>${content}</w:tc>`;
  const header = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${[900,900,1800,5400].map((width,i)=>cell(width,paragraph(`CedarHeading${i}`))).join('')}</w:tr>`;
  const crossingStart = Math.floor(rowCount / 2);
  const rows = Array.from({ length: rowCount }, (_, row) => {
    const number = String(row + 1).padStart(2, '0');
    const wide = cell(1800, row === 0 ? paragraph('CedarSpan', fittingImageRun(71,32,16)) : '<w:p/>',
      `<w:gridSpan w:val="2"/><w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`);
    const crossing = cell(1800, row < crossingStart ? paragraph(`Copse${number}`)
      : row === crossingStart ? paragraph('CedarOverlap') : '<w:p/>', row >= crossingStart
        ? `<w:vMerge${row === crossingStart ? ' w:val="restart"' : ''}/>` : '');
    const details = cell(5400, ['A','B','C'].map((suffix,index)=>paragraph(`Cedar${number}${suffix} checked`,
      row === rowCount-1 && index === 2 ? fittingImageRun(72,24,12) : '')).join(''));
    return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${wide}${crossing}${details}</w:tr>`;
  }).join('');
  return pageFittingDocx(`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>
      <w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>
    </w:tblPr><w:tblGrid>${[900,900,1800,5400].map(width=>`<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${header}${rows}</w:tbl>
    <w:p><w:r><w:t>Cedar ledger complete.</w:t></w:r></w:p>`);
}
