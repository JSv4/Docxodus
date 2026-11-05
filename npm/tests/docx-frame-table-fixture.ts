import { outOfFlowDocx, textBoxRun } from './docx-out-of-flow-fixture.js';

/** Three independently positioned labels whose source owners are table cells. */
export function frameTableDocx(framed = true, cellText = false, nestedDrawing = false): Uint8Array {
  const cells = ['Fern amber', 'Fern blue', 'Fern copper'].map((label, index) =>
    `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:spacing w:before="0" w:after="0"/>
        ${framed ? `<w:framePr w:w="2000" w:hAnchor="page" w:vAnchor="page" w:x="${1200 + index * 3600}" w:y="3000" w:wrap="none"/>` : ''}
      </w:pPr>${nestedDrawing ? textBoxRun('paragraph', [label]) : `<w:r><w:t>${label}</w:t></w:r>`}</w:p>
      ${cellText ? `<w:p><w:r><w:t>Slot ${index + 1} ready.</w:t></w:r></w:p>` : ''}</w:tc>`).join('');
  return outOfFlowDocx(`<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/>
    <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>
      <w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>
    <w:tblGrid>${'<w:gridCol w:w="2880"/>'.repeat(3)}</w:tblGrid><w:tr>${cells}</w:tr></w:tbl>
    <w:p><w:r><w:t>Fern inventory complete.</w:t></w:r></w:p>`);
}
