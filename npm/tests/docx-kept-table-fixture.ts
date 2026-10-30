import { pageFittingDocx } from './docx-page-fitting-fixture.js';

/** Original kept inventory paragraph in a table cell, plus an adjacent short cell. */
export function keptTableParagraphDocx(mode:'lines'|'next'|'fitting'):Uint8Array {
  const count=mode==='fitting'?20:83;
  const lines=Array.from({length:count},(_,i)=>`<w:r><w:t>Elm${String(i+1).padStart(3,'0')} recorded.</w:t></w:r>`).join('<w:r><w:br/></w:r>');
  const before=mode==='fitting'?`<w:p><w:pPr><w:spacing w:after="0" w:line="8000" w:lineRule="exact"/></w:pPr><w:r><w:t>Elm preface.</w:t></w:r></w:p>`:'';
  return pageFittingDocx(`${before}<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="4000"/></w:tblGrid>
    <w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
      <w:p><w:pPr>${mode==='next'?'<w:keepNext/>':'<w:keepLines/>'}<w:widowControl/>
      <w:spacing w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>${lines}</w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>Elm adjacent record.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>Elm inventory complete.</w:t></w:r></w:p>`);
}
