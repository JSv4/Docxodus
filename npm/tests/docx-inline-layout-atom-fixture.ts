import { fittingImageRun, pageFittingDocx } from './docx-page-fitting-fixture.js';

export function inlineLayoutAtomDocx(kind:'shape'|'tab'|'tab-cell'):Uint8Array {
  if(kind==='shape') return pageFittingDocx(`<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/></w:pPr>
    ${Array.from({length:5},(_,i)=>`<w:r><w:pict><v:oval xmlns:v="urn:schemas-microsoft-com:vml" id="AlderOval${i}" style="width:72pt;height:120pt"/></w:pict></w:r>`).join('')}
    ${fittingImageRun(81,320,590)}</w:p><w:p><w:r><w:t>Alder shape inventory complete.</w:t></w:r></w:p>`);
  const lines=Array.from({length:83},(_,i)=>`<w:r><w:t>Willow${String(i+1).padStart(3,'0')} recorded.</w:t></w:r>`).join('<w:r><w:br/></w:r>');
  const paragraph=`<w:p><w:pPr>${kind==='tab-cell'?'<w:keepNext/>':''}<w:keepLines w:val="0"/><w:widowControl w:val="0"/>
    <w:spacing w:after="0" w:line="320" w:lineRule="exact"/><w:tabs><w:tab w:val="left" w:pos="720" w:leader="dot"/></w:tabs></w:pPr>
    <w:r><w:t>R7.</w:t><w:tab/></w:r>${lines}</w:p>`;
  const body=kind==='tab'?paragraph:`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${paragraph}</w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>Willow neighbor.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  return pageFittingDocx(`${body}<w:p><w:r><w:t>Willow tab inventory complete.</w:t></w:r></w:p>`);
}
