import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

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

function pageFittingDocx(body: string): Uint8Array {
  const relations = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>`),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
