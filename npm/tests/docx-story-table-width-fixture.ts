import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original two-page inventory with the same three-column grid in each running story. */
export function storyTableWidthDocx(kind: 'wide' | 'normal' | 'margin'): Uint8Array {
  const width = kind === 'wide' ? 10800 : kind === 'normal' ? 5040 : 6600;
  const paragraph = (text: string) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const table = (story: string) => `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/>
    <w:tblLayout w:type="fixed"/>${kind === 'margin' ? '<w:tblInd w:w="-400" w:type="dxa"/>' : ''}
    <w:tblCellMar>${['top','bottom','left','right'].map(edge => `<w:${edge} w:w="0" w:type="dxa"/>`).join('')}</w:tblCellMar>
    </w:tblPr><w:tblGrid>${`<w:gridCol w:w="${width / 3}"/>`.repeat(3)}</w:tblGrid><w:tr>
    ${['amber','blue','copper'].map(label => `<w:tc><w:tcPr><w:tcW w:w="${width / 3}" w:type="dxa"/></w:tcPr>${paragraph(`Moss ${story} ${label}`)}</w:tc>`).join('')}
    </w:tr></w:tbl>`;
  const parts: Record<string, string> = {
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${paragraph('Moss inventory opens.')}
      <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Moss inventory closes.</w:t></w:r></w:p><w:sectPr>
      <w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/>
      <w:pgSz w:w="7200" w:h="9600"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/>
      </w:sectPr></w:body></w:document>`,
    'word/header.xml': `<w:hdr xmlns:w="${W_NS}">${table('top')}</w:hdr>`,
    'word/footer.xml': `<w:ftr xmlns:w="${W_NS}">${table('base')}</w:ftr>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="18"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  };
  const types = {document: 'document.main', header:'header', footer:'footer', styles:'styles'};
  const relationships = (entries: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      ${Object.entries(types).map(([name,type]) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': relationships(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relationships(['header','footer','styles'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')),
    ...parts,
  }).map(([name,value]) => ({name,data:xml(value)})));
}
