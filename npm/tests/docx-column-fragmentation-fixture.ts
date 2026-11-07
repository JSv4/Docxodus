import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export type ColumnContent = 'paragraph' | 'wrapped' | 'table';

export function columnFragmentationDocx(kind: ColumnContent, columns = 2, sequential = false, count = 77): Uint8Array {
  const labels = Array.from({ length: count }, (_, i) => `Maple${String(i + 1).padStart(3, '0')}`);
  const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const paragraph = (runs: string) => `<w:p>${runs}</w:p>`;
  const content = kind === 'table'
    ? `<w:tbl><w:tblPr><w:tblW w:w="2800" w:type="dxa"/><w:tblLayout w:type="fixed"/>
      <w:tblCellMar>${['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="0" w:type="dxa"/>`).join('')}</w:tblCellMar></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2800"/></w:tblGrid>${labels.map(label => `<w:tr><w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr>${paragraph(run(label))}</w:tc></w:tr>`).join('')}</w:tbl>`
    : paragraph(kind === 'wrapped' ? run(labels.map(label => `${label} register.`).join(' '))
      : labels.map(run).join('<w:r><w:br/></w:r>'));
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      ${['document', 'styles', 'settings'].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}${name === 'document' ? '.main' : ''}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${content}${paragraph(run('MapleEnd'))}
      <w:sectPr>${sequential ? '<w:type w:val="continuous"/>' : ''}<w:pgSz w:w="6800" w:h="4000"/>
      <w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="0" w:footer="0"/>
      <w:cols w:num="${columns}" w:space="400"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
      <w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}
