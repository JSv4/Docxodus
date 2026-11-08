import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export const fitTexts = {
  justify: 'Every autumn the ferry crews repaint the harbor signs before the storms reach our northern pier as the gulls complain about it all afternoon long.',
  left: 'We Took A Yellow Wagon To Avery Valley, Past Your Tower, Toward Lake Vale, Where Water Waits.',
};
export type FitNeighbor = 'plain' | 'bold' | 'table' | 'list' | 'link';

export function textFitDocx(alignment: keyof typeof fitTexts, neighbor: FitNeighbor,
  kern?: { owner: 'run' | 'defaults' | 'style'; halfPoints: number }): Uint8Array {
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const run = (text: string, props = '') => `<w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const p = (runs: string, props = '') => `<w:p><w:pPr>${props}</w:pPr>${runs}</w:p>`;
  const kerning = kern ? `<w:kern w:val="${kern.halfPoints}"/>` : '';
  const closing = {
    plain: p(run('Closing note.')),
    bold: p(run('A short ') + run('closing', '<w:b/>') + run(' note.')),
    table: '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' + p(run('Cell text.')) + '</w:tc></w:tr></w:tbl>',
    list: p(run('List text.'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
    link: p(`<w:hyperlink r:id="link">${run('Link text.')}</w:hyperlink>`),
  }[neighbor];
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="${wml}.document.main+xml"/>
      ${['styles', 'settings', 'numbering'].map(name => `<Override PartName="/word/${name}.xml" ContentType="${wml}.${name}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      ${['styles', 'settings', 'numbering'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}
      <Relationship Id="link" Type="${R_NS}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>${kern?.owner === 'defaults' ? kerning : ''}<w:sz w:val="24"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal">
      <w:name w:val="Normal"/><w:rPr>${kern?.owner === 'style' ? kerning : ''}</w:rPr></w:style></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
    'word/numbering.xml': `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>
      <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/>
      </w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
      p(run(fitTexts[alignment], kern?.owner === 'run' ? kerning : ''), `<w:jc w:val="${alignment === 'justify' ? 'both' : 'left'}"/>`) + closing +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
