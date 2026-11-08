import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export const minimumOwnerText = 'Teasel records the inventory before the ferry crossing. ';

/** The same uniform text in an ordinary cell or a numbered paragraph. */
export function minimumOwnerDocx(owner: 'cell' | 'list', natural: boolean, repeats: number, suffix: 'tab' | 'nothing' = 'tab'): Uint8Array {
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
  const line = natural ? 'w:line="240" w:lineRule="auto"' : 'w:line="320" w:lineRule="atLeast"';
  const numbering = owner === 'list' ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '';
  const paragraph = `<w:p><w:pPr><w:widowControl w:val="0"/>${numbering}<w:spacing w:before="0" w:after="0" ${line}/>
    </w:pPr><w:r><w:t xml:space="preserve">${minimumOwnerText.repeat(repeats).trim()}</w:t></w:r></w:p>`;
  const content = owner === 'cell' ? `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>
    ${paragraph}</w:tc></w:tr></w:tbl>` : paragraph;
  const rels = (value: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${value}</Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      ${[['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings'], ['numbering', 'numbering']]
        .map(([name, suffix]) => `<Override PartName="/word/${name}.xml" ContentType="${type}${suffix}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': rels(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': rels(['styles', 'settings', 'numbering'].map(name =>
      `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="40"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/numbering.xml': `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="${suffix}"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
      <w:sz w:val="40"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${content}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
