import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type MarginCase = 'same' | 'left' | 'right' | 'top' | 'bottom' | 'next' | 'paper';

/** Original Sequoia paragraphs and literal section properties; no imported document content. */
export function continuousMarginsDocx(change: MarginCase, long = false): Uint8Array {
  const section = (second: boolean) => `<w:sectPr><w:type w:val="${second && change !== 'next' ? 'continuous' : 'nextPage'}"/>
    <w:pgSz w:w="${second && change === 'paper' ? 14400 : 12240}" w:h="15840"/>
    <w:pgMar ${['top', 'right', 'bottom', 'left'].map(side => `w:${side}="${second && change === side ? 2880 : 1440}"`).join(' ')}
      w:header="720" w:footer="720"/></w:sectPr>`;
  const body = `<w:p><w:pPr>${section(false)}</w:pPr><w:r><w:t>SequoiaBefore</w:t></w:r></w:p>
    <w:p><w:r><w:t>SequoiaAfter${long ? ' ' + Array.from({ length: 750 }, (_, i) => `Sequoia${String(i).padStart(3, '0')}`).join(' ') : ''}</w:t></w:r></w:p>${section(true)}`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/><w:szCs w:val="22"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
