import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export type PageBreakCase = 'own' | 'end' | 'mid' | 'leading' | 'repeated' | 'trailing' | 'none' | 'blank' | 'line' | 'space';

export function pageBreakDocx(kind: PageBreakCase, splitMark?: boolean, linked = false): Uint8Array {
  const br = '<w:br w:type="page"/>';
  const p = (content: string) => `<w:p><w:r>${content}</w:r></w:p>`;
  const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
  const body = {
    own: p(t('First')) + p(br) + p(t('Last')),
    end: (linked ? `<w:p><w:hyperlink w:anchor="Target"><w:r>${t('First')}${br}</w:r></w:hyperlink>
      <w:bookmarkStart w:id="1" w:name="Target"/><w:bookmarkEnd w:id="1"/></w:p>` : p(t('First') + br)) + p(t('Last')),
    mid: p(t('First') + br + t('After')) + p(t('Last')),
    leading: p(br + t('Last')),
    repeated: p(t('First') + br + br + t('Last')),
    trailing: p(t('First') + br),
    none: p(t('First')) + p(t('Last')),
    blank: p(t('First')) + '<w:p/>' + p(t('Last')),
    line: p(t('First') + '<w:br/>' + t('Last')),
    space: p(t('First') + br + t(' ')) + p(t('Last')),
  }[kind];
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      ${['document', 'styles', 'settings'].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}${name === 'document' ? '.main' : ''}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat>${splitMark === undefined ? '' : `<w:splitPgBreakAndParaMark w:val="${splitMark ? '1' : '0'}"/>`}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}
