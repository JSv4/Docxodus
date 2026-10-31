import { storedZip, xml, W_NS as w, R_NS as r } from './docx-zip.js';

export type RubyFixture = 'simple' | 'zero-raise' | 'high-raise' | 'multiple' | 'long' | 'plain' | 'long-plain' | 'vertical' | 'tab';

export function rubyDocx(kind: RubyFixture): Uint8Array {
  const run = (text: string, properties = '') => `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const multiple = kind === 'multiple', plain = kind === 'plain' || kind === 'long-plain';
  const pair = plain ? run('CEDAR reading') : `<w:r><w:ruby><w:rubyPr>
    <w:rubyAlign w:val="${kind === 'vertical' ? 'rightVertical' : 'center'}"/><w:hps w:val="12"/>
    <w:hpsRaise w:val="${kind === 'zero-raise' ? '0' : kind === 'high-raise' ? '44' : '22'}"/><w:hpsBaseText w:val="22"/><w:lid w:val="en-US"/></w:rubyPr>
    <w:rt>${multiple ? run('read', '<w:b/>') + run('ing', '<w:i/>') : run('reading')}</w:rt>
    <w:rubyBase>${multiple ? run('CE', '<w:b/>') + run('DAR') : run('CEDAR')}</w:rubyBase></w:ruby></w:r>`;
  const words = Array.from({ length: 100 }, (_, i) => `Pine${String(i).padStart(3, '0')}`);
  const long = kind === 'long' || kind === 'long-plain';
  const body = run(long ? words.slice(0, 35).join(' ') + ' ' : 'CedarBefore ') + pair
    + (kind === 'tab' ? '<w:r><w:tab/></w:r>' : '') + run(long ? ' ' + words.slice(35).join(' ') : ' CedarAfter');
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + [['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings']].map(([name, type]) =>
        `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles', 'settings'].map(name => `<Relationship Id="${name}" Type="${r}/${name}" Target="${name}.xml"/>`).join('') + '</Relationships>',
    'word/document.xml': `<w:document xmlns:w="${w}"><w:body><w:p>${body}</w:p><w:p>${run('PineTail')}</w:p><w:sectPr>
      <w:pgSz w:w="6000" w:h="4000"/><w:pgMar w:top="600" w:left="600" w:bottom="600" w:right="600" w:header="300" w:footer="300"/>
      </w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
      <w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${w}"/>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
