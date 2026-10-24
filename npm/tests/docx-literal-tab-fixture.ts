import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type LiteralTabCase = 'body' | 'header' | 'footer' | 'multiple' | 'deleted';

/** Original tab-only, mixed-text and deleted-text inputs, with explicit-tab controls. */
export function literalTabFixture(kind: LiteralTabCase, literal = true): Uint8Array {
  const format = '<w:rPr><w:b/><w:color w:val="225588"/></w:rPr>';
  const tag = kind === 'deleted' ? 'delText' : 't';
  const text = (value: string) => '<w:' + tag + ' xml:space="preserve">' + value + '</w:' + tag + '>';
  let content: string;
  if (kind === 'multiple' || kind === 'deleted') {
    const tabs = kind === 'multiple' ? 2 : 1;
    content = '<w:r>' + format + (literal ? text('North' + '&#9;'.repeat(tabs) + 'HarborMark')
      : text('North') + '<w:tab/>'.repeat(tabs) + text('HarborMark')) + '</w:r>';
  } else {
    content = '<w:r><w:t/></w:r><w:r>' + format + (literal ? text('&#9;') : '<w:tab/>')
      + '</w:r><w:r>' + format + text('HarborMark') + '</w:r>';
  }
  if (kind === 'deleted') content = '<w:del w:id="7" w:author="Ferry reviewer" w:date="2026-01-01T00:00:00Z">' + content + '</w:del>';
  const paragraph = '<w:p><w:pPr><w:tabs>' + (kind === 'multiple' ? '<w:tab w:val="left" w:pos="2400"/>' : '')
    + '<w:tab w:val="right" w:pos="9360" w:leader="dot"/></w:tabs></w:pPr>' + content + '</w:p>';
  const story = kind === 'header' || kind === 'footer' ? kind : null;
  const overrides = ['document', 'styles', 'settings'].map(part => '<Override PartName="/word/' + part
    + '.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.'
    + (part === 'document' ? 'document.main' : part) + '+xml"/>').join('');
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>' + overrides
      + (story ? '<Override PartName="/word/' + story + '1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.' + story + '+xml"/>' : '') + '</Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="'
      + R_NS + '/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles', 'settings'].map(part => '<Relationship Id="r' + part + '" Type="' + R_NS + '/' + part + '" Target="' + part + '.xml"/>').join('')
      + (story ? '<Relationship Id="rStory" Type="' + R_NS + '/' + story + '" Target="' + story + '1.xml"/>' : '') + '</Relationships>',
    'word/styles.xml': '<w:styles xmlns:w="' + W_NS + '"><w:docDefaults><w:rPrDefault><w:rPr>'
      + '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
      + '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
    'word/settings.xml': '<w:settings xmlns:w="' + W_NS + '"><w:defaultTabStop w:val="720"/></w:settings>',
    'word/document.xml': '<w:document xmlns:w="' + W_NS + '" xmlns:r="' + R_NS + '"><w:body>'
      + (story ? '<w:p><w:r><w:t>BodyMark</w:t></w:r></w:p>' : paragraph)
      + '<w:sectPr>' + (story ? '<w:' + story + 'Reference w:type="default" r:id="rStory"/>' : '')
      + '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
      + '</w:sectPr></w:body></w:document>',
  };
  if (story) {
    const root = story === 'header' ? 'hdr' : 'ftr';
    parts['word/' + story + '1.xml'] = '<w:' + root + ' xmlns:w="' + W_NS + '">' + paragraph + '</w:' + root + '>';
  }
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}
