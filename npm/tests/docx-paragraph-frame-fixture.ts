import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original text, a positioned frame and an ordinary centered anchor paragraph.
export function paragraphFrameFixture(framed = true, story: 'body' | 'header' | 'footer' = 'footer', multiple = false, table = false): Uint8Array {
  const type = story === 'header' ? 'header' : 'footer';
  const part = `${type}1`;
  const frameProperties = framed ? '<w:framePr w:w="10080" w:hAnchor="page" w:vAnchor="text" w:x="840" w:y="20" w:wrap="auto"/>' : '';
  const paragraphs = ['CEDAR RECORD 731', ...(multiple ? ['CEDAR FRAME SECOND'] : [])]
    .map(text => `<w:p><w:pPr>${frameProperties}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`).join('')
    + '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Cedar page control</w:t></w:r></w:p>';
  const bodyParagraphs = table ? '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' + paragraphs + '</w:tc></w:tr></w:tbl>' : paragraphs;
  const namedParts = ['document', 'styles', 'settings', ...(story === 'body' ? [] : [part])];
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${namedParts.map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p === part ? type : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${namedParts.filter(p => p !== 'document').map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p === part ? type : p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${story === 'body' ? bodyParagraphs : '<w:p><w:r><w:t>Cedar body control</w:t></w:r></w:p>'}<w:sectPr>${story === 'body' ? '' : `<w:${type}Reference w:type="default" r:id="r${part}"/>`}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  };
  if (story !== 'body') parts[`word/${part}.xml`] = `<w:${type === 'header' ? 'hdr' : 'ftr'} xmlns:w="${W_NS}">${paragraphs}</w:${type === 'header' ? 'hdr' : 'ftr'}>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
