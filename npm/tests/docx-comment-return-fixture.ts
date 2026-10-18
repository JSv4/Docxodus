import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original comment range and ordinary bookmark link; the optional page break separates
// the first painted note from its one real return target.
export function commentReturnFixture(comments = true, later = false): Uint8Array {
  const mark = comments ? '<w:r><w:commentReference w:id="6"/></w:r>' : '';
  const ending = comments ? '<w:commentRangeEnd w:id="6"/>' + mark : '';
  const body = `<w:p>${comments ? '<w:commentRangeStart w:id="6"/>' : ''}<w:r><w:t>Maple navigation anchor.</w:t></w:r>${later ? '' : ending}</w:p>` +
    (later ? `<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>Maple return target.</w:t></w:r>${ending}</w:p>` : '') +
    '<w:p><w:bookmarkStart w:id="12" w:name="MapleDestination"/><w:r><w:t>Maple bookmark destination.</w:t></w:r><w:bookmarkEnd w:id="12"/></w:p>' +
    '<w:p><w:hyperlink w:anchor="MapleDestination"><w:r><w:t>MapleBodyReturn</w:t></w:r></w:hyperlink></w:p>';
  const names = ['document', 'styles', 'settings', ...(comments ? ['comments'] : [])];
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${names.map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name === 'document' ? 'document.main' : name}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDocument" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.filter(name => name !== 'document').map(name => `<Relationship Id="r${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  };
  if (comments) parts['word/comments.xml'] = `<w:comments xmlns:w="${W_NS}"><w:comment w:id="6" w:author=""><w:p><w:r><w:t>MapleNote Return to the marked text.</w:t></w:r></w:p></w:comment></w:comments>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
