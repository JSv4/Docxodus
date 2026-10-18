import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original thread and flat control; only the optional topology distinguishes them.
export function commentThreadFixture(threaded = true, longAuthor = true, resolved = false): Uint8Array {
  const author = longAuthor ? 'Cedar Archivistwithalongunbrokentoken' : 'Cy';
  const ids = ['00000010', '00000020', '00000030'];
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings', 'comments', 'commentsExtended'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings', 'comments'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}<Relationship Id="rTopology" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Cedar thread anchor.</w:t></w:r><w:commentRangeEnd w:id="0"/>${ids.map((_, i) => `<w:r><w:commentReference w:id="${i}"/></w:r>`).join('')}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/comments.xml': `<w:comments xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">${ids.map((id, i) => `<w:comment w:id="${i}" w:author="${author}" w:date="2026-09-10T12:00:00Z"><w:p w14:paraId="${id}"><w:r><w:t>CedarThread${i} Check the archive tray.</w:t></w:r></w:p></w:comment>`).join('')}</w:comments>`,
    'word/commentsExtended.xml': `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${ids.map((id, i) => `<w15:commentEx w15:paraId="${id}"${resolved ? ' w15:done="1"' : ''}${threaded && i > 0 ? ` w15:paraIdParent="${ids[i - 1]}"` : ''}/>`).join('')}</w15:commentsEx>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
