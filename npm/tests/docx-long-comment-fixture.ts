import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original long-note and short-note control text, independent of source documents.
export const LONG_COMMENT_TOKEN = 'CedarIdentifier' + 'X'.repeat(160);

export function longCommentFixture(count = 240, secondPage = false, longWord = false): Uint8Array {
  const note = Array.from({ length: count }, (_, i) => `CedarNote${String(i).padStart(3, '0')}`).join(' ') + (longWord ? ' ' + LONG_COMMENT_TOKEN : '');
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings', 'comments'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings', 'comments'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Cedar inventory note anchor.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>${secondPage ? '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Cedar second page control.</w:t></w:r></w:p>' : ''}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/comments.xml': `<w:comments xmlns:w="${W_NS}"><w:comment w:id="0" w:author="Cedar Reviewer" w:date="2026-09-10T12:00:00Z"><w:p><w:r><w:t>${note}</w:t></w:r></w:p></w:comment></w:comments>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
