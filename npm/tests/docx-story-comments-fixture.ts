import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Comment id, the commented text in its story, and the comment's own text (issue #1310). */
export const STORY_COMMENTS = [
  { id: 0, text: 'BODYTEXT', comment: 'BODYCOMMENT' },
  { id: 1, text: 'HEADERTEXT', comment: 'HEADERCOMMENT' },
  { id: 2, text: 'FOOTERTEXT', comment: 'FOOTERCOMMENT' },
  { id: 3, text: 'NOTETEXT', comment: 'NOTECOMMENT' },
  { id: 4, text: 'ENDTEXT', comment: 'ENDCOMMENT' },
] as const;

/**
 * Two pages under one default header and footer. Page 1 holds the body comment and cites a
 * footnote and an endnote; the endnote prints at the end of page 2. Each of the five stories wraps
 * its text in a comment, so a header or footer comment has two pages to repeat on. With `bodyOnly`
 * only the body comment remains: the control.
 */
export function storyCommentsFixture(bodyOnly = false): Uint8Array {
  const commented = (id: number, text: string) => !bodyOnly || id === 0
    ? `<w:commentRangeStart w:id="${id}"/><w:r><w:t>${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/>` +
      `<w:r><w:commentReference w:id="${id}"/></w:r>`
    : `<w:r><w:t>${text}</w:t></w:r>`;
  const separators = (kind: 'footnote' | 'endnote') =>
    `<w:${kind} w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;
  const body = `<w:p>${commented(0, 'BODYTEXT')}<w:r><w:footnoteReference w:id="2"/></w:r>` +
    '<w:r><w:endnoteReference w:id="2"/></w:r></w:p>' +
    '<w:p><w:r><w:br w:type="page"/><w:t>Second page.</w:t></w:r></w:p>';
  const comments = STORY_COMMENTS.filter(({ id }) => !bodyOnly || id === 0);
  const parts: Record<string, string> = {
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}<w:sectPr>` +
      '<w:headerReference w:type="default" r:id="rheader1"/><w:footerReference w:type="default" r:id="rfooter1"/>' +
      '<w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>',
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/header1.xml': `<w:hdr xmlns:w="${W_NS}"><w:p>${commented(1, 'HEADERTEXT')}</w:p></w:hdr>`,
    'word/footer1.xml': `<w:ftr xmlns:w="${W_NS}"><w:p>${commented(2, 'FOOTERTEXT')}</w:p></w:ftr>`,
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}">${separators('footnote')}<w:footnote w:id="2"><w:p>` +
      `<w:r><w:footnoteRef/></w:r>${commented(3, 'NOTETEXT')}</w:p></w:footnote></w:footnotes>`,
    'word/endnotes.xml': `<w:endnotes xmlns:w="${W_NS}">${separators('endnote')}<w:endnote w:id="2"><w:p>` +
      `<w:r><w:endnoteRef/></w:r>${commented(4, 'ENDTEXT')}</w:p></w:endnote></w:endnotes>`,
    'word/comments.xml': `<w:comments xmlns:w="${W_NS}">${comments.map(({ id, comment }) =>
      `<w:comment w:id="${id}" w:author="Reviewer" w:initials="R" w:date="2026-10-01T00:00:00Z">` +
      `<w:p><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment>`).join('')}</w:comments>`,
  };
  const types: Record<string, string> = {
    document: 'document.main', styles: 'styles', settings: 'settings', header1: 'header', footer1: 'footer',
    footnotes: 'footnotes', endnotes: 'endnotes', comments: 'comments',
  };
  const relationshipTypes: Record<string, string> = { header1: 'header', footer1: 'footer' };
  const related = Object.keys(types).filter((name) => name !== 'document');
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      Object.entries(types).map(([name, type]) => `<Override PartName="/word/${name}.xml" ` +
        `ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('') +
      '</Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rDocument" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      related.map((name) => `<Relationship Id="r${name}" Type="${R_NS}/${relationshipTypes[name] ?? name}" ` +
        `Target="${name}.xml"/>`).join('') + '</Relationships>',
    ...parts,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}
