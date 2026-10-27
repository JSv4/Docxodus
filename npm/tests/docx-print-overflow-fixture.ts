import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/**
 * A one-page Letter document of 25 short lines, each carrying a long comment: more than its markup
 * area holds, so most balloons are cut. Comment 12 runs to several paragraphs with blank ones between.
 * Every comment was written at 4:05 PM on its author's clock (w:date), which the modern comment parts
 * record as 21:05 UTC (commentsIds maps each comment's last paragraph to a durable id, and
 * commentsExtensible gives that id its UTC date).
 */
export function printOverflowFixture(): Uint8Array {
  const ids = Array.from({ length: 25 }, (_, id) => id);
  const paraId = (id: number) => (0x10000000 + id).toString(16).toUpperCase();
  const durableId = (id: number) => (0x20000000 + id).toString(16).toUpperCase();
  const remark = (id: number) => `Note ${id}: the warranty period in this section should run from acceptance, not delivery. ` +
    'Please align it with the acceptance certificate and the payment milestones in the schedule. ' +
    'The supplier has asked for twelve months; procurement prefers eighteen. Raise it at the next drafting call.';
  const paragraphsOf = (id: number) =>
    id === 12 ? ['First, the cap.', '', 'Second, the basket.', '', 'Third, the survival period.'] : [remark(id)];
  const run = (text: string) => (text ? `<w:r><w:t>${text}</w:t></w:r>` : '');

  const body = ids.map((id) => `<w:p><w:commentRangeStart w:id="${id}"/>` +
    `<w:r><w:t>Section ${id + 1} applies to the transferred assets.</w:t></w:r>` +
    `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`).join('');
  const comments = ids.map((id) => {
    const [author, initials] = id % 2 ? ['Kwame Asante', 'KA'] : ['Ines Moreau', 'IM'];
    const paragraphs = paragraphsOf(id);
    return `<w:comment w:id="${id}" w:author="${author}" w:initials="${initials}" w:date="2025-06-11T16:05:00Z">` +
      paragraphs.map((text, index) =>
        `<w:p${index === paragraphs.length - 1 ? ` w14:paraId="${paraId(id)}"` : ''}>${run(text)}</w:p>`).join('') +
      '</w:comment>';
  }).join('');

  const parts: Array<[string, string, string]> = [
    ['word/document.xml', `${WML}.document.main+xml`, `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr>` +
      '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" ' +
      'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'],
    ['word/styles.xml', `${WML}.styles+xml`, `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
      '</w:docDefaults><w:style w:type="paragraph" w:styleId="BalloonText"><w:name w:val="Balloon Text"/>' +
      '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="18"/></w:rPr></w:style></w:styles>'],
    ['word/settings.xml', `${WML}.settings+xml`, `<w:settings xmlns:w="${W_NS}"/>`],
    ['word/comments.xml', `${WML}.comments+xml`, `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14}">${comments}</w:comments>`],
    ['word/commentsIds.xml', `${WML}.commentsIds+xml`, `<w16cid:commentsIds xmlns:w16cid="${W16CID}">${ids.map((id) =>
      `<w16cid:commentId w16cid:paraId="${paraId(id)}" w16cid:durableId="${durableId(id)}"/>`).join('')}</w16cid:commentsIds>`],
    ['word/commentsExtensible.xml', `${WML}.commentsExtensible+xml`, `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}">` +
      ids.map((id) => `<w16cex:commentExtensible w16cex:durableId="${durableId(id)}" w16cex:dateUtc="2025-06-11T21:05:00Z"/>`).join('') +
      '</w16cex:commentsExtensible>'],
  ];
  const relationships: Array<[string, string]> = [
    [`${R_NS}/styles`, 'styles.xml'],
    [`${R_NS}/settings`, 'settings.xml'],
    [`${R_NS}/comments`, 'comments.xml'],
    ['http://schemas.microsoft.com/office/2016/09/relationships/commentsIds', 'commentsIds.xml'],
    ['http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible', 'commentsExtensible.xml'],
  ];
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      parts.map(([path, type]) => `<Override PartName="/${path}" ContentType="${type}"/>`).join('') + '</Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relationships.map(([type, target], index) => `<Relationship Id="r${index}" Type="${type}" Target="${target}"/>`).join('') +
      '</Relationships>'],
    ...parts.map(([path, , text]) => [path, text] as [string, string]),
  ].map(([name, text]) => ({ name, data: xml(text) })));
}
