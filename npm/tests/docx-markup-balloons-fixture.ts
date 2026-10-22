import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const DATE = 'w:date="2025-01-02T09:00:00Z"';

/**
 * Original All Markup sample: issue #1300's bold-run and left-indent changes, a comment with a
 * reply, and three comments on one line, on Letter or A4 paper.
 */
export function markupBalloonsFixture(paper: 'letter' | 'a4' = 'letter'): Uint8Array {
  const [width, height] = paper === 'letter' ? [12240, 15840] : [11906, 16838];
  const comment = (id: number) => (text: string) =>
    `<w:commentRangeStart w:id="${id}"/><w:r><w:t>${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:commentReference w:id="${id}"/></w:r>`;
  const body = [
    '<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r><w:r><w:rPr><w:b/>' +
      `<w:rPrChange w:id="11" w:author="Alice Archer" ${DATE}><w:rPr/></w:rPrChange></w:rPr><w:t>buyer</w:t></w:r>` +
      '<w:r><w:t xml:space="preserve"> pays on delivery.</w:t></w:r></w:p>',
    `<w:p><w:pPr><w:ind w:left="720"/><w:pPrChange w:id="12" w:author="Alice Archer" ${DATE}><w:pPr/></w:pPrChange></w:pPr>` +
      '<w:r><w:t>Indented clause.</w:t></w:r></w:p>',
    // The reply shares its root's range, as Word writes it.
    '<w:p><w:r><w:t xml:space="preserve">Risk passes </w:t></w:r><w:commentRangeStart w:id="1"/>' +
      `${comment(0)('on delivery')}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>` +
      '<w:r><w:t xml:space="preserve"> of the goods.</w:t></w:r></w:p>',
    `<w:p>${comment(2)('First')}<w:r><w:t xml:space="preserve"> and </w:t></w:r>${comment(3)('second')}` +
      `<w:r><w:t xml:space="preserve"> and </w:t></w:r>${comment(4)('third')}<w:r><w:t xml:space="preserve"> share a line.</w:t></w:r></w:p>`,
  ].join('');
  const comments: Array<[number, string, string, string]> = [
    [0, 'Alice Archer', 'AA', 'Is delivery the right trigger?'],
    [1, 'Bob Baker', 'BB', 'Yes, it matches the risk clause.'],
    [2, 'Bob Baker', 'BB', 'One.'],
    [3, 'Alice Archer', 'AA', 'Two.'],
    [4, 'Bob Baker', 'BB', 'Three.'],
  ];
  const paraId = (id: number) => `1000000${id}`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings', 'comments', 'commentsExtended'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings', 'comments'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}<Relationship Id="rTopology" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="${width}" w:h="${height}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="BalloonText"><w:name w:val="Balloon Text"/><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="18"/></w:rPr></w:style></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/comments.xml': `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14}">${comments.map(([id, author, initials, text]) => `<w:comment w:id="${id}" w:author="${author}" w:initials="${initials}" ${DATE}><w:p w14:paraId="${paraId(id)}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`).join('')}</w:comments>`,
    'word/commentsExtended.xml': `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="${paraId(0)}"/><w15:commentEx w15:paraId="${paraId(1)}" w15:paraIdParent="${paraId(0)}"/>${[2, 3, 4].map(id => `<w15:commentEx w15:paraId="${paraId(id)}"/>`).join('')}</w15:commentsEx>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
