import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const DATE = 'w:date="2025-01-02T09:00:00Z"';

/** A comment for `balloonDocx`; `parent` makes it a reply. */
export interface FixtureComment { id: number; author: string; initials: string; text: string; parent?: number }

/** Text wrapped in comment `id`'s range, followed by its reference. */
export const commented = (id: number, text: string) =>
  `<w:commentRangeStart w:id="${id}"/><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/>` +
  `<w:r><w:commentReference w:id="${id}"/></w:r>`;

/** A run whose formatting changed from nothing to `now`. */
export const formatted = (id: number, text: string, now = '<w:b/>') =>
  `<w:r><w:rPr>${now}<w:rPrChange w:id="${id}" w:author="Alice Archer" ${DATE}><w:rPr/></w:rPrChange></w:rPr>` +
  `<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A paragraph whose properties changed from nothing to `now`. */
export const reformatted = (id: number, text: string, now = '<w:ind w:left="720"/>') =>
  `<w:p><w:pPr>${now}<w:pPrChange w:id="${id}" w:author="Alice Archer" ${DATE}><w:pPr/></w:pPrChange></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

export const plain = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** A paragraph holding `content` whose own mark is a tracked insertion or deletion. */
export const markedParagraph = (id: number, mark: 'ins' | 'del', content: string) =>
  `<w:p><w:pPr><w:rPr><w:${mark} w:id="${id}" w:author="Alice Archer" ${DATE}/></w:rPr></w:pPr>${content}</w:p>`;

/** An original one-section document with the given body, comments and optional default header. */
export function balloonDocx(options: {
  body: string; comments?: FixtureComment[]; header?: string; paper?: 'letter' | 'a4';
}): Uint8Array {
  const [width, height] = options.paper === 'a4' ? [11906, 16838] : [12240, 15840];
  const comments = options.comments ?? [];
  const paraId = (id: number) => (0x10000000 + id).toString(16).toUpperCase();
  const parts: Record<string, string> = {
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${options.body}<w:sectPr>` +
      `${options.header ? '<w:headerReference w:type="default" r:id="rheader"/>' : ''}<w:pgSz w:w="${width}" w:h="${height}"/>` +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>',
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="BalloonText"><w:name w:val="Balloon Text"/><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="18"/></w:rPr></w:style></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/comments.xml': `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14}">${comments.map(({ id, author, initials, text }) =>
      `<w:comment w:id="${id}" w:author="${author}" w:initials="${initials}" ${DATE}><w:p w14:paraId="${paraId(id)}">` +
      `<w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`).join('')}</w:comments>`,
    'word/commentsExtended.xml': `<w15:commentsEx xmlns:w15="${W15}">${comments.map(({ id, parent }) =>
      `<w15:commentEx w15:paraId="${paraId(id)}"${parent === undefined ? '' : ` w15:paraIdParent="${paraId(parent)}"`}/>`).join('')}</w15:commentsEx>`,
    ...(options.header ? { 'word/header1.xml': `<w:hdr xmlns:w="${W_NS}">${options.header}</w:hdr>` } : {}),
  };
  const overrides = Object.keys(parts).map((path) => {
    const part = path.slice('word/'.length, -'.xml'.length);
    const type = part === 'document' ? 'document.main' : part === 'header1' ? 'header' : part;
    return `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
  }).join('');
  const relationships = ['styles', 'settings', 'comments'].map((p) =>
    `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('') +
    `<Relationship Id="rTopology" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>` +
    (options.header ? `<Relationship Id="rheader" Type="${R_NS}/header" Target="header1.xml"/>` : '');
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    ...parts,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

/**
 * Original All Markup sample: issue #1300's bold-run and left-indent changes, a comment with a
 * reply, and three comments on one line, on Letter or A4 paper.
 */
export function markupBalloonsFixture(paper: 'letter' | 'a4' = 'letter'): Uint8Array {
  return balloonDocx({
    paper,
    body: `<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r>${formatted(11, 'buyer')}` +
      '<w:r><w:t xml:space="preserve"> pays on delivery.</w:t></w:r></w:p>' +
      reformatted(12, 'Indented clause.') +
      // The reply shares its root's range, as Word writes it.
      '<w:p><w:r><w:t xml:space="preserve">Risk passes </w:t></w:r><w:commentRangeStart w:id="1"/>' +
      `${commented(0, 'on delivery')}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>` +
      '<w:r><w:t xml:space="preserve"> of the goods.</w:t></w:r></w:p>' +
      `<w:p>${commented(2, 'First')}<w:r><w:t xml:space="preserve"> and </w:t></w:r>${commented(3, 'second')}` +
      `<w:r><w:t xml:space="preserve"> and </w:t></w:r>${commented(4, 'third')}<w:r><w:t xml:space="preserve"> share a line.</w:t></w:r></w:p>`,
    comments: [
      { id: 0, author: 'Alice Archer', initials: 'AA', text: 'Is delivery the right trigger?' },
      { id: 1, author: 'Bob Baker', initials: 'BB', text: 'Yes, it matches the risk clause.', parent: 0 },
      { id: 2, author: 'Bob Baker', initials: 'BB', text: 'One.' },
      { id: 3, author: 'Alice Archer', initials: 'AA', text: 'Two.' },
      { id: 4, author: 'Bob Baker', initials: 'BB', text: 'Three.' },
    ],
  });
}

/** A point comment, an empty-range comment and a ranged comment on one line. */
export function anchorlessCommentsFixture(): Uint8Array {
  return balloonDocx({
    body: '<w:p><w:r><w:t xml:space="preserve">A point </w:t></w:r><w:r><w:commentReference w:id="0"/></w:r>' +
      '<w:r><w:t xml:space="preserve"> an empty range </w:t></w:r><w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/>' +
      `<w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r>${commented(2, 'a range')}</w:p>`,
    comments: [
      { id: 0, author: 'Alice Archer', initials: 'AA', text: 'PointNote' },
      { id: 1, author: 'Bob Baker', initials: 'BB', text: 'EmptyNote' },
      { id: 2, author: 'Alice Archer', initials: 'AA', text: 'RangeNote' },
    ],
  });
}

/** `count` long comments, `perLine` of them on each line: more than the markup area can hold. */
export function crowdedCommentsFixture(count: number, perLine: number, text: string): Uint8Array {
  const ids = Array.from({ length: count }, (_, id) => id);
  const lines = Array.from({ length: Math.ceil(count / perLine) }, (_, line) =>
    `<w:p>${ids.slice(line * perLine, (line + 1) * perLine).map((id) => commented(id, `w${id} `)).join('')}</w:p>`);
  return balloonDocx({
    body: lines.join(''),
    comments: ids.map((id) => ({ id, author: id % 2 ? 'Bob Baker' : 'Alice Archer', initials: id % 2 ? 'BB' : 'AA', text: `${id}: ${text}` })),
  });
}
