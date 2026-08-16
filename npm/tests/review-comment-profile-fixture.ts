import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const REVIEW_AUTHOR = 'Profile Reviewer';
const REVIEW_DATE = '2026-08-16T12:34:56Z';

export const REVIEW_PROFILE_TOKENS = Object.freeze({
  stable: 'BODY_STABLE',
  original: 'BODY_ORIGINAL',
  final: 'BODY_FINAL',
  insertedCommentTarget: 'INSERTION_COMMENT_TARGET',
  deletedCommentTarget: 'DELETION_COMMENT_TARGET',
  move: 'MOVED_CONTENT',
  moveSourceContext: 'SOURCE_LOCATION',
  moveDestinationContext: 'DESTINATION_LOCATION',
  format: 'FORMAT_TARGET',
});

export const COMMENT_REVISION_TOKENS = Object.freeze({
  original: 'COMMENT_BODY_ORIGINAL',
  final: 'COMMENT_BODY_FINAL',
});

export const STORY_PROFILE_TOKENS = Object.freeze({
  body: { original: 'BODY_STORY_ORIGINAL', final: 'BODY_STORY_FINAL' },
  header: { original: 'HEADER_STORY_ORIGINAL', final: 'HEADER_STORY_FINAL' },
  footer: { original: 'FOOTER_STORY_ORIGINAL', final: 'FOOTER_STORY_FINAL' },
  footnote: { original: 'FOOTNOTE_STORY_ORIGINAL', final: 'FOOTNOTE_STORY_FINAL' },
  endnote: { original: 'ENDNOTE_STORY_ORIGINAL', final: 'ENDNOTE_STORY_FINAL' },
});

export interface ExpectedComment {
  id: string;
  author: string;
  date: string;
  body: string;
  parentId?: string;
  resolved: 'true' | 'false' | 'unknown';
  story: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'reply';
}

export const EXPECTED_PROFILE_COMMENTS: readonly ExpectedComment[] = Object.freeze([
  {
    id: '0', author: 'Alice Root', date: '2026-08-01T08:00:00Z',
    body: 'Root overlap comment.', resolved: 'false', story: 'body',
  },
  {
    id: '1', author: 'Bob Reply', date: '2026-08-01T09:00:00Z',
    body: 'Ordered reply body.', parentId: '0', resolved: 'unknown', story: 'reply',
  },
  {
    id: '2', author: 'Carol Resolved', date: '2026-08-01T10:00:00Z',
    body: 'Inserted-only resolved comment.', resolved: 'true', story: 'body',
  },
  {
    id: '3', author: 'Dan Unknown', date: '2026-08-01T11:00:00Z',
    body: 'Deleted-only unknown comment.', resolved: 'unknown', story: 'body',
  },
  {
    id: '10', author: 'Body Reviewer', date: '2026-08-02T07:00:00Z',
    body: 'Body story comment.', resolved: 'false', story: 'body',
  },
  {
    id: '11', author: 'Header Reviewer', date: '2026-08-02T08:00:00Z',
    body: 'Header story comment.', resolved: 'false', story: 'header',
  },
  {
    id: '12', author: 'Footer Reviewer', date: '2026-08-02T09:00:00Z',
    body: 'Footer story comment.', resolved: 'false', story: 'footer',
  },
  {
    id: '13', author: 'Footnote Reviewer', date: '2026-08-02T10:00:00Z',
    body: 'Footnote story comment.', resolved: 'false', story: 'footnote',
  },
  {
    id: '14', author: 'Endnote Reviewer', date: '2026-08-02T11:00:00Z',
    body: 'Endnote story comment.', resolved: 'false', story: 'endnote',
  },
]);

const styles = xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/></w:style>
  <w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/></w:style>
  <w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="comment text"/></w:style>
  <w:style w:type="character" w:styleId="CommentReference"><w:name w:val="comment reference"/></w:style>
</w:styles>`);

function revisionPair(original: string, final: string, id: number): string {
  return `<w:del w:id="${id}" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}">`
    + `<w:r><w:delText>${original}</w:delText></w:r></w:del>`
    + `<w:ins w:id="${id + 1}" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}">`
    + `<w:r><w:t>${final}</w:t></w:r></w:ins>`;
}

function commentReference(id: number): string {
  return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>`
    + `<w:commentReference w:id="${id}"/></w:r>`;
}

function bodyContent(includeStoryProbe: boolean): string {
  const storyProbe = includeStoryProbe
    ? `<w:p><w:r><w:t xml:space="preserve">BODY_STORY: </w:t></w:r>`
      + `<w:commentRangeStart w:id="10"/>`
      + revisionPair(STORY_PROFILE_TOKENS.body.original, STORY_PROFILE_TOKENS.body.final, 100)
      + `<w:commentRangeEnd w:id="10"/>${commentReference(10)}</w:p>`
    : '';
  return `
  <w:p><w:r><w:t xml:space="preserve">${REVIEW_PROFILE_TOKENS.stable} </w:t></w:r>
    <w:commentRangeStart w:id="0"/>
    ${revisionPair(REVIEW_PROFILE_TOKENS.original, REVIEW_PROFILE_TOKENS.final, 1)}
    <w:commentRangeEnd w:id="0"/>${commentReference(0)}
  </w:p>
  <w:p><w:r><w:t xml:space="preserve">INSERT_ONLY: </w:t></w:r>
    <w:commentRangeStart w:id="2"/>
    <w:ins w:id="3" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"><w:r><w:t>${REVIEW_PROFILE_TOKENS.insertedCommentTarget}</w:t></w:r></w:ins>
    <w:commentRangeEnd w:id="2"/>${commentReference(2)}
  </w:p>
  <w:p><w:r><w:t xml:space="preserve">DELETE_ONLY: </w:t></w:r>
    <w:commentRangeStart w:id="3"/>
    <w:del w:id="4" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"><w:r><w:delText>${REVIEW_PROFILE_TOKENS.deletedCommentTarget}</w:delText></w:r></w:del>
    <w:commentRangeEnd w:id="3"/>${commentReference(3)}
  </w:p>
  <w:p><w:r><w:t xml:space="preserve">${REVIEW_PROFILE_TOKENS.moveSourceContext}: </w:t></w:r>
    <w:moveFromRangeStart w:id="20" w:name="profile-move" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"/>
    <w:moveFrom w:id="21" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"><w:r><w:t>${REVIEW_PROFILE_TOKENS.move}</w:t></w:r></w:moveFrom>
    <w:moveFromRangeEnd w:id="20"/>
  </w:p>
  <w:p><w:r><w:t xml:space="preserve">${REVIEW_PROFILE_TOKENS.moveDestinationContext}: </w:t></w:r>
    <w:moveToRangeStart w:id="22" w:name="profile-move" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"/>
    <w:moveTo w:id="23" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"><w:r><w:t>${REVIEW_PROFILE_TOKENS.move}</w:t></w:r></w:moveTo>
    <w:moveToRangeEnd w:id="22"/>
  </w:p>
  <w:p><w:r><w:rPr><w:b/><w:color w:val="0000FF"/>
    <w:rPrChange w:id="30" w:author="${REVIEW_AUTHOR}" w:date="${REVIEW_DATE}"><w:rPr><w:b w:val="0"/><w:color w:val="FF0000"/></w:rPr></w:rPrChange>
    </w:rPr><w:t>${REVIEW_PROFILE_TOKENS.format}</w:t></w:r></w:p>
  ${storyProbe}`;
}

function includedComments(includeStoryComments: boolean): readonly ExpectedComment[] {
  return EXPECTED_PROFILE_COMMENTS.filter((comment) =>
    includeStoryComments || Number(comment.id) < 10);
}

function commentsXml(includeStoryComments: boolean, longRootComment = false): Buffer {
  const overflowProbe = Array.from({ length: 420 }, (_, index) =>
    ` margin-overflow-token-${index}`).join('');
  const comments = includedComments(includeStoryComments)
    .map((comment, index) => `<w:comment w:id="${comment.id}" w:author="${comment.author}" w:date="${comment.date}" w:initials="C${index}">`
      + `<w:p w14:paraId="${(0x10000001 + index).toString(16).toUpperCase()}"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>`
      + `<w:r><w:t>${comment.body}${longRootComment && comment.id === '0' ? overflowProbe : ''}</w:t></w:r>`
      + (comment.id === '0'
        ? revisionPair(COMMENT_REVISION_TOKENS.original, COMMENT_REVISION_TOKENS.final, 200)
        : '')
      + `</w:p></w:comment>`)
    .join('');
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">${comments}</w:comments>`);
}

function commentsExtendedXml(includeStoryComments: boolean): Buffer {
  const comments = includedComments(includeStoryComments);
  const paraIdByComment = new Map(comments.map((comment, index) => [
    comment.id,
    (0x10000001 + index).toString(16).toUpperCase(),
  ]));
  const entries = comments
    // Omitted entries and omitted w15:done values deliberately exercise "unknown".
    .filter((comment) => comment.resolved !== 'unknown' || comment.parentId !== undefined)
    .map((comment) => {
      const parent = comment.parentId
        ? ` w15:paraIdParent="${paraIdByComment.get(comment.parentId)}"`
        : '';
      const done = comment.resolved === 'unknown' ? '' : ` w15:done="${comment.resolved === 'true' ? '1' : '0'}"`;
      return `<w15:commentEx w15:paraId="${paraIdByComment.get(comment.id)}"${parent}${done}/>`;
    })
    .join('');
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="${W15_NS}">${entries}</w15:commentsEx>`);
}

function contentTypes(includeStories: boolean): Buffer {
  const storyOverrides = includeStories ? `
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>` : '';
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>${storyOverrides}
</Types>`);
}

function packageRelationships(): Buffer {
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`);
}

function documentRelationships(includeStories: boolean): Buffer {
  const storyRelationships = includeStories ? `
  <Relationship Id="rIdHeader" Type="${R_NS}/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter" Type="${R_NS}/footer" Target="footer1.xml"/>
  <Relationship Id="rIdFootnotes" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdEndnotes" Type="${R_NS}/endnotes" Target="endnotes.xml"/>` : '';
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rIdComments" Type="${R_NS}/comments" Target="comments.xml"/>
  <Relationship Id="rIdCommentsExtended" Type="${COMMENTS_EXTENDED_REL}" Target="commentsExtended.xml"/>${storyRelationships}
</Relationships>`);
}

function documentXml(includeStories: boolean): Buffer {
  const noteReferences = includeStories
    ? `<w:p><w:r><w:t xml:space="preserve">NOTE_REFERENCES </w:t></w:r>`
      + `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>`
      + `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/></w:r></w:p>`
    : '';
  const sectionReferences = includeStories
    ? `<w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/>`
    : '';
  const repeatedPageProbe = includeStories
    ? `<w:p><w:r><w:t>PAGE_TWO</w:t><w:br w:type="page"/></w:r></w:p>`
      + `<w:p><w:r><w:t>PAGE_THREE</w:t><w:br w:type="page"/></w:r></w:p>`
      + `<w:p><w:r><w:t>PAGE_FOUR</w:t></w:r></w:p>`
    : '';
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${bodyContent(includeStories)}
  ${noteReferences}
  ${repeatedPageProbe}
  <w:sectPr>${sectionReferences}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360"/></w:sectPr>
</w:body></w:document>`);
}

function storyPart(kind: 'header' | 'footer', commentId: number, revisionId: number): Buffer {
  const tokens = STORY_PROFILE_TOKENS[kind];
  const root = kind === 'header' ? 'hdr' : 'ftr';
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${root} xmlns:w="${W_NS}"><w:p><w:pPr><w:pStyle w:val="${kind === 'header' ? 'Header' : 'Footer'}"/></w:pPr>
  <w:r><w:t xml:space="preserve">${kind.toUpperCase()}_STORY: </w:t></w:r><w:commentRangeStart w:id="${commentId}"/>
  ${revisionPair(tokens.original, tokens.final, revisionId)}
  <w:commentRangeEnd w:id="${commentId}"/>${commentReference(commentId)}
</w:p></w:${root}>`);
}

function notesPart(kind: 'footnote' | 'endnote', commentId: number, revisionId: number): Buffer {
  const isFootnote = kind === 'footnote';
  const plural = isFootnote ? 'footnotes' : 'endnotes';
  const singular = isFootnote ? 'footnote' : 'endnote';
  const style = isFootnote ? 'FootnoteText' : 'EndnoteText';
  const reference = isFootnote ? 'footnoteRef' : 'endnoteRef';
  const tokens = STORY_PROFILE_TOKENS[kind];
  return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${plural} xmlns:w="${W_NS}">
  <w:${singular} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${singular}>
  <w:${singular} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${singular}>
  <w:${singular} w:id="1"><w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>
    <w:r><w:${reference}/></w:r><w:r><w:t xml:space="preserve"> ${kind.toUpperCase()}_STORY: </w:t></w:r>
    <w:commentRangeStart w:id="${commentId}"/>${revisionPair(tokens.original, tokens.final, revisionId)}
    <w:commentRangeEnd w:id="${commentId}"/>${commentReference(commentId)}
  </w:p></w:${singular}>
</w:${plural}>`);
}

function buildFixture(includeStories: boolean, longRootComment = false): Uint8Array {
  const entries = [
    { name: '[Content_Types].xml', data: contentTypes(includeStories) },
    { name: '_rels/.rels', data: packageRelationships() },
    { name: 'word/_rels/document.xml.rels', data: documentRelationships(includeStories) },
    { name: 'word/styles.xml', data: styles },
    { name: 'word/comments.xml', data: commentsXml(includeStories, longRootComment) },
    { name: 'word/commentsExtended.xml', data: commentsExtendedXml(includeStories) },
    { name: 'word/document.xml', data: documentXml(includeStories) },
  ];
  if (includeStories) {
    entries.push(
      { name: 'word/header1.xml', data: storyPart('header', 11, 110) },
      { name: 'word/footer1.xml', data: storyPart('footer', 12, 120) },
      { name: 'word/footnotes.xml', data: notesPart('footnote', 13, 130) },
      { name: 'word/endnotes.xml', data: notesPart('endnote', 14, 140) },
    );
  }
  return storedZip(entries);
}

/** Small body-only fixture for focused debugging of intersecting revision/comment ranges. */
export function generateReviewCommentOverlapDocx(): Uint8Array {
  return buildFixture(false);
}

/** Acceptance fixture: the overlap probe plus revision/comment ranges in every supported story. */
export function generateReviewCommentStoriesDocx(): Uint8Array {
  return buildFixture(true);
}

/** Oversized thread used to prove finite page margins fail closed instead of clipping. */
export function generateLongMarginCommentDocx(): Uint8Array {
  return buildFixture(false, true);
}
