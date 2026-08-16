const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

export const REVIEW_TOKENS = Object.freeze({
  stable: "PROFILE_STABLE",
  original: "PROFILE_ORIGINAL",
  final: "PROFILE_FINAL",
  moved: "PROFILE_MOVED",
  format: "PROFILE_FORMAT",
  rootComment: "Root printable review comment.",
  replyComment: "Ordered printable reply.",
  commentOriginal: "CMT_OLD",
  commentFinal: "CMT_NEW",
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function xml(value) {
  return Buffer.from(value, "utf8");
}

function storedZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const checksum = crc32(entry.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(entry.data.length, 18);
    local.writeUInt32LE(entry.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    localParts.push(local, name, entry.data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(entry.data.length, 20);
    central.writeUInt32LE(entry.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, name);
    offset += local.length + name.length + entry.data.length;
  }
  const directory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...localParts, directory, end]));
}

/** A compact, readable package for the supported Node/PDF profile matrix. */
export function generateReviewCommentDocx({
  unsupportedParagraphChange = false,
  unsupportedRevisionFamily = false,
  malformedCommentTopology = false,
  duplicateCommentIdentity = false,
  oversizedMarginComment = false,
  relocatedCommentsExtended = false,
} = {}) {
  const paragraphChange = unsupportedParagraphChange
    ? `<w:pPrChange w:id="40" w:author="Paragraph Reviewer" w:date="2026-08-03T12:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`
    : "";
  const unsupportedRange = unsupportedRevisionFamily
    ? `<w:customXmlMoveFromRangeStart w:id="50" w:author="Unsupported Reviewer" w:date="2026-08-03T12:00:00Z"/><w:r><w:t>UNSUPPORTED CUSTOM XML MOVE</w:t></w:r><w:customXmlMoveFromRangeEnd w:id="50"/>`
    : "";
  const duplicateMetadata = duplicateCommentIdentity
    ? `\n  <w15:commentEx w15:paraId="10000001" w15:done="0"/>`
    : "";
  const commentsExtended = malformedCommentTopology
    ? `<w15:commentEx w15:paraId="10000001" w15:paraIdParent="10000002" w15:done="1"/>
  <w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001"/>
  <w15:commentEx w15:paraId="10000003" w15:paraIdParent="99999999"/>`
    : `<w15:commentEx w15:paraId="10000001" w15:done="1"/>
  <w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001"/>${duplicateMetadata}`;
  const malformedComment = malformedCommentTopology
    ? `<w:comment w:id="2" w:author="Orphan Reply" w:date="2026-08-01T10:00:00Z"><w:p w14:paraId="10000003"><w:r><w:t>Orphaned printable reply.</w:t></w:r></w:p></w:comment>`
    : "";
  const malformedCommentReference = malformedCommentTopology
    ? `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="2"/></w:r>`
    : "";
  const duplicateComments = duplicateCommentIdentity
    ? `<w:comment w:id="0" w:author="Duplicate Id" w:date="2026-08-01T10:00:00Z"><w:p w14:paraId="10000004"><w:r><w:t>Duplicate id definition.</w:t></w:r></w:p></w:comment>
  <w:comment w:id="2" w:author="Duplicate Paragraph" w:date="2026-08-01T11:00:00Z"><w:p w14:paraId="10000001"><w:r><w:t>Duplicate paragraph identity.</w:t></w:r></w:p></w:comment>`
    : "";
  const marginOverflow = oversizedMarginComment
    ? ` ${Array.from({ length: 420 }, (_, index) => `MARGIN_OVERFLOW_${String(index).padStart(3, "0")}`).join(" ")}`
    : "";
  const commentsExtendedEntry = relocatedCommentsExtended
    ? "word/custom/reviewTopology.xml"
    : "word/commentsExtended.xml";
  const commentsExtendedTarget = relocatedCommentsExtended
    ? "custom/reviewTopology.xml"
    : "commentsExtended.xml";
  const entries = [
    {
      name: "[Content_Types].xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/${commentsExtendedEntry}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/_rels/document.xml.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>
  <Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>
  <Relationship Id="rIdCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="${commentsExtendedTarget}"/>
</Relationships>`),
    },
    {
      name: "word/styles.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="comment text"/></w:style>
  <w:style w:type="character" w:styleId="CommentReference"><w:name w:val="comment reference"/></w:style>
</w:styles>`),
    },
    {
      name: "word/comments.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W}" xmlns:w14="${W14}">
  <w:comment w:id="0" w:author="Alice Root" w:date="2026-08-01T08:00:00Z"><w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">${REVIEW_TOKENS.rootComment} </w:t></w:r><w:del w:id="60" w:author="Comment Reviewer" w:date="2026-08-15T11:22:33Z"><w:r><w:delText>${REVIEW_TOKENS.commentOriginal}</w:delText></w:r></w:del><w:ins w:id="61" w:author="Comment Reviewer" w:date="2026-08-15T11:22:33Z"><w:r><w:t>${REVIEW_TOKENS.commentFinal}</w:t></w:r></w:ins><w:r><w:t>${marginOverflow}</w:t></w:r></w:p></w:comment>
  <w:comment w:id="1" w:author="Bob Reply" w:date="2026-08-01T09:00:00Z"><w:p w14:paraId="10000002"><w:r><w:t>${REVIEW_TOKENS.replyComment}</w:t></w:r></w:p></w:comment>
  ${malformedComment}
  ${duplicateComments}
</w:comments>`),
    },
    {
      name: commentsExtendedEntry,
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="${W15}">
  ${commentsExtended}
</w15:commentsEx>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>
  <w:p><w:pPr>${paragraphChange}</w:pPr><w:r><w:t xml:space="preserve">${REVIEW_TOKENS.stable} </w:t></w:r>
    ${unsupportedRange}
    <w:commentRangeStart w:id="0"/>
    <w:del w:id="1" w:author="Profile Reviewer" w:date="2026-08-16T12:34:56Z"><w:r><w:delText>${REVIEW_TOKENS.original}</w:delText></w:r></w:del>
    <w:ins w:id="2" w:author="Profile Reviewer" w:date="2026-08-16T12:34:56Z"><w:r><w:t>${REVIEW_TOKENS.final}</w:t></w:r></w:ins>
    <w:commentRangeEnd w:id="0"/>
    <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
    ${malformedCommentReference}
  </w:p>
  <w:p><w:r><w:t xml:space="preserve">MOVE SOURCE </w:t></w:r><w:moveFromRangeStart w:id="20" w:name="node-profile-move"/><w:moveFrom w:id="21" w:author="Profile Reviewer" w:date="2026-08-16T12:34:56Z"><w:r><w:t>${REVIEW_TOKENS.moved}</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="20"/></w:p>
  <w:p><w:r><w:t xml:space="preserve">MOVE DESTINATION </w:t></w:r><w:moveToRangeStart w:id="22" w:name="node-profile-move"/><w:moveTo w:id="23" w:author="Profile Reviewer" w:date="2026-08-16T12:34:56Z"><w:r><w:t>${REVIEW_TOKENS.moved}</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="22"/></w:p>
  <w:p><w:r><w:rPr><w:b/><w:color w:val="0000FF"/><w:rPrChange w:id="30" w:author="Profile Reviewer" w:date="2026-08-16T12:34:56Z"><w:rPr><w:b w:val="0"/><w:color w:val="FF0000"/></w:rPr></w:rPrChange></w:rPr><w:t>${REVIEW_TOKENS.format}</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr>
</w:body></w:document>`),
    },
  ];
  return storedZip(entries);
}
