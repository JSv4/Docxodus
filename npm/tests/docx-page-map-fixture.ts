import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

const CONTENT_TYPES = (partName: string, contentType: string) => xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="${partName}" ContentType="${contentType}"/>
</Types>`);

const PACKAGE_RELS = xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`);

const PAGE = `
  <w:sectPr>
    <w:pgSz w:w="9000" w:h="7200"/>
    <w:pgMar w:top="720" w:right="1800" w:bottom="720" w:left="720"
      w:header="360" w:footer="360" w:gutter="0"/>
  </w:sectPr>`;

const STYLES = xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="comment text"/></w:style>
  <w:style w:type="character" w:styleId="CommentReference"><w:name w:val="comment reference"/></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
</w:styles>`);

/** A two-paragraph native comment whose range lives in a table cell. */
export function generateTableCommentDocx(collapsed = false): Uint8Array {
  const commentedBody = collapsed
    ? `<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/>`
    : `<w:commentRangeStart w:id="0"/><w:r><w:t>Cell comment target</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>`;
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: CONTENT_TYPES(
        '/word/comments.xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      ),
    },
    { name: '_rels/.rels', data: PACKAGE_RELS },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/comments" Target="comments.xml"/>
</Relationships>`),
    },
    { name: 'word/styles.xml', data: STYLES },
    {
      name: 'word/comments.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}">
  <w:comment w:id="0" w:author="Reviewer" w:initials="RV">
    <w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
      <w:r><w:annotationRef/></w:r><w:r><w:t>First comment paragraph.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
      <w:r><w:t>Second comment paragraph.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  <w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tr><w:tc>
    <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
    <w:p>${commentedBody}<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
        <w:commentReference w:id="0"/></w:r></w:p>
  </w:tc></w:tr></w:tbl>
  <w:p><w:r><w:t>Following body paragraph.</w:t></w:r></w:p>
  ${PAGE}
</w:body></w:document>`),
    },
  ]);
}

/** A real native endnote with one deliberately oversized text paragraph. */
export function generateLongEndnoteDocx(wordCount = 1200): Uint8Array {
  const words = Array.from({ length: wordCount }, (_, index) => `endnote${index}`).join(' ');
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: CONTENT_TYPES(
        '/word/endnotes.xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
      ),
    },
    { name: '_rels/.rels', data: PACKAGE_RELS },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/endnotes" Target="endnotes.xml"/>
</Relationships>`),
    },
    { name: 'word/styles.xml', data: STYLES },
    {
      name: 'word/endnotes.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="${W_NS}">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> ${words}</w:t></w:r></w:p></w:endnote>
</w:endnotes>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  <w:p><w:r><w:t>Body with a long endnote</w:t></w:r><w:r>
    <w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/>
  </w:r></w:p>
  ${PAGE}
</w:body></w:document>`),
    },
  ]);
}
