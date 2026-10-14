import { storedZip, xml } from './docx-zip.js';

/**
 * A generated cached table of contents (issues #397/#427) — an actual complex TOC field around
 * dotted-leader entries and an ordinary hyperlink control after the field result. Both kinds of
 * link use the same character style, so field context is the only styling variable.
 *
 * The interesting property is that the entry text carries `w:rStyle w:val="Hyperlink"`, and the
 * `Hyperlink` character style declares a color and an underline. Word suppresses that presentation
 * for hyperlinks inside the cached TOC result, but applies it to the otherwise-identical ordinary
 * link. A renderer must therefore use field semantics, not paragraph-style or anchor heuristics.
 */

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The `Hyperlink` character style's declared appearance — Word's own default values. */
export const TOC_HYPERLINK_COLOR = '0563C1';
export const TOC_HYPERLINK_RGB = 'rgb(5, 99, 193)';

/** docDefaults automatic line spacing, in 240ths of a line. Word's default since Word 2013. */
export const TOC_LINE_TWIPS = 259;
/** `w:after` on the TOC1 style, in twips. */
export const TOC_AFTER_TWIPS = 100;
/** Body font size in half-points. */
export const TOC_SIZE_HALF_POINTS = 22;
/** Right tab stop for the page number, in twips. */
export const TOC_TAB_TWIPS = 9350;

export interface TocEntry {
  text: string;
  page: string;
}

export const TOC_ENTRIES: TocEntry[] = [
  { text: 'The first heading of the generated document', page: '1' },
  { text: 'The second heading of the generated document', page: '2' },
  { text: 'The third heading of the generated document', page: '3' },
];
export const ORDINARY_HYPERLINK_TEXT = 'An ordinary hyperlink using the same character style';

function entryXml(entry: TocEntry, index: number): string {
  const fieldStart = index === 0
    ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    : '';
  const fieldEnd = index === TOC_ENTRIES.length - 1
    ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    : '';
  // `w:webHidden` on the leader/page-number runs is what Word writes; they stay visible in print.
  return `<w:p><w:pPr><w:pStyle w:val="TOC1"/>` +
    `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${TOC_TAB_TWIPS}"/></w:tabs>` +
    `</w:pPr>${fieldStart}` +
    `<w:hyperlink w:anchor="_Toc${index}" w:history="1">` +
    `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr>` +
    `<w:t xml:space="preserve">${entry.text}</w:t></w:r>` +
    `<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc${index} \\h </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>${entry.page}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `</w:hyperlink>${fieldEnd}</w:p>`;
}

export function generateTocDocx(): Uint8Array {
  const body = TOC_ENTRIES.map(entryXml).join('\n    ');

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${w}" xmlns:r="${r}">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>
    ${body}
    <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
      <w:hyperlink w:anchor="_Ordinary" w:history="1">
        <w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr>
          <w:t>${ORDINARY_HYPERLINK_TEXT}</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
        w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${w}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="${TOC_SIZE_HALF_POINTS}"/><w:szCs w:val="${TOC_SIZE_HALF_POINTS}"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="${TOC_LINE_TWIPS}" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr><w:color w:val="${TOC_HYPERLINK_COLOR}"/><w:u w:val="single"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOCHeading">
    <w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="9"/></w:pPr>
    <w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TOC1">
    <w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="${TOC_AFTER_TWIPS}"/></w:pPr>
  </w:style>
</w:styles>`;

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`),
    },
    // Style resolution needs the settings part: without it the converter emits the character
    // style's CLASS but an empty rule, so `w:rStyle` silently loses its declared properties.
    {
      name: 'word/settings.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${w}"/>`),
    },
    { name: 'word/styles.xml', data: xml(stylesXml) },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}
