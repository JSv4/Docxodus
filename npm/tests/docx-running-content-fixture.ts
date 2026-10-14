import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

/**
 * A generated two-section document with first/even/default running stories, for pinning where
 * headers, footers, and body text sit on the page.
 *
 * Word's page setup is FOUR independent distances from the paper edge, not two nested boxes:
 * `w:header` to the top of the header story, `w:footer` to the bottom of the footer story, and
 * `w:top`/`w:bottom` to the body text. Section 2 declares its own distances and margins while
 * declaring NO header/footer references, so it also covers the transition where a later section
 * inherits stories from an earlier one but must not inherit its geometry.
 */

/** Page setup for one section, in twips (1/20 pt) — the units `w:pgMar` itself uses. */
export interface SectionGeometryTwips {
  marginTop: number;
  marginBottom: number;
  headerDistance: number;
  footerDistance: number;
}

export interface RunningContentGeometry {
  first: SectionGeometryTwips;
  second: SectionGeometryTwips;
}

/** Twips → points, the unit the rendered page boxes are sized in. */
export const twipsToPt = (twips: number): number => twips / 20;

/** Section 1 uses Word's own defaults; section 2 deliberately differs in all four distances. */
export const DEFAULT_RUNNING_CONTENT_GEOMETRY: RunningContentGeometry = {
  first: { marginTop: 1440, marginBottom: 1440, headerDistance: 720, footerDistance: 720 },
  second: { marginTop: 1800, marginBottom: 1080, headerDistance: 1080, footerDistance: 360 },
};

/** Distinct, single-line story text so a spec can tell which variant a page selected. */
export const STORY_TEXT = {
  headerFirst: 'HEADER FIRST',
  headerDefault: 'HEADER DEFAULT',
  headerEven: 'HEADER EVEN',
  footerFirst: 'FOOTER FIRST',
  footerDefault: 'FOOTER DEFAULT',
  footerEven: 'FOOTER EVEN',
} as const;

const PAGE_WIDTH_TWIPS = 12240; // 8.5 in
const PAGE_HEIGHT_TWIPS = 15840; // 11 in
const SIDE_MARGIN_TWIPS = 1440;

/** Page width and height in points, for a spec that asserts against the paper edge. */
export const PAGE_WIDTH_PT = twipsToPt(PAGE_WIDTH_TWIPS);
export const PAGE_HEIGHT_PT = twipsToPt(PAGE_HEIGHT_TWIPS);

function story(text: string): string {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

/** `lines` paragraphs, the first carrying `text` so a spec can still identify the variant. */
function multiLineStory(text: string, lines: number): string {
  return Array.from({ length: Math.max(1, lines) }, (_, i) =>
    story(i === 0 ? text : `${text} cont ${i}`)).join('');
}

function headerPart(text: string, lines: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W_NS}">${multiLineStory(text, lines)}</w:hdr>`;
}

function footerPart(text: string, lines: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W_NS}">${multiLineStory(text, lines)}</w:ftr>`;
}

function bodyParagraphs(label: string, count: number): string {
  return Array.from({ length: count }, (_, i) => story(`${label} body line ${i + 1}`)).join('');
}

function sectPr(geometry: SectionGeometryTwips, references: string): string {
  return `${references}<w:pgSz w:w="${PAGE_WIDTH_TWIPS}" w:h="${PAGE_HEIGHT_TWIPS}"/>` +
    `<w:pgMar w:top="${geometry.marginTop}" w:right="${SIDE_MARGIN_TWIPS}"` +
    ` w:bottom="${geometry.marginBottom}" w:left="${SIDE_MARGIN_TWIPS}"` +
    ` w:header="${geometry.headerDistance}" w:footer="${geometry.footerDistance}" w:gutter="0"/>` +
    `<w:cols w:space="720"/><w:titlePg/>`;
}

/**
 * @param geometry per-section page setup; the default gives the two sections different
 *   header/footer distances AND margins, so a renderer that reuses section 1's numbers fails.
 * @param linesPerSection body lines per section — enough for at least THREE pages each, so the
 *   first, even, and odd/default story variants all get a page to be checked on.
 * @param storyLines paragraphs per running story. More than one, against a small margin, is how
 *   a story is made to reach past its margin and push the body — Word's overflow case.
 */
export function generateRunningContentDocx(
  geometry: RunningContentGeometry = DEFAULT_RUNNING_CONTENT_GEOMETRY,
  linesPerSection = 180,
  storyLines = 1,
): Uint8Array {
  const references =
    '<w:headerReference w:type="first" r:id="rId10"/>' +
    '<w:headerReference w:type="default" r:id="rId11"/>' +
    '<w:headerReference w:type="even" r:id="rId12"/>' +
    '<w:footerReference w:type="first" r:id="rId13"/>' +
    '<w:footerReference w:type="default" r:id="rId14"/>' +
    '<w:footerReference w:type="even" r:id="rId15"/>';

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
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/settings" Target="settings.xml"/>
  <Relationship Id="rId10" Type="${R_NS}/header" Target="header1.xml"/>
  <Relationship Id="rId11" Type="${R_NS}/header" Target="header2.xml"/>
  <Relationship Id="rId12" Type="${R_NS}/header" Target="header3.xml"/>
  <Relationship Id="rId13" Type="${R_NS}/footer" Target="footer1.xml"/>
  <Relationship Id="rId14" Type="${R_NS}/footer" Target="footer2.xml"/>
  <Relationship Id="rId15" Type="${R_NS}/footer" Target="footer3.xml"/>
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: 'word/settings.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}"><w:evenAndOddHeaders/></w:settings>`),
    },
    { name: 'word/header1.xml', data: xml(headerPart(STORY_TEXT.headerFirst, storyLines)) },
    { name: 'word/header2.xml', data: xml(headerPart(STORY_TEXT.headerDefault, storyLines)) },
    { name: 'word/header3.xml', data: xml(headerPart(STORY_TEXT.headerEven, storyLines)) },
    { name: 'word/footer1.xml', data: xml(footerPart(STORY_TEXT.footerFirst, storyLines)) },
    { name: 'word/footer2.xml', data: xml(footerPart(STORY_TEXT.footerDefault, storyLines)) },
    { name: 'word/footer3.xml', data: xml(footerPart(STORY_TEXT.footerEven, storyLines)) },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${bodyParagraphs('S1', linesPerSection)}
  <w:p><w:pPr><w:sectPr>${sectPr(geometry.first, references)}</w:sectPr></w:pPr></w:p>
  ${bodyParagraphs('S2', linesPerSection)}
  <w:sectPr>${sectPr(geometry.second, '')}</w:sectPr>
</w:body></w:document>`),
    },
  ]);
}
