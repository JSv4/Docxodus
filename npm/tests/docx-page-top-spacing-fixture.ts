import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

/**
 * Generated pagination probe for Word's paragraph space-before rule.
 *
 * The first section leaves less than one spaced paragraph on page 1, so NATURAL TOP is moved by
 * ordinary pagination.
 * PAGE-BREAK-BEFORE TOP starts page 3 by paragraph formatting. The second section then checks
 * the important exception: Word retains space-before on the first page of a new section.
 */

export const PAGE_TOP_SPACE_BEFORE_TWIPS = 360; // 18 pt
export const PAGE_TOP_LINE_TWIPS = 400; // 20 pt, exact
export const PAGE_TOP_MARGIN_TWIPS = 720; // 36 pt
export const PAGE_TOP_HEIGHT_TWIPS = 5760; // 4 in
export const PAGE_TOP_WIDTH_TWIPS = 5760;

export const PAGE_TOP_LABELS = {
  sectionStart: 'SECTION START',
  samePage: 'SAME-PAGE CONTROL',
  natural: 'NATURAL TOP',
  pageBreakBefore: 'PAGE-BREAK-BEFORE TOP',
  nextSection: 'NEXT-SECTION START',
} as const;

const paragraph = (
  text: string,
  options: { before?: number; pageBreakBefore?: boolean } = {},
): string => {
  const before = options.before ?? 0;
  const pageBreakBefore = options.pageBreakBefore ? '<w:pageBreakBefore/>' : '';
  return `<w:p><w:pPr><w:spacing w:before="${before}" w:after="0" ` +
    `w:line="${PAGE_TOP_LINE_TWIPS}" w:lineRule="exact"/>${pageBreakBefore}</w:pPr>` +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`;
};

const sectionProperties = (type = ''): string =>
  `${type ? `<w:type w:val="${type}"/>` : ''}` +
  `<w:pgSz w:w="${PAGE_TOP_WIDTH_TWIPS}" w:h="${PAGE_TOP_HEIGHT_TWIPS}"/>` +
  `<w:pgMar w:top="${PAGE_TOP_MARGIN_TWIPS}" w:right="${PAGE_TOP_MARGIN_TWIPS}" ` +
  `w:bottom="${PAGE_TOP_MARGIN_TWIPS}" w:left="${PAGE_TOP_MARGIN_TWIPS}" ` +
  `w:header="360" w:footer="360" w:gutter="0"/><w:cols w:space="720"/>`;

export function generatePageTopSpacingDocx(): Uint8Array {
  // Leave 20 pt of the 216 pt body: NATURAL TOP needs 18+20 pt and must move as a unit.
  const fillers = Array.from({ length: 6 }, (_, index) =>
    paragraph(`FILLER ${index + 1}`)).join('');

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${paragraph(PAGE_TOP_LABELS.sectionStart, { before: PAGE_TOP_SPACE_BEFORE_TWIPS })}
  ${paragraph(PAGE_TOP_LABELS.samePage, { before: PAGE_TOP_SPACE_BEFORE_TWIPS })}
  ${fillers}
  ${paragraph(PAGE_TOP_LABELS.natural, { before: PAGE_TOP_SPACE_BEFORE_TWIPS })}
  ${paragraph('PAGE 2 CONTROL', { before: PAGE_TOP_SPACE_BEFORE_TWIPS })}
  ${paragraph(PAGE_TOP_LABELS.pageBreakBefore, {
    before: PAGE_TOP_SPACE_BEFORE_TWIPS,
    pageBreakBefore: true,
  })}
  <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${PAGE_TOP_LINE_TWIPS}" w:lineRule="exact"/>
    <w:sectPr>${sectionProperties('nextPage')}</w:sectPr>
  </w:pPr></w:p>
  ${paragraph(PAGE_TOP_LABELS.nextSection, { before: PAGE_TOP_SPACE_BEFORE_TWIPS })}
  <w:sectPr>${sectionProperties()}</w:sectPr>
</w:body></w:document>`;

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="20"/><w:szCs w:val="20"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}
