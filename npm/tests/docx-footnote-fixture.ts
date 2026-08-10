import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * A generated document with footnotes, for pinning where the note area sits on the page.
 *
 * The note area is bottom-aligned to the body text band, and above the notes Word draws the
 * `w:separator` note's paragraph — so its geometry is a composition of the body bottom, the
 * separator's line box, and the notes' own spacing. Every part of that has to be independently
 * observable, which is why the fixture's body and note text are fixed and single-line.
 */

const PAGE_WIDTH_TWIPS = 12240; // 8.5 in
const PAGE_HEIGHT_TWIPS = 15840; // 11 in
const MARGIN_TWIPS = 1440; // 1 in

export const PAGE_HEIGHT_PT = PAGE_HEIGHT_TWIPS / 20;
export const MARGIN_PT = MARGIN_TWIPS / 20;
/** Word's built-in footnote separator length. */
export const SEPARATOR_WIDTH_IN = 2;

export interface FootnoteFixtureOptions {
  /** Body paragraphs, in order; each element is the number of notes that paragraph cites. */
  paragraphs: number[];
  /** Lines in each note. More than one is how a note is made long enough to split a page. */
  linesPerNote?: number;
  /** Bottom margin in twips, when a test needs the body band's bottom edge to move. */
  marginBottomTwips?: number;
  /**
   * Paragraphs in a default footer story, or 0 for no footer at all.
   *
   * Enough of them, against the default 1 inch bottom margin and 0.5 inch footer distance, and the
   * footer reaches ABOVE its margin and raises the body band's bottom edge — which is the only
   * shape that can tell "the notes are anchored to the body band" apart from "the notes are
   * anchored to the bottom margin".
   */
  footerLines?: number;
}

/** `w:pgMar/@w:footer` in twips: the distance from the paper's bottom edge to the footer story. */
export const FOOTER_DISTANCE_TWIPS = 720;
export const FOOTER_DISTANCE_PT = FOOTER_DISTANCE_TWIPS / 20;

function noteBody(index: number, lines: number): string {
  return Array.from({ length: lines }, (_, i) =>
    `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
    (i === 0
      ? `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>`
      : '') +
    `<w:r><w:t xml:space="preserve"> Note ${index} line ${i + 1}.</w:t></w:r></w:p>`).join('');
}

export function generateFootnoteDocx(options: FootnoteFixtureOptions): Uint8Array {
  const {
    paragraphs, linesPerNote = 1, marginBottomTwips = MARGIN_TWIPS, footerLines = 0,
  } = options;

  let nextNoteId = 1;
  const notes: string[] = [];
  const body = paragraphs
    .map((citations, p) => {
      const refs = Array.from({ length: citations }, () => {
        const id = nextNoteId++;
        notes.push(
          `<w:footnote w:id="${id}">${noteBody(id, linesPerNote)}</w:footnote>`);
        return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
          `<w:footnoteReference w:id="${id}"/></w:r>`;
      }).join('');
      return `<w:p>${refs}<w:r><w:t>Body paragraph ${p + 1}.</w:t></w:r></w:p>`;
    })
    .join('');

  // Word's two reserved notes. Their paragraphs are what the separator bands are drawn from.
  const reserved =
    `<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>` +
    `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>` +
    `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`;

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
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>${footerLines > 0 ? `
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` : ''}
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
  <Relationship Id="rId3" Type="${R_NS}/footnotes" Target="footnotes.xml"/>${footerLines > 0 ? `
  <Relationship Id="rId4" Type="${R_NS}/footer" Target="footer1.xml"/>` : ''}
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
  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
</w:styles>`),
    },
    {
      name: 'word/settings.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}"><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>`),
    },
    {
      name: 'word/footnotes.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W_NS}">${reserved}${notes.join('')}</w:footnotes>`),
    },
    ...(footerLines > 0 ? [{
      name: 'word/footer1.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W_NS}">${Array.from({ length: footerLines }, (_unused, i) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
  `<w:r><w:t>Footer line ${i + 1}</w:t></w:r></w:p>`).join('')}</w:ftr>`),
    }] : []),
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${body}
  <w:sectPr>${footerLines > 0 ? '<w:footerReference w:type="default" r:id="rId4"/>' : ''}
    <w:pgSz w:w="${PAGE_WIDTH_TWIPS}" w:h="${PAGE_HEIGHT_TWIPS}"/>
    <w:pgMar w:top="${MARGIN_TWIPS}" w:right="${MARGIN_TWIPS}" w:bottom="${marginBottomTwips}"
             w:left="${MARGIN_TWIPS}" w:header="720" w:footer="${FOOTER_DISTANCE_TWIPS}" w:gutter="0"/>
    <w:cols w:space="720"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}
