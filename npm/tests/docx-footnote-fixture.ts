import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

/**
 * A generated one-page document with body text citing footnotes, for pinning where the paginated
 * footnote area sits on the page (issue #378).
 *
 * Word's model: the footnote area is anchored to the BOTTOM of the text column — the last note
 * line ends on the bottom margin line — with the separator rule drawn on the baseline of one
 * empty FootnoteText line above the first note. FootnoteText is 10pt single-spaced with zero
 * spacing-after, so the area contains no vertical spacing of its own; any air between the margin
 * line and the note ink is renderer-invented.
 */

export const FOOTNOTE_PAGE = {
  widthTwips: 12240, // 8.5in
  heightTwips: 15840, // 11in
  marginTwips: 1440, // 1in on all sides
} as const;

/** Twips → points, the unit the rendered page boxes are sized in. */
export const twipsToPt = (twips: number): number => twips / 20;

function bodyParagraph(index: number, noteId: number | null): string {
  const citation = noteId === null
    ? ''
    : `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
      `<w:footnoteReference w:id="${noteId}"/></w:r>`;
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">Body line ${index + 1}</w:t></w:r>${citation}</w:p>`;
}

function footnote(id: number, paragraphCount: number, wordsPerParagraph: number): string {
  const paragraphs = Array.from({ length: paragraphCount }, (_, index) =>
    `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
      (index === 0
        ? `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>`
        : '') +
      `<w:r><w:t xml:space="preserve"> ${wordsPerParagraph > 0
        ? Array.from({ length: wordsPerParagraph }, (_, wordIndex) =>
          `footnote-${id}-${index + 1}-${wordIndex + 1}`).join(' ')
        : `Footnote ${id} paragraph ${index + 1} text.`}</w:t></w:r></w:p>`)
    .join('');
  return `<w:footnote w:id="${id}">${paragraphs}</w:footnote>`;
}

/**
 * @param noteCount distinct footnotes, cited from the first `noteCount` body paragraphs.
 * @param bodyLines total body paragraphs — few enough to keep everything on one page, so the
 *   note area's position is determined purely by the page geometry, not by flow pressure.
 * @param paragraphsPerNote paragraphs emitted inside each note.
 * @param wordsPerParagraph when positive, emits this many uniquely numbered words in each note
 *   paragraph so one real converter paragraph can be made taller than a note band.
 */
export function generateFootnoteDocx(
  noteCount = 1,
  bodyLines = 5,
  paragraphsPerNote = 1,
  wordsPerParagraph = 0,
): Uint8Array {
  const noteIds = Array.from({ length: noteCount }, (_, i) => i + 1);
  const body = Array.from({ length: bodyLines }, (_, i) =>
    bodyParagraph(i, i < noteCount ? i + 1 : null)).join('');

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
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
  <Relationship Id="rId2" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
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
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
</w:styles>`),
    },
    {
      name: 'word/footnotes.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  ${noteIds.map((id) => footnote(id, paragraphsPerNote, wordsPerParagraph)).join('\n  ')}
</w:footnotes>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${body}
  <w:sectPr>
    <w:pgSz w:w="${FOOTNOTE_PAGE.widthTwips}" w:h="${FOOTNOTE_PAGE.heightTwips}"/>
    <w:pgMar w:top="${FOOTNOTE_PAGE.marginTwips}" w:right="${FOOTNOTE_PAGE.marginTwips}"
      w:bottom="${FOOTNOTE_PAGE.marginTwips}" w:left="${FOOTNOTE_PAGE.marginTwips}"
      w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}
