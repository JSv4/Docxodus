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

/** Literal six-part source used by the native Word kept-note/widow controls. */
export function generateNormalKeptNoteDocx(lineCount: number, widow?: boolean, compatibilityMode?: number): Uint8Array {
  const spacing = '<w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/>';
  const body = `<w:p><w:pPr>${spacing}</w:pPr><w:r><w:t>SpruceCite01</w:t></w:r>`
    + '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>'
    + `<w:p><w:pPr>${spacing}</w:pPr><w:r><w:t>SpruceTail01</w:t></w:r></w:p>`;
  const widowControl = widow === undefined ? '' : widow ? '<w:widowControl/>' : '<w:widowControl w:val="0"/>';
  const content = '<w:r><w:footnoteRef/></w:r>' + Array.from({ length: lineCount }, (_, i) =>
    `<w:r><w:t>SpruceNoteLine${String(i + 1).padStart(2, '0')}</w:t></w:r>`).join('<w:r><w:br/></w:r>');
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const overrides = [['document', 'document.main'], ['styles', 'styles'], ['footnotes', 'footnotes'],
    ...(compatibilityMode === undefined ? [] : [['settings', 'settings']])]
    .map(([part, kind]) => `<Override PartName="/word/${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`).join('');
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
    '_rels/.rels': `<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R_NS}/footnotes" Target="footnotes.xml"/>${compatibilityMode === undefined ? '' : `<Relationship Id="rIdFixtureSettings" Type="${R_NS}/settings" Target="settings.xml"/>`}</Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${spacing}</w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}"><w:footnote w:id="1"><w:p><w:pPr><w:keepLines/>${widowControl}${spacing}</w:pPr>${content}</w:p></w:footnote></w:footnotes>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="8000" w:h="6400"/><w:pgMar w:top="200" w:right="200" w:bottom="200" w:left="200" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  if (compatibilityMode !== undefined) parts['word/settings.xml'] = compatibilitySettings(compatibilityMode);
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

function compatibilitySettings(mode: number): string {
  return `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`;
}

function bodyParagraph(index: number, noteId: number | null): string {
  const citation = noteId === null
    ? ''
    : `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
      `<w:footnoteReference w:id="${noteId}"/></w:r>`;
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">Body line ${index + 1}</w:t></w:r>${citation}</w:p>`;
}

function footnote(
  id: number,
  paragraphCount: number,
  wordCounts: number | readonly number[],
  paragraphProperties: string,
): string {
  const paragraphWordCounts = Array.isArray(wordCounts) ? wordCounts : undefined;
  const wordsPerParagraph = typeof wordCounts === 'number' ? wordCounts : 0;
  const effectiveParagraphCount = paragraphWordCounts?.length ?? paragraphCount;
  const paragraphs = Array.from({ length: effectiveParagraphCount }, (_, index) => {
    const requestedWords = paragraphWordCounts?.[index];
    if (requestedWords !== undefined && (!Number.isInteger(requestedWords) || requestedWords < 1)) {
      throw new Error(`Footnote paragraph ${index + 1} needs a positive integer word count.`);
    }
    // Emit EXACTLY the requested number of words. Padding a five-word stem with `repeat(n - 5)`
    // clamped every count below five to the same text, so a case asking for uneven continuation
    // pressure silently got none and the test it fed proved nothing about the uneven path.
    const text = requestedWords === undefined
      ? wordsPerParagraph > 0
        ? Array.from({ length: wordsPerParagraph }, (_, wordIndex) =>
          `footnote-${id}-${index + 1}-${wordIndex + 1}`).join(' ')
        : `Footnote ${id} paragraph ${index + 1} text.`
      : Array.from({ length: requestedWords }, (_, wordIndex) =>
        `note${id}p${index + 1}w${wordIndex + 1}`).join(' ');
    return (
      `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/>${paragraphProperties}</w:pPr>` +
      (index === 0
        ? `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>`
        : '') +
      `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`
    );
  })
    .join('');
  return `<w:footnote w:id="${id}">${paragraphs}</w:footnote>`;
}

/**
 * @param noteCount distinct footnotes, cited from the first `noteCount` body paragraphs.
 * @param bodyLines total body paragraphs — few enough to keep everything on one page, so the
 *   note area's position is determined purely by the page geometry, not by flow pressure.
 * @param paragraphsPerNote paragraphs emitted inside each note.
 * @param wordCounts either a shared exact word count for every note paragraph, or per-paragraph
 *   word counts for constructing notes with uneven continuation pressure.
 * @param noteParagraphProperties extra `w:pPr` children for every note paragraph, e.g. a
 *   hanging `w:ind`.
 * @param firstCitationLine one-based body paragraph that cites the first note; later notes follow.
 */
export function generateFootnoteDocx(
  noteCount = 1,
  bodyLines = 5,
  paragraphsPerNote = 1,
  wordCounts: number | readonly number[] = 0,
  separatorStories: {
    normalText?: string;
    continuationText?: string;
  } = {},
  noteParagraphProperties = '',
  firstCitationLine = 1,
  options: { compatibilityMode?: number; fontFamily?: 'Liberation Serif' | 'Times New Roman' } = {},
): Uint8Array {
  const { compatibilityMode, fontFamily = 'Liberation Serif' } = options;
  const noteIds = Array.from({ length: noteCount }, (_, i) => i + 1);
  const body = Array.from({ length: bodyLines }, (_, i) =>
    bodyParagraph(i, i >= firstCitationLine - 1 && i < firstCitationLine - 1 + noteCount
      ? i - firstCitationLine + 2 : null)).join('');

  const separatorRun = (marker: 'separator' | 'continuationSeparator', text?: string) =>
    `<w:p><w:r><w:${marker}/>${text ? `<w:t xml:space="preserve"> ${text}</w:t>` : ''}</w:r></w:p>`;

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
  ${compatibilityMode === undefined ? '' : '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'}
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
  ${compatibilityMode === undefined ? '' : `<Relationship Id="rIdFixtureSettings" Type="${R_NS}/settings" Target="settings.xml"/>`}
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/>
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
  <w:footnote w:type="separator" w:id="-1">${separatorRun(
    'separator', separatorStories.normalText,
  )}</w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">${separatorRun(
    'continuationSeparator', separatorStories.continuationText,
  )}</w:footnote>
  ${noteIds.map((id) => footnote(id, paragraphsPerNote, wordCounts, noteParagraphProperties)).join('\n  ')}
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
    ...(compatibilityMode === undefined ? [] : [{ name: 'word/settings.xml', data: xml(compatibilitySettings(compatibilityMode)) }]),
  ]);
}
