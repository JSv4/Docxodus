import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

export const NOTE_TEXT = 'A spare quartz plate remains in the cabinet.';
const CLIPPED_TEXT = 'Quartz tail below the page.';

/**
 * One body line citing one single-line 10 pt footnote whose only run carries a formatting
 * revision (issue #1271). `formatRevision: false` is the control: the same note without the
 * `w:rPrChange`. `exactRowTwips` adds a one-row exact-height table to the note, whose text sits at
 * the foot of the row (under the same revision), for a note too tall for any page.
 */
export function noteFormatRevisionDocx({ formatRevision = true, exactRowTwips = 0,
  inheritedRevision = false, sourceMarker = false, separatorSize = 0 } = {}): Uint8Array {
  const change = formatRevision
    ? `<w:rPrChange w:id="31" w:author="Review Example" w:date="2026-09-12T10:00:00Z"><w:rPr>${inheritedRevision ? '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>' : '<w:i/>'}</w:rPr></w:rPrChange>`
    : '';
  // An exact-height row cannot split, so a row taller than the page cuts off the text at its foot.
  const table = exactRowTwips > 0
    ? `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${exactRowTwips}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:rPr>${change.replace('w:id="31"', 'w:id="32"')}</w:rPr><w:t>${CLIPPED_TEXT}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`
    : '';
  const font = '<w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="20"/>';
  const noteStyle = inheritedRevision ? `<w:style w:type="paragraph" w:styleId="SageNote"><w:name w:val="Sage Note"/>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/>${change}</w:rPr></w:style>` : '';
  const separator = separatorSize ? `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${separatorSize}"/></w:rPr>` : '';
  const marker = sourceMarker ? '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' : '';
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rNotes" Type="${R_NS}/footnotes" Target="footnotes.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>${noteStyle}</w:styles>`,
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}"><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>${separator}</w:pPr><w:r>${separator}<w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:pPr>${inheritedRevision ? '<w:pStyle w:val="SageNote"/>' : ''}<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${marker}<w:r><w:rPr>${inheritedRevision ? '' : font + change}</w:rPr><w:t xml:space="preserve">${NOTE_TEXT}</w:t></w:r></w:p>${table}</w:footnote></w:footnotes>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t xml:space="preserve">Receiver maintenance</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}
