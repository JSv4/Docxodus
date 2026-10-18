import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * A Letter page of Times New Roman 12pt one-line body paragraphs, one of which cites a footnote
 * holding a one-column table of `rows` rows (10pt FootnoteText), authored as literal XML. Word
 * reference PDFs of two such shapes record how Word splits a note's table between pages (issue
 * #1178; docs/ooxml_corner_cases.md).
 */
export function generateNoteTableDocx(bodyLines: number, citingLine: number, rows: number, headerRow = false): Uint8Array {
  const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
  const reference = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
  const body = Array.from({ length: bodyLines }, (_, index) => {
    const label = `Body line ${String(index + 1).padStart(3, '0')}.`;
    return `<w:p>${text(label)}${index + 1 === citingLine ? reference : ''}</w:p>`;
  }).join('');
  const cell = (value: string, first: boolean) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>`
    + (first ? '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' : '')
    + `${text(value)}</w:p></w:tc>`;
  const table = '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>'
    + ['top', 'bottom', 'insideH'].map(side => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')
    + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
    + (headerRow ? `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('Note table heading.', true)}</w:tr>` : '')
    + Array.from({ length: rows }, (_, index) =>
      `<w:tr>${cell(`Note table row ${String(index + 1).padStart(3, '0')}.`, index === 0 && !headerRow)}</w:tr>`).join('')
    + '</w:tbl>';
  const separator = (type: string, id: number, mark: string) =>
    `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${mark}/></w:r></w:p></w:footnote>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${R_NS}/footnotes" Target="footnotes.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
      + '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
      + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>',
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}">${separator('separator', -1, 'separator')}${separator('continuationSeparator', 0, 'continuationSeparator')}<w:footnote w:id="1">${table}</w:footnote></w:footnotes>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
