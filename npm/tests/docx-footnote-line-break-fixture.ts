import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export interface FootnoteLineBreakOptions {
  paragraphs?: boolean;
  formatted?: boolean;
  blankEvery?: number;
  lineCount?: number;
  hyperlink?: boolean;
}

/** Original Spruce archive note, authored from literal OOXML without a template. */
export function generateFootnoteLineBreakDocx(options: FootnoteLineBreakOptions = {}): Uint8Array {
  const props = '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>';
  const lines: string[] = [];
  for (let index = 0; index < (options.lineCount ?? 100); index++) {
    const text = `Note${String(index + 1).padStart(3, '0')} Spruce archive verified.`;
    const run = `<w:r><w:rPr>${options.formatted ? (index % 2 ? '<w:i/>' : '<w:b/>') : ''}</w:rPr><w:t>${text}</w:t></w:r>`;
    lines.push(options.hyperlink ? `<w:hyperlink r:id="link">${run}</w:hyperlink>` : run);
    if (options.blankEvery && (index + 1) % options.blankEvery === 0 && index + 1 < (options.lineCount ?? 100)) lines.push('');
  }
  const marker = '<w:r><w:footnoteRef/></w:r>';
  const note = options.paragraphs
    ? lines.map((line, index) => `<w:p>${props}${index === 0 ? marker : ''}${line}</w:p>`).join('')
    : `<w:p>${props}${marker}${lines.join('<w:r><w:br/></w:r>')}</w:p>`;
  const relations = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="notes" Type="${R_NS}/footnotes" Target="footnotes.xml"/>`),
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}" xmlns:r="${R_NS}">
      <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
      <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
      <w:footnote w:id="1">${note}</w:footnote></w:footnotes>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
      <w:p><w:r><w:t>BodyMarker Spruce archive has one long note.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
      </w:body></w:document>`,
  };
  if (options.hyperlink) parts['word/_rels/footnotes.xml.rels'] = relations(`<Relationship Id="link" Type="${R_NS}/hyperlink" Target="https://example.org/archive" TargetMode="External"/>`);
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
