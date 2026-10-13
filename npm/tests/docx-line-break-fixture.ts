import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export interface LineBreakOptions {
  paragraphs?: boolean;
  formatted?: boolean;
  blankEvery?: number;
  lineCount?: number;
  keepLines?: boolean;
  widowControl?: boolean;
  hyperlink?: boolean;
}

/** Original register lines; no external DOCX, template, or generation helper is loaded. */
export function generateLineBreakDocx(options: LineBreakOptions = {}): Uint8Array {
  const properties = '<w:pPr>' + (options.keepLines ? '<w:keepLines/>' : '')
    + `<w:widowControl w:val="${options.widowControl ? 1 : 0}"/>`
    + '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/></w:pPr>';
  const lines = Array.from({ length: options.lineCount ?? 100 }, (_, index) => {
    const text = `Line${String(index + 1).padStart(3, '0')} Café sample.`;
    const formatting = options.formatted ? (index % 2 ? '<w:i/>' : '<w:b/>') : '';
    const run = `<w:r><w:rPr>${formatting}</w:rPr><w:t>${text}</w:t></w:r>`;
    return options.hyperlink ? `<w:hyperlink r:id="link">${run}</w:hyperlink>` : run;
  });
  const blocks: string[] = [];
  for (const [index, line] of lines.entries()) {
    blocks.push(line);
    if (options.blankEvery && (index + 1) % options.blankEvery === 0 && index + 1 < lines.length) blocks.push('');
  }
  const body = options.paragraphs
    ? blocks.map(line => `<w:p>${properties}${line}</w:p>`).join('')
    : `<w:p>${properties}${blocks.join('<w:r><w:br/></w:r>')}</w:p>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      ${options.hyperlink ? `<Relationship Id="link" Type="${R_NS}/hyperlink" Target="https://example.org/register" TargetMode="External"/>` : ''}</Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440"
      w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
