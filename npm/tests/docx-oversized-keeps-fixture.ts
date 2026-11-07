import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** A newly authored inventory with enough lines to exceed a page's available text band. */
export function oversizedKeepDocx(
  keep: 'keepLines' | 'widowControl' | 'keepNext' | 'styledKeepNext' | 'bothKeeps' | 'none',
  lines = 73, openingLines = 0,
): Uint8Array {
  const text = Array.from({ length: lines }, (_, index) => `<w:r><w:t>Copper${String(index + 1).padStart(3, '0')} transfer received.</w:t></w:r>`)
    .join('<w:r><w:br/></w:r>');
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Kept">
      <w:name w:val="Kept"/><w:pPr><w:keepNext/></w:pPr></w:style></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>
      ${openingLines ? `<w:p><w:pPr><w:spacing w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>
        ${Array.from({ length: openingLines }, (_, i) => `<w:r><w:t>Opening${i + 1}.</w:t></w:r>`).join('<w:r><w:br/></w:r>')}</w:p>` : ''}
      <w:p><w:pPr>${keep === 'styledKeepNext' ? '<w:pStyle w:val="Kept"/>'
        : keep === 'bothKeeps' ? '<w:keepNext/><w:keepLines/>' : keep === 'none' ? '' : `<w:${keep}/>`}
        <w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>${text}</w:p>
      <w:p><w:r><w:t>Final receipt signed.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440"
        w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
    </w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}
