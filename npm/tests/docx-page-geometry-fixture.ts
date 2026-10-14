import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original compact archive pages, authored directly in OOXML. Sizes are points. */
export function generatePageGeometryDocx(sizes: readonly (readonly [number, number])[]): Uint8Array {
  const section = ([width, height]: readonly [number, number]) => `<w:sectPr><w:type w:val="nextPage"/>
    <w:pgSz w:w="${width * 20}" w:h="${height * 20}"/><w:pgMar w:top="1440" w:right="1440"
    w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;
  const paragraphs = sizes.map((size, index) => `<w:p><w:pPr><w:widowControl w:val="0"/>
    <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>
    ${index < sizes.length - 1 ? section(size) : ''}</w:pPr><w:r><w:t>Geometry${index + 1} Willow archive page.</w:t></w:r></w:p>`).join('');
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${paragraphs}
      ${section(sizes[sizes.length - 1])}</w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
