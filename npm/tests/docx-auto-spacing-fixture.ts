import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** Original multi-spaced inventory with alternating emphasis and an optional short footnote. */
export function autoSpacingParagraphDocx(note: boolean): Uint8Array {
  const lines = Array.from({ length: 83 }, (_, index) => `<w:r><w:rPr>${index % 2 ? '<w:i/>' : '<w:b/>'}</w:rPr>
    <w:t>Juniper${String(index + 1).padStart(3, '0')} checkpoint recorded.</w:t></w:r>
    ${note && index === 35 ? '<w:r><w:footnoteReference w:id="1"/></w:r>' : ''}`).join('<w:r><w:br/></w:r>');
  const relations = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="notes" Type="${R_NS}/footnotes" Target="footnotes.xml"/>`),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}">
      <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
      <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
      <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Juniper note checked.</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
      <w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/></w:pPr>${lines}</w:p>
      <w:p><w:r><w:t>Checkpoint inventory complete.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440"
      w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
