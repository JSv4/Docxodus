import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export interface LegacyPageNumberOptions {
  kind?: 'legacy' | 'field' | 'empty' | 'tab';
  start?: number;
  roman?: boolean;
  bodyNumber?: boolean;
  pages?: number;
}

/** Original archive pages and independently authored running story. */
export function generateLegacyPageNumberDocx(options: LegacyPageNumberOptions = {}): Uint8Array {
  const runProperties = '<w:rPr><w:b/><w:i/><w:sz w:val="24"/></w:rPr>';
  const number = options.kind === 'field'
    ? `<w:r><w:fldChar w:fldCharType="begin"/></w:r>
       <w:r><w:instrText xml:space="preserve"> PAGE \\* Arabic </w:instrText></w:r>
       <w:r><w:fldChar w:fldCharType="separate"/></w:r>
       <w:r>${runProperties}<w:t>1</w:t></w:r>
       <w:r><w:fldChar w:fldCharType="end"/></w:r>`
    : `<w:r>${runProperties}<w:pgNum/></w:r>`;
  const header = options.kind === 'empty' ? '' : options.kind === 'tab' ? '<w:r><w:tab/></w:r>' : number;
  const body = Array.from({ length: options.pages ?? 3 }, (_, index) => `<w:p><w:pPr>
    ${index ? '<w:pageBreakBefore/>' : ''}<w:widowControl w:val="0"/><w:spacing w:after="0"/></w:pPr>
    <w:r><w:t xml:space="preserve">Body${index + 1} Cedar archive page.${options.bodyNumber ? ' Current page: ' : ''}</w:t></w:r>
    ${options.bodyNumber ? number : ''}</w:p>`).join('');
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="header" Type="${R_NS}/header" Target="header1.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/header1.xml': `<w:hdr xmlns:w="${W_NS}"><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${header}</w:p></w:hdr>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}
      <w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      <w:pgNumType w:fmt="${options.roman ? 'upperRoman' : 'decimal'}" w:start="${options.start ?? 1}"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
