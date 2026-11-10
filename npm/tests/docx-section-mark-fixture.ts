import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// The issue #1572 documents, rebuilt from their description: Liberation Serif 12 pt, Letter pages.
const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
  + 'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
export const SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
export const SPACED = '<w:spacing w:before="0" w:after="160" w:line="259" w:lineRule="auto"/>';
export const CONTINUOUS = `<w:sectPr><w:type w:val="continuous"/>${PAGE}</w:sectPr>`;
export const NEXT_PAGE = `<w:sectPr>${PAGE}</w:sectPr>`;
export const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

export const line = (text: string, props = SINGLE) =>
  `<w:p><w:pPr>${props}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
/** A paragraph carrying a section break (next-page by default), optionally with text of its own. */
export const mark = (props = SINGLE, text = '', section = NEXT_PAGE) =>
  `<w:p><w:pPr>${props}${section}</w:pPr>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
/** A next-page section break laid out in two balanced columns. */
export const TWO_COLUMNS = NEXT_PAGE.replace('</w:sectPr>', '<w:cols w:num="2" w:space="720"/></w:sectPr>');
export const table = (text: string) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
  + '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
  + `${line(text)}</w:tc></w:tr></w:tbl>`;

export function sectionMarkDocx(body: string, finalSection = CONTINUOUS, compat = 15): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + ['document', 'styles', 'settings'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles', 'settings'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('') + '</Relationships>',
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}${finalSection}</w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compat}"/></w:compat></w:settings>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}
