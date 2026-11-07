import { storedZip } from "./mixed-section-fixture.mjs";

export const MISSING_FAMILY = "Absent Orchard Sans";
export const CONFIGURED_FAMILY = "Docxodus Canvas Mono";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const escape = text => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;");

export function fontObservationDocx({ family = CONFIGURED_FAMILY, text = "Orchard shelves hold seven tools.", unused } = {}) {
  const parts = {
    "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + ['document', 'styles', 'settings', 'fontTable'].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}${name === 'document' ? '.main' : ''}+xml"/>`).join('') + '</Types>',
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles', 'settings', 'fontTable'].map(name => `<Relationship Id="${name}" Type="${R}/${name}" Target="${name}.xml"/>`).join('') + '</Relationships>',
    "word/document.xml": `<w:document xmlns:w="${W}"><w:body>`
      + (Array.isArray(text) ? text : [text]).map(value => `<w:p><w:r><w:t xml:space="preserve">${escape(value)}</w:t></w:r></w:p>`).join('')
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
    "word/styles.xml": `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
      + (unused === 'style' ? `<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/><w:rPr><w:rFonts w:ascii="${MISSING_FAMILY}" w:hAnsi="${MISSING_FAMILY}"/></w:rPr></w:style>` : '') + '</w:styles>',
    "word/settings.xml": `<w:settings xmlns:w="${W}"/>`,
    "word/fontTable.xml": `<w:fonts xmlns:w="${W}"><w:font w:name="${family}"/>${unused === 'table' ? `<w:font w:name="${MISSING_FAMILY}"/>` : ''}</w:fonts>`,
  };
  return storedZip(Object.entries(parts).map(([name, text]) => ({ name, data: Buffer.from(text) })));
}
