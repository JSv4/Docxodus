import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original formatting revisions, with actual font/size changes and an unchanged control.
export function formatRevisionFixture(): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${[
      ['BoldExample', '<w:b/>', ''],
      ['SizeExample', '<w:sz w:val="28"/>', '<w:sz w:val="24"/>'],
      ['FontExample', '<w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>', '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>'],
    ].map(([text, current, previous], i) => `<w:p><w:r><w:rPr>${current}<w:rPrChange w:id="${i + 1}" w:author="Cedar Reviewer" w:date="2026-09-10T12:00:00Z"><w:rPr>${previous}</w:rPr></w:rPrChange></w:rPr><w:t>${text} Cedar revision sample.</w:t></w:r></w:p>`).join('')}<w:p><w:r><w:t>PlainExample Cedar control.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

