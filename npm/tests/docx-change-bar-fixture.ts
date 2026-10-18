import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original changed-paragraph pair: zero indent and a half-inch indented control.
export function changeBarFixture(): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${[0, 720].map((left, i) =>
      `<w:p><w:pPr><w:ind w:left="${left}"/></w:pPr><w:r><w:t>Cedar ledger </w:t></w:r><w:ins w:id="${i + 1}" w:author="Cedar Reviewer" w:date="2026-09-10T12:00:00Z"><w:r><w:t>updated.</w:t></w:r></w:ins></w:p>`).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"/>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

