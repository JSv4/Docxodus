import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export const tabEntries = [
  ['3.2.', 'Cedar sample inventory', '7'],
  ['3.12.', 'Quarterly cedar storage', '18'],
  ['3.13.', 'Cedar map register', '104'],
] as const;

// Original manual labels, two tab stops and hanging indentation; the first entry is a narrow-label control.
export function multipleTabFixture(): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${tabEntries.map(([label, title, number]) => `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1800"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs><w:ind w:left="1800" w:hanging="720"/></w:pPr><w:r><w:t>${label}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${title}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${number}</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}
