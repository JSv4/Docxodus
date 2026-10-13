import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type CustomPropertiesCase = 'none' | 'text' | 'number' | 'boolean' | 'mismatched' | 'truncated' | 'entity';

/** Literal original Cedar ledger and metadata, without a source DOCX. */
export function generateCustomPropertiesDocx(kind: CustomPropertiesCase): Uint8Array {
  const present = kind !== 'none';
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${present ? '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' : ''}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>${present ? `<Relationship Id="custom" Type="${R_NS}/custom-properties" Target="docProps/custom.xml"/>` : ''}</Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>CedarTop All body records stay visible.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  if (present) {
    const value = kind === 'number' ? '<vt:i4>37</vt:i4>' : kind === 'boolean' ? '<vt:bool>true</vt:bool>'
      : kind === 'entity' ? '<vt:lpwstr>&missing;</vt:lpwstr>' : '<vt:lpwstr>CedarMetadataOnly &amp; &lt;kept&gt;.</vt:lpwstr>';
    let metadata = `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="CedarRecord">${value}</property></Properties>`;
    if (kind === 'mismatched') metadata = metadata.replace('</property>', '</wrong>');
    if (kind === 'truncated') metadata = metadata.replace('</Properties>', '');
    parts['docProps/custom.xml'] = metadata;
  }
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
