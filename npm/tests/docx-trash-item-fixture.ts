import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Aspen resource ledger, authored as literal XML without a source DOCX. */
export function generateTrashItemDocx(name: string | null, datDefault = false, corrupt = false): Uint8Array {
  const marker = 'DiscardedAspenChecksumMarker';
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${datDefault ? '<Default Extension="dat" ContentType="application/xml"/>' : ''}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>AspenTop The resource ledger retains every original entry.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  if (name) parts[name] = `<obsolete>${marker}`;
  if (datDefault) parts['resources/live.dat'] = '<live/>';
  const bytes = storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
  if (corrupt) {
    const position = Buffer.from(bytes).indexOf(marker);
    if (position < 0) throw new Error('Checksum fixture marker is missing');
    bytes[position] ^= 1;
  }
  return bytes;
}
