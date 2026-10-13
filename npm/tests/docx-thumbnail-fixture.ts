import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Maple ledger and literal 2x2 RGB tile; no source document or image. */
export function generateThumbnailDocx(state: 'none' | 'present' | 'missing',
  owner: 'package' | 'document' = 'package', absolute = false): Uint8Array {
  const target = absolute ? '/docProps/preview.png'
    : owner === 'package' ? 'docProps/preview.png' : '../docProps/preview.png';
  const thumbnail = state === 'none' ? '' : `<Relationship Id="preview" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="${target}"/>`;
  const parts: Record<string, string | Buffer> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>${owner === 'package' ? thumbnail : ''}</Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>${owner === 'document' ? thumbnail : ''}</Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>MapleTop The original records remain complete.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  if (state === 'present') parts['docProps/preview.png'] = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAE0lEQVR4nGNwSes4c/cdAxADWQA0QAeTbjIQcQAAAABJRU5ErkJggg==', 'base64');
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: typeof value === 'string' ? xml(value) : value })));
}
