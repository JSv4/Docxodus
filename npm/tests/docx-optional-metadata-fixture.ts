import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export type OptionalMetadataCase = 'clean' | 'broken-properties' | 'missing-thumbnail'
  | 'orphan-xml' | 'broken-styles' | 'missing-image' | 'properties-used-as-styles'
  | 'document-thumbnail' | 'thumbnail-used-as-image' | 'broken-relationships';

/** A newly authored stock receipt with a cached field and optional package metadata. */
export function optionalMetadataDocx(kind: OptionalMetadataCase): Uint8Array {
  const properties = kind === 'broken-properties' || kind === 'properties-used-as-styles';
  const thumbnail = kind === 'missing-thumbnail' || kind === 'document-thumbnail' || kind === 'thumbnail-used-as-image';
  const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  const rels = (value: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${value}</Relationships>`;
  const preview = rel('preview', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail', 'meta/preview.png');
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      ${properties ? '<Override PartName="/meta/editor.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' : ''}</Types>`,
    '_rels/.rels': rels(rel('document', `${R_NS}/officeDocument`, 'word/document.xml')
      + (properties ? rel('properties', `${R_NS}/custom-properties`, 'meta/editor.xml') : '')
      + (thumbnail && kind !== 'document-thumbnail' ? preview : '')),
    'word/_rels/document.xml.rels': rels(rel('styles', `${R_NS}/styles`,
      kind === 'properties-used-as-styles' ? '../meta/editor.xml' : 'styles.xml')
      + (kind === 'missing-image' ? rel('picture', `${R_NS}/image`, 'media/absent.png') : '')
      + (kind === 'thumbnail-used-as-image' ? rel('picture', `${R_NS}/image`, '../meta/preview.png') : '')
      + (kind === 'document-thumbnail' ? preview : '')),
    'word/styles.xml': kind === 'broken-styles' ? '<w:styles>'
      : `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
        <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
        </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r><w:t>Quartz receipt: twelve gauges arrived.</w:t></w:r></w:p>
      <w:p><w:fldSimple w:instr="DOCPROPERTY BatchTag"><w:r><w:t>QuartzCache 12.</w:t></w:r></w:fldSimple></w:p>
      <w:p><w:del w:id="41" w:author="Mira" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>Old rack.</w:delText></w:r></w:del>
      <w:ins w:id="42" w:author="Mira" w:date="2024-01-01T00:00:00Z"><w:r><w:t>New rack.</w:t></w:r></w:ins></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>
      </w:body></w:document>`,
  };
  if (properties) parts['meta/editor.xml'] = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property name="Unused"><unknown:value>Spare label</unknown:value></property></Properties>';
  if (kind === 'orphan-xml') parts['unused/scratch.xml'] = '<note><unknown:value>Unreferenced draft label</unknown:value></note>';
  if (kind === 'broken-relationships') parts['word/_rels/document.xml.rels'] = '<Relationships>';
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}
