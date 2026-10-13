import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type PageMarginName = 'top' | 'right' | 'bottom' | 'left' | 'header' | 'footer';

/** Original Juniper register with two running stories, authored from literal XML. */
export function generatePageMarginUnitDocx(margins: Partial<Record<PageMarginName, string>> = {}, rows = 16,
  omitMargins = false): Uint8Array {
  const values = { top: '1080', right: '1080', bottom: '1080', left: '1080', header: '360', footer: '360', ...margins };
  const paragraph = (text: string) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const texts = ['PageTop Juniper margin records begin here.',
    'PageWide ' + 'The survey crew measures the sheet before recording each entry. '.repeat(5).trim(),
    ...Array.from({ length: rows }, (_, i) => `Juniper${String(i + 1).padStart(3, '0')} The page register retains this original line.`),
    'PageEnd Juniper margin records end here.'];
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'header1', 'footer1'].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name === 'document' ? 'document.main' : name.replace('1', '')}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="header" Type="${R_NS}/header" Target="header1.xml"/><Relationship Id="footer" Type="${R_NS}/footer" Target="footer1.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/header1.xml': `<w:hdr xmlns:w="${W_NS}">${paragraph('PageHead Juniper running header.')}</w:hdr>`,
    'word/footer1.xml': `<w:ftr xmlns:w="${W_NS}">${paragraph('PageFoot Juniper running footer.')}</w:ftr>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${texts.map(paragraph).join('')}<w:sectPr><w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="12240" w:h="15840"/>${omitMargins ? '' : `<w:pgMar ${Object.entries(values).map(([name, value]) => `w:${name}="${value}"`).join(' ')} w:gutter="0"/>`}</w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
