import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export function vmlLineDocx(to: string | null, mixed: boolean, options: {
  from?: string | null; strokeWeight?: string | null; stroked?: boolean; extra?: string; host?: 'pict' | 'object';
} = {}): Uint8Array {
  const relations = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const { from = '0,0', strokeWeight = '2pt', stroked = true, extra = '', host = 'pict' } = options;
  const line = `<w:r><w:${host}><v:line${from === null ? '' : ` from="${from}"`}${to === null ? '' : ` to="${to}"`}` +
    ` strokecolor="#16777A"${strokeWeight === null ? '' : ` strokeweight="${strokeWeight}"`} stroked="${stroked ? 't' : 'f'}">${extra}</v:line></w:${host}></w:r>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`) },
    { name: '_rels/.rels', data: xml(relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`)) },
    { name: 'word/_rels/document.xml.rels', data: xml(relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>`)) },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p>${run('LineBefore')}</w:p><w:p>${mixed ? run('LineAnchor') : ''}${line}${mixed ? run('LineTail') : ''}</w:p><w:p>${run('LineAfter')}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`) },
  ]);
}
