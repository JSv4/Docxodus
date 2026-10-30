import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** Original page decoration next to two short inventory paragraphs. */
export function positionedVmlShapeDocx(kind: 'oval' | 'rect', positioned = true): Uint8Array {
  const placement = positioned ? 'position:absolute;margin-left:18pt;margin-top:8pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;' : '';
  const shape = `<w:pict><v:${kind} xmlns:v="urn:schemas-microsoft-com:vml" style="${placement}width:140pt;height:${positioned ? 550 : 72}pt;z-index:-1"><v:fill color="#eaf2e2"/></v:${kind}></w:pict>`;
  const relations = (value: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${value}</Relationships>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`) },
    { name: '_rels/.rels', data: xml(relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`)) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>
      <w:p><w:r>${shape}</w:r><w:r><w:t>Maple inventory begins.</w:t></w:r></w:p>
      <w:p><w:r><w:t>Maple inventory complete.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="8640" w:h="11520"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>
      </w:body></w:document>`) },
  ]);
}
