import { storedZip, xml } from './docx-zip.js';

/**
 * The issue #1270 reproductions: a 2 × 1 inch teal rectangle (a DrawingML `wps:wsp` that is not a
 * text box), optionally inside a `wpg:wgp` group and optionally inside a tracked deletion, in its
 * own paragraph between two text paragraphs. The XML is the issue's generator's, verbatim.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

const SHAPE = '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>'
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="19647E"/></a:solidFill></wps:spPr>'
  + '<wps:bodyPr/></wps:wsp>';

export function shapeDocx({ grouped, deleted }: { grouped: boolean; deleted: boolean }): Uint8Array {
  const data = grouped
    ? '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>'
      + `<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>${SHAPE}</wpg:wgp>`
    : SHAPE;
  const drawing = '<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/>'
    + '<wp:docPr id="7" name="Teal inspection panel"/>'
    + `<a:graphic><a:graphicData uri="${grouped ? WPG : WPS}">${data}</a:graphicData></a:graphic>`
    + '</wp:inline></w:drawing></w:r>';
  const payload = deleted
    ? `<w:del w:id="31" w:author="Review Example" w:date="2026-09-12T10:00:00Z">${drawing}</w:del>`
    : drawing;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}"><w:body>
  <w:p><w:r><w:t xml:space="preserve">Optical panel</w:t></w:r></w:p>
  <w:p>${payload}</w:p>
  <w:p><w:r><w:t xml:space="preserve">End of inspection</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
  </w:sectPr>
</w:body></w:document>`;
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}

export const SHAPE_VARIANTS = [
  { name: 'shape-plain', grouped: false, deleted: false },
  { name: 'shape-deleted', grouped: false, deleted: true },
  { name: 'group-plain', grouped: true, deleted: false },
  { name: 'group-deleted', grouped: true, deleted: true },
] as const;
