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
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const SHAPE = '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>'
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="19647E"/></a:solidFill></wps:spPr>'
  + '<wps:bodyPr/></wps:wsp>';

export function shapeDocx({ grouped, deleted }: { grouped: boolean; deleted: boolean }): Uint8Array {
  const data = grouped
    ? '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>'
      + `<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>${SHAPE}</wpg:wgp>`
    : SHAPE;
  return inlineDrawingDocx(grouped ? WPG : WPS, data, deleted);
}

interface ExtraPart { id: string; type: string; contentType: string; name: string; body: string }

const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const diagramPart = (id: string, relationship: string, contentType: string, root: string): ExtraPart => ({
  id, type: `${REL}/${relationship}`, name: `diagrams/${relationship}1.xml`,
  contentType: `application/vnd.openxmlformats-officedocument.drawingml.${contentType}+xml`,
  body: `<dgm:${root} xmlns:dgm="${DGM}"/>`,
});

/**
 * Inline drawings or fills the converter cannot draw: unsupported shape geometry or paint, SmartArt,
 * a chart type it does not project (scatter), and a picture with no image (issue #1270 review). Each
 * keeps its extent, and the PDF export reports it as unsupported content.
 */
export const UNDRAWABLE_DRAWINGS: Record<string, { uri: string; data: string; parts?: ExtraPart[] }> = {
  triangle: { uri: WPS, data: SHAPE.replace('prst="rect"', 'prst="triangle"') },
  presetColorFill: { uri: WPS, data: SHAPE.replace('<a:srgbClr val="19647E"/>', '<a:prstClr val="red"/>') },
  pictureFill: { uri: WPS, data: SHAPE.replace('<a:solidFill><a:srgbClr val="19647E"/></a:solidFill>',
    '<a:blipFill><a:stretch><a:fillRect/></a:stretch></a:blipFill>') },
  smartArt: {
    uri: DGM,
    data: `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>`,
    parts: [diagramPart('rIdDm', 'diagramData', 'diagramData', 'dataModel'),
      diagramPart('rIdLo', 'diagramLayout', 'diagramLayout', 'layoutDef'),
      diagramPart('rIdQs', 'diagramQuickStyle', 'diagramStyle', 'styleDef'),
      diagramPart('rIdCs', 'diagramColors', 'diagramColors', 'colorsDef')],
  },
  chart: {
    uri: C,
    data: `<c:chart xmlns:c="${C}" r:id="rIdChart"/>`,
    parts: [{
      id: 'rIdChart', type: `${REL}/chart`, name: 'charts/chart1.xml',
      contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      body: `<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="lineMarker"/>`
        + '</c:scatterChart></c:plotArea></c:chart></c:chartSpace>',
    }],
  },
  pictureWithNoImage: {
    uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    data: '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="p"/>'
      + '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip/></pic:blipFill><pic:spPr/></pic:pic>',
  },
};

/** The issue's three paragraphs, the middle one holding one 2 × 1 inch inline drawing. */
export function inlineDrawingDocx(uri: string, data: string, deleted = false, parts: ExtraPart[] = []): Uint8Array {
  const drawing = '<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/>'
    + '<wp:docPr id="7" name="Teal inspection panel"/>'
    + `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic>`
    + '</wp:inline></w:drawing></w:r>';
  return drawingDocx(drawing, deleted, parts);
}

export function drawingDocx(drawing: string, deleted = false, parts: ExtraPart[] = [],
  after = 'End of inspection', before = ['Optical panel']): Uint8Array {
  const payload = deleted
    ? `<w:del w:id="31" w:author="Review Example" w:date="2026-09-12T10:00:00Z">${drawing}</w:del>`
    : drawing;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}"><w:body>
  ${before.map(text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('\n  ')}
  <w:p>${payload}</w:p>
  <w:p><w:r><w:t xml:space="preserve">${after}</w:t></w:r></w:p>
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
  ${parts.map((part) => `<Override PartName="/word/${part.name}" ContentType="${part.contentType}"/>`).join('')}
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
    ...(parts.length === 0 ? [] : [{
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${parts.map((part) =>
  `<Relationship Id="${part.id}" Type="${part.type}" Target="${part.name}"/>`).join('')}</Relationships>`),
    }]),
    ...parts.map((part) => ({ name: `word/${part.name}`, data: xml(part.body) })),
  ]);
}

export const SHAPE_VARIANTS = [
  { name: 'shape-plain', grouped: false, deleted: false },
  { name: 'shape-deleted', grouped: false, deleted: true },
  { name: 'group-plain', grouped: true, deleted: false },
  { name: 'group-deleted', grouped: true, deleted: true },
] as const;
