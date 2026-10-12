import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type PrintGeometryCase = 'zero-width' | 'ordinary' | 'drawing-only' | 'drawing-with-text';

/** Literal OOXML and original text; no external document or image is loaded. */
export function generatePrintGeometryDocx(kind: PrintGeometryCase): Uint8Array {
  const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const drawing = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0"
    simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
    <wp:extent cx="2743200" cy="914400"/><wp:wrapNone/><wp:docPr id="1" name="Original archive label"/><wp:cNvGraphicFramePr/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
    <wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>
    <wps:txbx><w:txbxContent>${paragraph('BoxMarker Amber sample tray.')}</w:txbxContent></wps:txbx>
    <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr></wps:wsp>
    </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const body = kind.startsWith('drawing')
    ? '<w:p>' + (kind === 'drawing-with-text' ? '<w:r><w:t>HostMarker The archive label.</w:t></w:r>' : '')
      + drawing + '</w:p>' + paragraph('FollowingMarker The tray is ready.')
    : `<w:p>${kind === 'zero-width' ? '<w:pPr><w:ind w:left="5000" w:right="5000"/></w:pPr>' : ''}<w:r><w:t>ABC</w:t></w:r></w:p>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
