import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export const placementPins = ['MicaTopPin', 'MicaTailPin'];
export type PlacementCase = 'outside' | 'visible' | 'straddling' | 'internal-overflow' | 'unknown-basis';

/** Original explicit-overflow box: its paragraph anchor and margin alignment use separate bases. */
export function marginParagraphPlacementDocx(mode: PlacementCase): Uint8Array {
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const offsetPt = mode === 'outside' || mode === 'unknown-basis' ? 530 : mode === 'straddling' ? 424 : 60;
  const heightPt = mode === 'internal-overflow' ? 13 : 26;
  const content = placementPins.map((text, index) => '<w:p><w:pPr>' +
    `<w:spacing w:before="${mode === 'internal-overflow' && index === 1 ? 100 : 0}" w:after="0" w:line="240" w:lineRule="exact"/>` +
    `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`).join('');
  const drawing = '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="3" behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1">' +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${mode === 'unknown-basis' ? 'unrecognized' : 'margin'}"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetPt * 12700}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="812800" cy="${heightPt * 12700}"/><wp:wrapNone/><wp:docPr id="3" name="Mica pin window"/>` +
    `<a:graphic><a:graphicData uri="${wps}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
    `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
    '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" vertOverflow="overflow" horzOverflow="overflow"><a:noAutofit/></wps:bodyPr>' +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>';
  const parts = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="18"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:wps="${wps}" ` +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<w:body><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>' +
      `<w:r><w:t>Mica placement ledger</w:t></w:r><w:r>${drawing}</w:r></w:p>` +
      '<w:p><w:r><w:t>Mica ledger closed</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="7680" w:h="9600"/>' +
      '<w:pgMar w:top="960" w:right="960" w:bottom="960" w:left="960"/></w:sectPr></w:body></w:document>',
  };
  return storedZip(Object.entries(parts).map(([name, data]) => ({ name, data: xml(data) })));
}
