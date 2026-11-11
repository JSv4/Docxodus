import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export const sourceClipLines = ['Quail red marker', 'Quail blue marker', 'Quail hidden marker'];
export type SourceClipKind = 'drawing' | 'vml';
export type SourceClipFit = 'fixed' | 'auto' | 'implicit' | 'overflow' | 'ellipsis';

/** Original three-line story: the source viewport holds two lines, unless it can grow. */
export function sourceTextboxClipDocx(kind: SourceClipKind, fit: SourceClipFit = 'fixed', offPaper = false,
  story: 'body' | 'header' | 'footer' = 'body', paragraphRelative = false,
  placement: { markup?: boolean; right?: boolean; bottom?: boolean; visible?: boolean; percent?: boolean } = {}): Uint8Array {
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const content = sourceClipLines.map((text, index) => '<w:p><w:pPr>' +
    `<w:spacing w:before="${index === 2 ? 160 : 0}" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr>` +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`).join('');
  const geometry = `<wp:extent cx="2032000" cy="${placement.percent ? 1778000 : 431800}"/>`;
  const x = placement.visible || placement.bottom ? 457200 : placement.right ? 8229600 : paragraphRelative ? -3556000 : -2794000;
  const y = placement.bottom ? 10287000 : 762000;
  const identity = '<wp:docPr id="1" name="Quail window"/>';
  const anchor = '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1">' +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${paragraphRelative ? 'column' : 'page'}"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${paragraphRelative ? 'paragraph' : 'page'}"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` + geometry + '<wp:wrapNone/>' + identity;
  const drawing = kind === 'vml'
    ? '<w:pict><v:shape id="quail-shape" style="width:160pt;height:34pt;' +
      (offPaper ? `position:absolute;z-index:-2;margin-left:${placement.right ? 648 : -220}pt;margin-top:60pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page` : '') +
      '"><v:textbox inset="0pt,0pt,0pt,0pt"' + (fit === 'auto' ? ' style="mso-fit-shape-to-text:true"' : '') + '>' +
      `<w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict>`
    : '<w:drawing>' + (offPaper ? anchor : '<wp:inline>' + geometry + identity) +
      `<a:graphic><a:graphicData uri="${wps}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
      `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
      '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"' +
      (fit === 'overflow' || fit === 'ellipsis' ? ` vertOverflow="${fit}"` : '') + '>' +
      (fit === 'fixed' || fit === 'overflow' || fit === 'ellipsis' ? '<a:noAutofit/>' : fit === 'auto' ? '<a:spAutoFit/>' : '') +
      '</wps:bodyPr></wps:wsp></a:graphicData></a:graphic>' + (offPaper ? (placement.percent ? '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>20000</wp14:pctWidth></wp14:sizeRelH>' : '') + '</wp:anchor>' : '</wp:inline>') + '</w:drawing>';
  const namespaces = `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wps="${wps}" ` +
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"';
  const boxParagraph = '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>' +
    (offPaper ? '<w:r><w:t>Quail placement host</w:t></w:r>' : '') + `<w:r>${drawing}</w:r></w:p>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      (story === 'body' ? '' : `<Override PartName="/word/${story}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${story}+xml"/>`) + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>` +
      (story === 'body' ? '' : `<Relationship Id="story" Type="${R_NS}/${story}" Target="${story}.xml"/>`) + '</Relationships>',
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/>' +
      '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>',
    'word/document.xml': `<w:document ${namespaces}><w:body>` + (story === 'body' ? boxParagraph : '') +
      '<w:p>' + (placement.markup ? '<w:pPr><w:jc w:val="center"/>' +
        '<w:pPrChange w:id="9" w:author="Quail Editor" w:date="2026-03-04T09:00:00Z"><w:pPr/></w:pPrChange></w:pPr>' : '') +
      '<w:r><w:t>Quail ledger complete</w:t></w:r></w:p>' +
      '<w:sectPr>' + (story === 'body' ? '' : `<w:${story}Reference w:type="default" r:id="story"/>`) +
      '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="720" w:bottom="1080" w:left="720" w:header="240" w:footer="240"/></w:sectPr></w:body></w:document>',
  };
  if (story !== 'body') {
    const tag = story === 'header' ? 'hdr' : 'ftr';
    parts[`word/${story}.xml`] = `<w:${tag} ${namespaces}>${boxParagraph}</w:${tag}>`;
  }
  return storedZip(Object.entries(parts).map(([name, data]) => ({ name, data: xml(data) })));
}
