import { storedZip, xml } from './docx-zip.js';

export const textboxLineHeightLines = ['Heron copper lantern', 'Heron violet meadow', 'Heron amber window', 'Heron silver basket'];

/** Original fixed text story inside a host paragraph with an independent line-spacing rule. */
export function textboxLineHeightDocx(legacy: boolean, hostMultiple: number, exactInner = false): Uint8Array {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const height = 68;
  const content = textboxLineHeightLines.map(text => '<w:p><w:pPr><w:pStyle w:val="Heron"/></w:pPr>' +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`).join('');
  const drawing = legacy
    ? `<w:pict><v:shape style="width:190pt;height:${height}pt"><v:textbox inset="0,0,0,0">` +
      `<w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict>`
    : `<w:drawing><wp:inline><wp:extent cx="2413000" cy="${height * 12700}"/>` +
      '<wp:docPr id="1" name="Heron frame"/><a:graphic>' +
      `<a:graphicData uri="${wps}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
      `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
      '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr>' +
      '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>';
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>') },
    { name: '_rels/.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>') },
    { name: 'word/_rels/document.xml.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>') },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${w}">` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>' +
      '<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Heron"><w:name w:val="Heron"/><w:pPr>' +
      '<w:spacing w:before="0" w:after="0"' + (exactInner ? ' w:line="280" w:lineRule="exact"' : '') + '/>' + '</w:pPr></w:style></w:styles>') },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${w}" xmlns:wps="${wps}" ` +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">' +
      `<w:body><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${hostMultiple * 240}" w:lineRule="auto"/></w:pPr><w:r>${drawing}</w:r></w:p><w:p><w:r><w:t>Heron body marker</w:t></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
      '</w:sectPr></w:body></w:document>') },
  ]);
}
