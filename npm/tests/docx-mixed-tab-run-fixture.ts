import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export const mixedTabRows = Array.from({ length: 6 }, (_, index) =>
  [`Kestrel shelf ${index + 1}`, `Copper lantern inventory ${index + 1}`, `Window alcove ${index + 1}`]);

/** Original three-column story; changing run boundaries must not change its layout. */
export function mixedTabRunDocx(story: 'body' | 'drawing' | 'vml', partition: 'split' | 'mixed' | 'single'): Uint8Array {
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const run = (content: string) => `<w:r><w:rPr><w:b/><w:color w:val="254C68"/></w:rPr>${content}</w:r>`;
  const text = (value: string) => `<w:t>${value}</w:t>`;
  const content = mixedTabRows.map(([prefix, middle, tail]) => {
    const runs = partition === 'single' ? run(text(prefix) + '<w:tab/>' + text(middle) + '<w:tab/>' + text(tail))
      : partition === 'mixed' ? run(text(prefix)) + run('<w:tab/>' + text(middle) + '<w:tab/>' + text(tail))
      : [text(prefix), '<w:tab/>', text(middle), '<w:tab/>', text(tail)].map(run).join('');
    return '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2800"/><w:tab w:val="left" w:pos="6000"/></w:tabs>' +
      '<w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr>' + runs + '</w:p>';
  }).join('');
  const shape = story === 'vml'
    ? '<w:pict><v:shape style="width:510pt;height:94pt"><v:textbox inset="0,0,0,0">' +
      `<w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict>`
    : '<w:drawing><wp:inline><wp:extent cx="6477000" cy="1193800"/><wp:docPr id="1" name="Kestrel ledger"/>' +
      `<a:graphic><a:graphicData uri="${wps}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
      `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
      '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr>' +
      '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>';
  const body = story === 'body' ? content : `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r>${shape}</w:r></w:p>`;
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:wps="${wps}" ` +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">' +
      `<w:body>${body}<w:p><w:r><w:t>Kestrel ledger complete</w:t></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body></w:document>',
  }).map(([name, data]) => ({ name, data: xml(data) })));
}
