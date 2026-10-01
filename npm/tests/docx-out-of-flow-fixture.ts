import { storedZip, xml } from './docx-zip.js';

/**
 * Documents whose paragraphs host a floating DrawingML text box (issue #849). The paginated view
 * promotes such a box out of its paragraph into the page box, which leaves a paragraph whose only
 * content was the box with an empty, zero-height line in the text column.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

/** A run holding one anchored text box whose body is one paragraph per entry (empty → `<w:p/>`). */
export function textBoxRun(verticalFrom: string, boxParagraphs: string[]): string {
  const body = boxParagraphs
    .map((text) => (text ? `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` : '<w:p/>'))
    .join('');
  return `<w:r><w:drawing>
    <wp:anchor distT="0" distR="114300" distB="0" distL="114300" simplePos="0" relativeHeight="10"
      behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="${verticalFrom}"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="2286000" cy="914400"/>
      <wp:wrapSquare wrapText="bothSides"/>
      <wp:docPr id="1" name="Text Box 1"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic><a:graphicData uri="${WPS}"><wps:wsp>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="914400"/></a:xfrm></wps:spPr>
        <wps:txbx><w:txbxContent>${body}</w:txbxContent></wps:txbx>
        <wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
      </wps:wsp></a:graphicData></a:graphic>
    </wp:anchor>
  </w:drawing></w:r>`;
}

/** A one-section Letter-size document with `bodyXml` as its body blocks. */
export function outOfFlowDocx(bodyXml: string): Uint8Array {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}"><w:body>
  ${bodyXml}
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
