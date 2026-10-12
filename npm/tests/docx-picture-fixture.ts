import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

export function pictureDocx(tight: boolean, uniform = false): Uint8Array {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const wrap = tight ? `<wp:wrapTight wrapText="right" distR="114300"><wp:wrapPolygon edited="1">
    <wp:start x="0" y="0"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="0" y="0"/>
    </wp:wrapPolygon></wp:wrapTight>` : '<wp:wrapSquare wrapText="right" distR="114300"/>';
  const words = Array.from({ length: 120 }, (_, i) => `token${i}`).join(' ');
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      ${uniform ? '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' : ''}
      </Types>`) },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="office" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="picture" Type="${R_NS}/image" Target="media/black.png"/>
      ${uniform ? `<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>` : ''}</Relationships>`) },
    { name: 'word/media/black.png', data: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAC0lEQVR4nGNgQAYAAA4AAamRc7EAAAAASUVORK5CYII=', 'base64') },
    ...(uniform ? [{ name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"><w:docDefaults>
      <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
      </w:docDefaults></w:styles>`) }] : []),
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:pic="${pic}">
      <w:body><w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="68580" distR="68580" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="914400" cy="457200"/>${wrap}<wp:docPr id="1" name="Authored triangle"/>
      <a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="black.png"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r><w:r><w:t>${words}</w:t></w:r></w:p>
      ${uniform ? '<w:p><w:r><w:t>Following paragraph preserves its original position and source identity.</w:t></w:r></w:p>' : ''}
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body></w:document>`) },
  ]);
}
