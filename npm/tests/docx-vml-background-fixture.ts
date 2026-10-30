import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';
import { encodePng } from './visual-parity/png.js';

/** Original Oriole page stationery, including its generated two-color bitmap. */
export function generateVmlBackgroundDocx(kind: 'vml' | 'drawing' | 'inline'): Uint8Array {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const width = kind === 'inline' ? 96 : 648, height = kind === 'inline' ? 48 : 828;
  const vmlStyle = (kind === 'inline' ? '' :
    'position:absolute; z-index:-53; margin-left:0; margin-top:0; mso-position-horizontal:center; mso-position-horizontal-relative:margin; mso-position-vertical:center; mso-position-vertical-relative:margin;')
    + `width:${width}pt;height:${height}pt`;
  const vml = `<w:pict><v:shape id="OrioleStationery" style="${vmlStyle}"><v:imagedata r:id="stationery"/></v:shape></w:pict>`;
  const drawing = `<w:drawing><wp:anchor simplePos="0" relativeHeight="7" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"
      distT="0" distB="0" distL="0" distR="0"><wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
    <wp:positionV relativeFrom="margin"><wp:align>center</wp:align></wp:positionV>
    <wp:extent cx="${width * 12700}" cy="${height * 12700}"/><wp:wrapNone/><wp:docPr id="7" name="Oriole stationery"/>
    <a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Oriole stationery"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="stationery"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width * 12700}" cy="${height * 12700}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship
      Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="header" Type="${R_NS}/header" Target="header.xml"/></Relationships>`,
    'word/_rels/header.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="stationery" Type="${R_NS}/image" Target="media/oriole.png"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/header.xml': `<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:wp="${wp}" xmlns:a="${a}" xmlns:pic="${pic}"><w:p><w:r><w:t>Oriole background ledger.</w:t></w:r>
      <w:r>${kind === 'drawing' ? drawing : vml}</w:r></w:p></w:hdr>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
      <w:p><w:r><w:t>Oriole body record.</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/>
      <w:pgSz w:w="12960" w:h="16560"/><w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200"
      w:header="600" w:footer="600"/></w:sectPr></w:body></w:document>`,
  };
  const pixels = new Uint8Array(16 * 16 * 4);
  for (let index = 0; index < 16 * 16; index++) pixels.set(index < 32 ? [41, 117, 137, 255] : [242, 232, 212, 255], index * 4);
  return storedZip([...Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })),
    { name: 'word/media/oriole.png', data: encodePng({ width: 16, height: 16, data: pixels }) }]);
}
