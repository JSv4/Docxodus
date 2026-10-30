import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** Original paired image labels, with either a promoted drawing or inline siblings. */
export function imageIdentityDocx(promoted: boolean, broken = true): Uint8Array {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const image = (id: number) => {
    const tag = promoted ? 'anchor' : 'inline';
    return `<w:drawing><wp:${tag} ${promoted ? 'simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"' : ''}>
      ${promoted ? '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>508000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>508000</wp:posOffset></wp:positionV>' : ''}
      <wp:extent cx="254000" cy="254000"/>${promoted ? '<wp:wrapNone/>' : ''}<wp:docPr id="${id}" name="Juniper tile"/>
      <a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Juniper tile"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="tile${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic></a:graphicData></a:graphic></wp:${tag}></w:drawing>`;
  };
  const relations = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  // A two-pixel black tile, generated entirely for the test fixtures.
  const healthy = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAC0lEQVR4nGNgQAYAAA4AAamRc7EAAAAASUVORK5CYII=', 'base64');
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`) },
    { name: '_rels/.rels', data: xml(relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`)) },
    { name: 'word/_rels/document.xml.rels', data: xml(relations([1, 2].map(id => `<Relationship Id="tile${id}" Type="${R_NS}/image" Target="media/tile${id}.png"/>`).join(''))) },
    { name: 'word/media/tile1.png', data: broken ? Buffer.from('Juniper deliberately undecodable tile') : healthy },
    { name: 'word/media/tile2.png', data: healthy },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:pic="${pic}"><w:body>
      <w:p><w:r>${image(1)}${promoted ? '' : image(2)}</w:r><w:r><w:t>Juniper sample register.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="8640" w:h="11520"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>
      </w:body></w:document>`) },
  ]);
}
