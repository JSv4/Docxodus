import { deflateSync } from 'node:zlib';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** An original pale two-color tile, encoded from four RGB pixels. */
function backgroundPng(): Buffer {
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type), data]);
    let crc = 0xffffffff;
    for (const byte of body) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
    const size = Buffer.alloc(4), checksum = Buffer.alloc(4);
    size.writeUInt32BE(data.length); checksum.writeUInt32BE((crc ^ 0xffffffff) >>> 0);
    return Buffer.concat([size, body, checksum]);
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(2, 0); header.writeUInt32BE(2, 4); header[8] = 8; header[9] = 2;
  return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]), chunk('IHDR', header),
    chunk('IDAT', deflateSync(Buffer.from([0,238,247,243,246,240,226,0,246,240,226,238,247,243]))),
    chunk('IEND', Buffer.alloc(0))]);
}

export function backgroundPictureDocx(origin: 'page' | 'paragraph', small = false, relative = false): Uint8Array {
  const width = small || relative ? 18 : 432, height = small || relative ? 18 : 576;
  const cx = width * 12700, cy = height * 12700;
  const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const relations = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const namespace = `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="header" Type="${R_NS}/header" Target="header.xml"/>`),
    'word/_rels/header.xml.rels': relations(`<Relationship Id="tile" Type="${R_NS}/image" Target="media/background.png"/>`),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document ${namespace}><w:body>${paragraph('Copper archive: first delivery.')}
      <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Copper archive: second delivery.</w:t></w:r></w:p>
      ${paragraph('Six crates accounted for.')}
      <w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="8640" w:h="11520"/>
      <w:pgMar w:top="1080" w:right="720" w:bottom="1080" w:left="720" w:header="240" w:footer="240"/></w:sectPr>
      </w:body></w:document>`,
    'word/header.xml': `<w:hdr ${namespace}><w:p><w:r><w:drawing>
      <wp:anchor simplePos="0" relativeHeight="17" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="${origin === 'page' ? 'page' : 'column'}"><wp:posOffset>${origin === 'page' ? 0 : -36 * 12700}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="${origin}"><wp:posOffset>${origin === 'page' ? 0 : -12 * 12700}</wp:posOffset></wp:positionV>
      <wp:extent cx="${cx}" cy="${cy}"/><wp:wrapNone/><wp:docPr id="17" name="Copper background"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
      <pic:nvPicPr><pic:cNvPr id="17" name="Copper tile"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="tile"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic></a:graphicData></a:graphic>
      ${relative ? '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>100000</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom="page"><wp14:pctHeight>100000</wp14:pctHeight></wp14:sizeRelV>' : ''}
      </wp:anchor></w:drawing></w:r></w:p>${paragraph('Copper archive')}</w:hdr>`,
  };
  return storedZip([...Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })),
    { name: 'word/media/background.png', data: backgroundPng() }]);
}
