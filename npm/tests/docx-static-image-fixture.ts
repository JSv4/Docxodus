import { R_NS as r, W_NS as w, storedZip, xml } from './docx-zip.js';
import { encodePng } from './visual-parity/png.js';

export type StaticImage = 'png' | 'tiff' | 'wmf' | 'anisotropic-wmf' | 'tiny-wmf' | 'hairline-wmf' | 'wide-wmf' | 'tall-wmf' | 'group-wmf' | 'unsupported-tiff' | 'unsupported-wmf';

/** Six original color pixels and a separate authored WMF rectangle/ellipse. */
function imageBytes(kind: StaticImage): Buffer {
  const rgb = Buffer.from([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0, 255, 0, 255, 0, 255, 255]);
  if (kind === 'png') return encodePng({ width: 2, height: 3,
    data: Uint8Array.from(Array.from({ length: 6 }, (_, i) => [...rgb.subarray(i * 3, i * 3 + 3), 255]).flat()) });
  if (kind.endsWith('tiff')) {
    const entries = [[256, 4, 1, 2], [257, 4, 1, 3], [258, 3, 3, 134], [259, 3, 1, kind === 'tiff' ? 1 : 5],
      [262, 3, 1, 2], [273, 4, 1, 140], [277, 3, 1, 3], [278, 4, 1, 3], [279, 4, 1, 18], [284, 3, 1, 1]];
    const data = Buffer.alloc(158); data.write('II'); data.writeUInt16LE(42, 2); data.writeUInt32LE(8, 4); data.writeUInt16LE(entries.length, 8);
    entries.forEach(([tag, type, count, value], i) => {
      const offset = 10 + i * 12; data.writeUInt16LE(tag, offset); data.writeUInt16LE(type, offset + 2);
      data.writeUInt32LE(count, offset + 4); data.writeUInt32LE(value, offset + 8);
    });
    for (let i = 0; i < 3; i++) data.writeUInt16LE(8, 134 + i * 2);
    rgb.copy(data, 140); return data;
  }
  const record = (fn: number, ...words: number[]) => {
    const data = Buffer.alloc(6 + words.length * 2); data.writeUInt32LE(3 + words.length, 0); data.writeUInt16LE(fn, 4);
    words.forEach((word, i) => data.writeInt16LE(word, 6 + i * 2)); return data;
  };
  const records = [record(0x0103, 8), record(0x020b, 0, 0), record(0x020c, 64, kind === 'anisotropic-wmf' ? 48 : 96),
    record(0x02fa, 0, kind === 'hairline-wmf' ? 0 : kind === 'anisotropic-wmf' ? 4 : 2, 0, 0x5030, 0x0080),
    record(0x012d, 0), record(0x02fc, 1, 0, 0, 0), record(0x012d, 1), record(0x041b, 55, 43, 8, 8), record(0x0418, 48, 88, 12, 52)];
  if (kind === 'unsupported-wmf') records.push(record(0x061d, 0, 0, 0, 0, 0, 0));
  records.push(record(0));
  const header = Buffer.alloc(40); header.writeUInt32LE(0x9ac6cdd7, 0); header.writeInt16LE(96, 10); header.writeInt16LE(64, 12); header.writeUInt16LE(96, 14);
  if (kind === 'tiny-wmf') { header.writeInt16LE(1, 10); header.writeInt16LE(1, 12); header.writeUInt16LE(65535, 14); }
  let checksum = 0; for (let i = 0; i < 20; i += 2) checksum ^= header.readUInt16LE(i); header.writeUInt16LE(checksum, 20);
  header.writeUInt16LE(1, 22); header.writeUInt16LE(9, 24); header.writeUInt16LE(0x300, 26);
  header.writeUInt32LE(9 + records.reduce((n, b) => n + b.length / 2, 0), 28); header.writeUInt16LE(2, 32);
  header.writeUInt32LE(Math.max(...records.map(b => b.length / 2)), 34);
  return Buffer.concat([header, ...records]);
}

export function staticImageDocx(kind: StaticImage, header: boolean): Uint8Array {
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main', pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const format = kind.split('-').at(-1)!, mime = format === 'wmf' ? 'image/x-wmf' : `image/${format}`;
  const width = kind === 'wide-wmf' ? 1828800 : 914400, height = kind === 'tall-wmf' ? 1219200 : 609600;
  const namespaces = `xmlns:w="${w}" xmlns:r="${r}" xmlns:a="${a}" xmlns:pic="${pic}" xmlns:wp="${wp}"`;
  const drawing = kind === 'group-wmf' ? `<w:p><w:r><w:pict><v:group xmlns:v="urn:schemas-microsoft-com:vml"
    style="width:144pt;height:48pt" coordsize="96,64"><v:shape id="CypressGroupImage" style="left:0;top:0;width:96;height:64">
    <v:imagedata r:id="picture"/></v:shape></v:group></w:pict></w:r></w:p>`
    : `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="${width}" cy="${height}"/><wp:docPr id="1" name="Cypress image"/>
    <a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Cypress image"/><pic:cNvPicPr/></pic:nvPicPr>
    <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const imageRelation = `<Relationship Id="picture" Type="${r}/image" Target="media/cypress.${format}"/>`;
  const entries = [
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="${format}" ContentType="${mime}"/>${[['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings'], ...(header ? [['header', 'header']] : [])].map(([name, type]) =>
        `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('')}</Types>`) },
    { name: '_rels/.rels', data: xml(rels(`<Relationship Id="main" Type="${r}/officeDocument" Target="word/document.xml"/>`)) },
    { name: 'word/_rels/document.xml.rels', data: xml(rels(['styles', 'settings', ...(header ? ['header'] : [])].map(name =>
      `<Relationship Id="${name}" Type="${r}/${name}" Target="${name}.xml"/>`).join('') + (header ? '' : imageRelation))) },
    { name: 'word/media/cypress.' + format, data: imageBytes(kind) },
    { name: 'word/document.xml', data: xml(`<w:document ${namespaces}><w:body>${header ? '' : drawing}<w:p><w:r><w:t>CypressAfter</w:t></w:r></w:p><w:sectPr>
      ${header ? '<w:headerReference w:type="default" r:id="header"/>' : ''}<w:pgSz w:w="6120" w:h="7920"/>
      <w:pgMar w:top="1800" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/></w:sectPr></w:body></w:document>`) },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
      <w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`) },
    { name: 'word/settings.xml', data: xml(`<w:settings xmlns:w="${w}"/>`) },
  ];
  if (header) entries.push({ name: 'word/header.xml', data: xml(`<w:hdr ${namespaces}>${drawing}</w:hdr>`) },
    { name: 'word/_rels/header.xml.rels', data: xml(rels(imageRelation)) });
  return storedZip(entries);
}
