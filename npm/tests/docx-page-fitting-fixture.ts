import { deflateSync } from 'node:zlib';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/** A small original two-color tile, generated from RGB samples rather than a stored image. */
export function fittingTilePng(): Buffer {
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type), data]);
    let crc = 0xffffffff;
    for (const byte of body) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
    const length = Buffer.alloc(4), checksum = Buffer.alloc(4);
    length.writeUInt32BE(data.length); checksum.writeUInt32BE((crc ^ 0xffffffff) >>> 0);
    return Buffer.concat([length, body, checksum]);
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(4, 0); header.writeUInt32BE(2, 4); header[8] = 8; header[9] = 2;
  const pixels = Buffer.from([0, 22, 81, 97, 22, 81, 97, 241, 166, 54, 241, 166, 54,
    0, 241, 166, 54, 241, 166, 54, 22, 81, 97, 22, 81, 97]);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header),
    chunk('IDAT', deflateSync(pixels)), chunk('IEND', Buffer.alloc(0))]);
}

export function fittingImageRun(id: number, widthPt: number, heightPt: number): string {
  const cx = widthPt * 12700, cy = heightPt * 12700;
  return `<w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/>
    <wp:docPr id="${id}" name="Tile ${id}" descr="Two-color calibration tile ${id}"/>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
    <pic:nvPicPr><pic:cNvPr id="${id}" name="Tile ${id}"/><pic:cNvPicPr/></pic:nvPicPr>
    <pic:blipFill><a:blip r:embed="tile"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

export function pageFittingDocx(body: string): Uint8Array {
  const relations = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="tile" Type="${R_NS}/image" Target="media/tile.png"/>`),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip([...Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })),
    { name: 'word/media/tile.png', data: fittingTilePng() }]);
}

export function wideBodyTableDocx(): Uint8Array {
  const widths = [2300, 4300, 2900, 3500, 2600];
  const rows = Array.from({ length: 39 }, (_, row) => `<w:tr>${widths.map((width, column) => `<w:tc>
    <w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p>
    <w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/><w:spacing w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>
    <w:r><w:t>Ochre${String(row + 1).padStart(2, '0')}${String.fromCharCode(65 + column)} recorded.</w:t></w:r>
    ${row === 0 ? fittingImageRun(column + 1, 20, 10) : ''}</w:p></w:tc>`).join('')}</w:tr>`).join('');
  return pageFittingDocx(`<w:tbl><w:tblPr><w:tblW w:w="15600" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/></w:tblCellMar>
    </w:tblPr><w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>
    <w:p><w:r><w:t>Batch complete.</w:t></w:r></w:p>`);
}

export function inlineImageParagraphDocx(labels: boolean): Uint8Array {
  const sizes = [[410, 250], [360, 280], [425, 260], [390, 300]];
  const images = sizes.map(([width, height], index) => fittingImageRun(index + 1, width, height)
    + (labels ? `<w:r><w:t xml:space="preserve"> Maple${index + 1} recorded. </w:t></w:r>` : '')).join('');
  return pageFittingDocx(`<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
    <w:spacing w:after="0"/></w:pPr>${images}</w:p>
    <w:p><w:r><w:t>Image inventory complete.</w:t></w:r></w:p>`);
}

export function oversizedInlineImageDocx(height = 870, count = 1): Uint8Array {
  return pageFittingDocx(`<w:p><w:r><w:t>Before the tall tile.</w:t></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${Array.from({ length: count }, (_, index) => fittingImageRun(index + 1, 320, height)).join('')}</w:p>
    <w:p><w:r><w:t>After the tall tile.</w:t></w:r></w:p>`);
}

export function floatingPictureDocx(height = 300): Uint8Array {
  const picture = fittingImageRun(1, 240, height)
    .replace('<wp:inline>', `<wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>`)
    .replace('<wp:docPr', '<wp:wrapSquare wrapText="bothSides"/><wp:docPr')
    .replace('</wp:inline>', '</wp:anchor>');
  const lines = Array.from({ length: 26 }, (_, index) => `<w:r><w:t>Hazel${String(index + 1).padStart(2, '0')} registered.</w:t></w:r>`).join('<w:r><w:br/></w:r>');
  return pageFittingDocx(`<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/>
    <w:spacing w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>${lines}</w:p>
    <w:p><w:r><w:t>Hazel anchor.</w:t></w:r>${picture}</w:p>
    <w:p><w:r><w:t>Float inventory complete.</w:t></w:r></w:p>`);
}
