import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

/**
 * Bytes that are declared `image/png` and are not a PNG.
 *
 * The export path base64-encodes an image part verbatim into a data URL — nothing between the
 * OPC package and the browser ever decodes it — so a part like this is the shortest honest way
 * to make `HTMLImageElement.decode()` reject inside the readiness barrier. A real but truncated
 * PNG would be decoder-dependent; this is undecodable everywhere.
 */
const UNDECODABLE_PNG = Buffer.from('not a png, and deliberately so', 'utf8');

/** A one-paragraph DOCX whose only inline image cannot be decoded by any browser. */
export function generateUndecodableImageDocx(): Uint8Array {
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
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
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/broken.png"/>
</Relationships>`),
    },
    { name: 'word/media/broken.png', data: UNDECODABLE_PNG },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:pic="${pic}">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Before the image. </w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Broken" descr="undecodable fixture image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="1" name="Broken" descr="undecodable fixture image"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rIdImage"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
      <w:r><w:t xml:space="preserve"> After the image.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`),
    },
  ]);
}
