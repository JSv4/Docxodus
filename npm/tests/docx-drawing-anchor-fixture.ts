import { storedZip, xml } from './docx-zip.js';

export type HorizontalAnchorOrigin =
  | 'page' | 'margin' | 'column' | 'character';
export type VerticalAnchorOrigin =
  | 'page' | 'margin' | 'paragraph' | 'line';

export interface DrawingAnchorFixtureOptions {
  horizontal: { relativeFrom: HorizontalAnchorOrigin; offsetEmu?: number; align?: 'left' | 'center' | 'right' };
  vertical: { relativeFrom: VerticalAnchorOrigin; offsetEmu?: number; align?: 'top' | 'center' | 'bottom' };
  prefix?: string;
  /**
   * Emit `<a:spAutoFit/>` in `wps:bodyPr`. Word then sizes the shape to its laid-out text plus
   * the body insets and treats the stored `wp:extent`/`a:ext` as a stale cache; without it the
   * stored extent is the height (issue #396).
   */
  autoFit?: boolean;
  /** Textbox body text; longer text wraps to more lines and, under auto-fit, a taller box. */
  text?: string;
  /** Stored extent. Auto-fit boxes deliberately carry a wrong one, as Word's own files do. */
  extentEmu?: { cx: number; cy: number };
  /**
   * `w:spacing w:line` in 240ths of a line with `w:lineRule="auto"` on the textbox paragraph.
   * Auto line spacing is a multiple of the FONT's line box, so this is what a content-driven
   * height must track.
   */
  lineTwips?: number;
}

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

function positionXml(
  axis: 'H' | 'V',
  value: DrawingAnchorFixtureOptions['horizontal'] | DrawingAnchorFixtureOptions['vertical'],
): string {
  const position = value.offsetEmu !== undefined
    ? `<wp:posOffset>${value.offsetEmu}</wp:posOffset>`
    : `<wp:align>${value.align}</wp:align>`;
  return `<wp:position${axis} relativeFrom="${value.relativeFrom}">${position}</wp:position${axis}>`;
}

/** A generated, image-free DrawingML textbox with deliberately asymmetric geometry. */
export function generateDrawingAnchorDocx(options: DrawingAnchorFixtureOptions): Uint8Array {
  const prefix = options.prefix
    ? `<w:r><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/>` +
      `<w:sz w:val="24"/></w:rPr><w:t>${options.prefix}</w:t></w:r>`
    : '';
  const extent = options.extentEmu ?? { cx: 1828800, cy: 914400 };
  const bodySpacing = options.lineTwips !== undefined
    ? `<w:pPr><w:spacing w:line="${options.lineTwips}" w:lineRule="auto"/></w:pPr>`
    : '';
  const bodyText = options.text ?? 'ANCHOR';
  const autoFit = options.autoFit ? '<a:spAutoFit/>' : '';

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${w}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:wps="${wps}">
  <w:body>
    <w:p>${prefix}<w:r><w:drawing>
      <wp:anchor distT="152400" distR="304800" distB="381000" distL="457200"
        simplePos="0" relativeHeight="10" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        ${positionXml('H', options.horizontal)}
        ${positionXml('V', options.vertical)}
        <wp:extent cx="${extent.cx}" cy="${extent.cy}"/>
        <wp:wrapSquare wrapText="bothSides"/>
        <wp:docPr id="1" name="Generated anchor"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="${wps}"><wps:wsp>
          <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${extent.cx}" cy="${extent.cy}"/></a:xfrm></wps:spPr>
          <wps:txbx><w:txbxContent><w:p>${bodySpacing}<w:r><w:t xml:space="preserve">${bodyText}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
          <wps:bodyPr lIns="114300" tIns="76200" rIns="228600" bIns="12700">${autoFit}</wps:bodyPr>
        </wps:wsp></a:graphicData></a:graphic>
      </wp:anchor>
    </w:drawing></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
        w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${w}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}
