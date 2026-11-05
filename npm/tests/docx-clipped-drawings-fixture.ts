import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export interface ClippedDrawingOptions {
  story: 'body' | 'header';
  top: number;
  line?: boolean;
}

/** Original two-page Merlin ledger with a page-relative drawing. */
export function generateClippedDrawingDocx({ story, top, line = false }: ClippedDrawingOptions): Uint8Array {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const namespaces = `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:wps="${wps}"`;
  const width = 1524000, height = line ? 0 : 304800;
  const drawing = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0"
      simplePos="0" relativeHeight="7" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>${top * 12700}</wp:posOffset></wp:positionV>
    <wp:extent cx="${width}" cy="${height}"/><wp:wrapNone/><wp:docPr id="7" name="Merlin purple marker"/>
    <a:graphic><a:graphicData uri="${wps}"><wps:wsp><wps:cNvSpPr/><wps:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
      <a:prstGeom prst="${line ? 'line' : 'rect'}"><a:avLst/></a:prstGeom>
      ${line ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="53407A"/></a:solidFill>'}
      <a:ln w="19050"><a:solidFill><a:srgbClr val="53407A"/></a:solidFill></a:ln>
    </wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></w:r>`;
  const paragraph = (text: string, extra = '', properties = '') =>
    `<w:p>${properties}<w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship
      Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="header" Type="${R_NS}/header" Target="header.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/header.xml': `<w:hdr ${namespaces}>${paragraph('Merlin survey ledger.', story === 'header' ? drawing : '')}</w:hdr>`,
    'word/document.xml': `<w:document ${namespaces}><w:body>
      ${paragraph('Merlin ridge record.')}
      ${paragraph('Merlin river record.', story === 'body' ? drawing : '', '<w:pPr><w:pageBreakBefore/></w:pPr>')}
      <w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="12400" w:h="16000"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
      </w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
