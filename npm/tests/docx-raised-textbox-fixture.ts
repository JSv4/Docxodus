import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export function raisedTextboxDocx(position: number): Uint8Array {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const rels = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const types = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`;
  const body = `<w:document xmlns:w="${W_NS}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:wps="${wps}"><w:body><w:p><w:r><w:rPr><w:position w:val="${position}"/></w:rPr>
    <w:drawing><wp:inline><wp:extent cx="2286000" cy="457200"/><wp:docPr id="1" name="Cobalt label box"/>
    <a:graphic><a:graphicData uri="${wps}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="457200"/></a:xfrm></wps:spPr>
    <wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr><w:t>Cobalt label inside.</w:t></w:r></w:p></w:txbxContent></wps:txbx>
    <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
    <w:p><w:r><w:t>Cobalt inventory complete.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="8640" w:h="11520"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body></w:document>`;
  const styles = `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`;
  const parts = {'[Content_Types].xml': types, '_rels/.rels': rels(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': rels(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/>`),
    'word/document.xml': body, 'word/styles.xml': styles, 'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`};
  return storedZip(Object.entries(parts).map(([name, value]) => ({name, data: xml(value)})));
}
