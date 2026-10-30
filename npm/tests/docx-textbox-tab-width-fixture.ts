import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export function textboxTabWidthDocx(drawing: boolean, long: boolean): Uint8Array {
  const words = Array.from({ length: long ? 45 : 2 }, (_, index) => `Willow${String(index).padStart(2, '0')}`).join(' ');
  const paragraph = `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>
    <w:r><w:t>${words}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>RegisterClosed</w:t></w:r></w:p>`;
  const box = drawing
    ? `<w:drawing><wp:inline><wp:extent cx="2209800" cy="4572000"/><wp:docPr id="11" name="Willow register frame"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>
      <wps:txbx><w:txbxContent>${paragraph}</w:txbxContent></wps:txbx>
      <wps:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
    : `<w:pict><v:shape id="WillowRegister" style="width:174pt;height:360pt"><v:textbox inset="0,0,0,0">
      <w:txbxContent>${paragraph}</w:txbxContent></v:textbox></v:shape></w:pict>`;
  const relations = (value: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${value}</Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>`),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/>
      </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body><w:p><w:r>${box}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
