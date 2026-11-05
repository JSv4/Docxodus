import { storedZip, xml } from './docx-zip.js';

export const defaultFrameLines = Array.from({ length: 80 }, (_, index) =>
  `Tern record ${String(index + 1).padStart(2, '0')}: amber sail and cobalt stone.`);

/** Original flowing paragraphs whose style carries only inactive frame defaults. */
export function defaultFrameDocx(frame: 'none' | 'empty' | 'auto'): Uint8Array {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const frameProperties = frame === 'none' ? '' : `<w:framePr${frame === 'auto' ? ' w:wrap="auto"' : ''}/>`;
  const paragraphs = defaultFrameLines.map(text => '<w:p><w:pPr><w:pStyle w:val="Tern"/></w:pPr>' +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`).join('');
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>') },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="${rel}">` +
      '<Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>') },
    { name: 'word/_rels/document.xml.rels', data: xml(`<Relationships xmlns="${rel}">` +
      '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>') },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${w}"><w:docDefaults>` +
      '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="TernBase"><w:name w:val="Tern Base"/>' +
      `<w:pPr>${frameProperties}</w:pPr></w:style>` +
      '<w:style w:type="paragraph" w:styleId="Tern"><w:name w:val="Tern"/><w:basedOn w:val="TernBase"/></w:style></w:styles>') },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${w}"><w:body>` +
      '<w:p><w:r><w:t>Tern opening marker</w:t></w:r></w:p>' + paragraphs +
      '<w:p><w:r><w:t>Tern final marker</w:t></w:r></w:p><w:sectPr>' +
      '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440"/>' +
      '</w:sectPr></w:body></w:document>') },
  ]);
}
