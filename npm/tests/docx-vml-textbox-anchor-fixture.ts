import { storedZip, xml } from './docx-zip.js';

export const vmlPanelLabels = Array.from({ length: 8 }, (_, index) => `Plover panel ${index + 1}`);

/** Original page panels whose combined source heights exceed the ordinary body band. */
export function vmlTextboxAnchorDocx(basis: 'page' | 'margin' | 'inline', autoFit = false): Uint8Array {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const labels = basis === 'inline' ? vmlPanelLabels.slice(0, 1) : vmlPanelLabels;
  const boxes = labels.map((text, index) => {
    const placement = basis === 'inline' ? '' : `position:absolute;margin-left:${29 + (index % 2) * 254}pt;` +
      `margin-top:${58 + Math.floor(index / 2) * 137}pt;mso-position-horizontal-relative:${basis};` +
      `mso-position-vertical-relative:${basis};z-index:${index % 2 === 0 ? -23 : 23};`;
    return `<w:r><w:pict><v:shape id="Plover${index}" style="${placement}width:242pt;height:96pt" fillcolor="#d1e7e3" strokecolor="#3c6963">` +
      `<v:textbox inset="0,0,0,0"${autoFit ? ' style="mso-fit-shape-to-text:true"' : ''}><w:txbxContent>` +
      `<w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:p><w:r><w:t>Copper sail</w:t></w:r></w:p>` +
      '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>';
  }).join('');
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
      '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault>' +
      '</w:docDefaults></w:styles>') },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${w}" xmlns:v="urn:schemas-microsoft-com:vml">` +
      `<w:body><w:p>${boxes}</w:p><w:p><w:r><w:t>Plover body control</w:t></w:r></w:p><w:sectPr>` +
      '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440"/>' +
      '</w:sectPr></w:body></w:document>') },
  ]);
}
