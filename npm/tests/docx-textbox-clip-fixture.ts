import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** A fixed-height observation box with an independently authored trailing paragraph. */
export function textboxClipDocx(visibleTrailingText = false): Uint8Array {
  const paragraph = (text: string) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>
    <w:rPr><w:sz w:val="16"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>
      <w:p><w:r><w:pict><v:shape style="width:144pt;height:12pt" stroked="f" filled="f">
        <v:textbox inset="0,0,0,0"><w:txbxContent>${paragraph('Sighting logged.')}${paragraph(visibleTrailingText ? 'Off-box signal.' : '')}</w:txbxContent></v:textbox>
      </v:shape></w:pict></w:r></w:p>${paragraph('Orbit verified.')}
      <w:sectPr><w:pgSz w:w="6000" w:h="7200"/>
        <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
