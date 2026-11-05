import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** A running-story table authored to extend into one side margin while staying on the paper. */
export function storyMarginsDocx(story: 'header' | 'footer', side: 'left' | 'right'): Uint8Array {
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
  const part = `${story}1.xml`;
  const paragraph = (text: string) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const widths = side === 'left' ? [480, 4560] : [4560, 480];
  const root = story === 'header' ? 'hdr' : 'ftr';
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="${type}document.main+xml"/>
      <Override PartName="/word/${part}" ContentType="${type}${story}+xml"/>
    </Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="runningStory" Type="${R_NS}/${story}" Target="${part}"/>
    </Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
      ${paragraph('Orbital survey.')}
      <w:sectPr><w:${story}Reference w:type="default" r:id="runningStory"/>
        <w:pgSz w:w="6000" w:h="7200"/>
        <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="240" w:footer="240"/>
      </w:sectPr></w:body></w:document>`,
    [`word/${part}`]: `<w:${root} xmlns:w="${W_NS}"><w:tbl>
      <w:tblPr><w:tblW w:w="5040" w:type="dxa"/><w:tblLayout w:type="fixed"/>
        <w:tblInd w:w="${side === 'left' ? -480 : 0}" w:type="dxa"/>
        <w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
      </w:tblPr><w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
      <w:tr>${widths.map((width, i) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>
        ${paragraph(width === 480 ? 'Rim' : 'Survey table')}</w:tc>`).join('')}</w:tr>
    </w:tbl></w:${root}>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
