import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** A small observing log with independently authored, widely spaced running-story paragraphs. */
export function storySpacingDocx(story: 'header' | 'footer'): Uint8Array {
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
  const part = `${story}1.xml`;
  const paragraph = (text: string, spaced = false) => `<w:p><w:pPr>
    <w:spacing w:before="${spaced ? 360 : 0}" w:after="${spaced ? 360 : 0}" w:line="240" w:lineRule="exact"/>
    </w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>
    <w:t>${text}</w:t></w:r></w:p>`;
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
      ${paragraph('Night log caption.')}
      <w:sectPr><w:${story}Reference w:type="default" r:id="runningStory"/>
        <w:pgSz w:w="6000" w:h="7200"/>
        <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="240" w:footer="240"/>
      </w:sectPr></w:body></w:document>`,
    [`word/${part}`]: `<w:${root} xmlns:w="${W_NS}">
      ${Array.from({ length: 4 }, (_, i) => paragraph(`Observatory label ${i}`, true)).join('')}
    </w:${root}>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
