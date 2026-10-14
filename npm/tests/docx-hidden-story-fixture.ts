import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type HiddenStory = 'body' | 'header' | 'footer';
export type HiddenMode = 'hidden' | 'empty' | 'visible' | 'mixed' | 'inherited';

/** Original text and literal OOXML; no external document or asset is read. */
export function generateHiddenStoryDocx(story: HiddenStory, mode: HiddenMode): Uint8Array {
  const runProperties = mode === 'inherited' ? '<w:rStyle w:val="Concealed"/>'
    : `<w:vanish w:val="${mode === 'visible' ? '0' : '1'}"/>`;
  const paragraph = mode === 'empty' ? '<w:p/>'
    : `<w:p><w:r><w:rPr>${runProperties}</w:rPr><w:t>Concealed archive note.</w:t></w:r>`
      + (mode === 'mixed' ? '<w:r><w:t>Visible story label.</w:t></w:r>' : '') + '</w:p>';
  const part = story === 'body' ? '' : `<Override PartName="/word/${story}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${story}+xml"/>`;
  const storyReference = story === 'body' ? '' : `<w:${story}Reference w:type="default" r:id="story"/>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${part}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>`
      + (story === 'body' ? '' : `<Relationship Id="story" Type="${R_NS}/${story}" Target="${story}1.xml"/>`) + '</Relationships>',
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>`
      + '<w:style w:type="character" w:styleId="Concealed"><w:name w:val="Concealed"/><w:rPr><w:vanish/></w:rPr></w:style></w:styles>',
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${story === 'body' ? paragraph : ''}
      <w:p><w:r><w:t>BODYMARK The archive tray is ready.</w:t></w:r></w:p>
      <w:sectPr>${storyReference}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
      </w:body></w:document>`,
  };
  if (story !== 'body') parts[`word/${story}1.xml`] = `<w:${story === 'header' ? 'hdr' : 'ftr'} xmlns:w="${W_NS}">${paragraph}</w:${story === 'header' ? 'hdr' : 'ftr'}>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
