import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export const fitTexts = {
  justify: 'Every autumn the ferry crews repaint the harbor signs before the storms reach our northern pier as the gulls complain about it all afternoon long.',
  left: 'We Took A Yellow Wagon To Avery Valley, Past Your Tower, Toward Lake Vale, Where Water Waits.',
};
export type FitNeighbor = 'plain' | 'bold' | 'table' | 'list' | 'link';
/**
 * Where the fitted text sits: an ordinary paragraph, one with a bold word (the issue's ` harb`, the
 * first `the`, or the last word), one with a hyperlinked word, a numbered list item, a table cell
 * with Word's default cell margins, or a paragraph that states no line spacing.
 */
export type FitTarget = 'paragraph' | 'bold' | 'bold-the' | 'bold-last' | 'link' | 'list' | 'cell' | 'unspaced';
/** An ordinary or last-word-bold paragraph's text follows its alignment; other targets hold the justified text. */
export const fitText = (alignment: keyof typeof fitTexts, target: FitTarget = 'paragraph') =>
  fitTexts[target === 'paragraph' || target === 'bold-last' ? alignment : 'justify'];

export function textFitDocx(alignment: keyof typeof fitTexts, neighbor: FitNeighbor,
  kern?: { owner: 'run' | 'defaults' | 'style'; halfPoints: number }, target: FitTarget = 'paragraph',
  compatibilityMode = 15): Uint8Array {
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const run = (text: string, props = '') => `<w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const p = (runs: string, props = '') => `<w:p><w:pPr>${props}</w:pPr>${runs}</w:p>`;
  const kerning = kern ? `<w:kern w:val="${kern.halfPoints}"/>` : '';
  const text = fitText(alignment, target), props = kern?.owner === 'run' ? kerning : '';
  const split = (at: number, length: number, wrap: (runs: string) => string) =>
    run(text.slice(0, at), props) + wrap(run(text.slice(at, at + length), props)) + run(text.slice(at + length), props);
  const bold = (at: number, length: number) => split(at, length, () => run(text.slice(at, at + length), props + '<w:b/>'));
  const last = text.lastIndexOf(' ') + 1;
  const runs = {
    bold: bold(40, 5),
    'bold-the': bold(13, 3),
    'bold-last': bold(last, text.length - last),
    link: split(40, 5, inner => `<w:hyperlink r:id="link">${inner}</w:hyperlink>`),
  }[target as string] ?? run(text, props);
  const jc = `<w:jc w:val="${alignment === 'justify' ? 'both' : 'left'}"/>`;
  const fitted = target === 'list' ? p(runs, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' + jc)
    : target === 'cell' ? '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
      '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' +
      p(runs, jc) + '</w:tc></w:tr></w:tbl>'
    : p(runs, jc);
  const closing = {
    plain: p(run('Closing note.')),
    bold: p(run('A short ') + run('closing', '<w:b/>') + run(' note.')),
    table: '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' + p(run('Cell text.')) + '</w:tc></w:tr></w:tbl>',
    list: p(run('List text.'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
    link: p(`<w:hyperlink r:id="link">${run('Link text.')}</w:hyperlink>`),
  }[neighbor];
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="${wml}.document.main+xml"/>
      ${['styles', 'settings', 'numbering'].map(name => `<Override PartName="/word/${name}.xml" ContentType="${wml}.${name}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      ${['styles', 'settings', 'numbering'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}
      <Relationship Id="link" Type="${R_NS}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>${kern?.owner === 'defaults' ? kerning : ''}<w:sz w:val="24"/>
      </w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${target === 'unspaced' ? ''
        : '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'}
      </w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal">
      <w:name w:val="Normal"/><w:rPr>${kern?.owner === 'style' ? kerning : ''}</w:rPr></w:style></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/></w:compat></w:settings>`,
    'word/numbering.xml': `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>
      <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/>
      </w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` + fitted + closing +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
