import { storedZip, xml } from './docx-zip.js';

/**
 * Seeded, synthetic documents that mix the content shapes behind the paginated export's past
 * failures (epic #846): fonts named only through an absent theme (#847), paragraphs whose only
 * content is a floating text box (#849), tables after zero-spacing paragraphs and exact line spacing
 * below the font height (#848), plus tab-aligned label lines, lists, keep-with-next chains, borders and
 * running headers and footers. The same seed always yields the same bytes.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

const WORDS = ('the parties agree that the supplier shall deliver the services described in the schedule '
  + 'within thirty days of notice and any material breach may be cured before termination').split(' ');

function random(seed: number): () => number {
  let state = seed;
  return () => {
    state = (state * 1103515245 + 12345) & 0x7fffffff;
    return state / 0x7fffffff;
  };
}

export function mixedContentDocx(seed: number): Uint8Array {
  const next = random(seed);
  const pick = <T>(values: T[]): T => values[Math.floor(next() * values.length)];
  const words = (count: number) => Array.from({ length: count }, () => pick(WORDS)).join(' ');
  const themeOnlyFonts = next() < 0.3;

  const paragraph = (text: string, extra = ''): string => {
    const lineRoll = next();
    const line = lineRoll < 0.2
      ? `w:line="${pick([240, 276, 360])}" w:lineRule="auto"`
      : lineRoll < 0.35 ? `w:line="${pick([240, 280])}" w:lineRule="exact"` : '';
    const keepNext = next() < 0.2 ? '<w:keepNext/>' : '';
    const border = next() < 0.06
      ? '<w:pBdr><w:top w:val="single" w:sz="8" w:space="4" w:color="000000"/></w:pBdr>' : '';
    const size = pick([18, 20, 22, 24, 28, 32]);
    return `<w:p><w:pPr>${keepNext}${border}${extra}<w:spacing w:before="${pick([0, 0, 120, 240])}" `
      + `w:after="${pick([0, 0, 120, 160])}" ${line}/></w:pPr>`
      + `<w:r><w:rPr><w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  };
  // A short label and an amount on one tab stop, as forms and price lists use. A tab after text
  // longer than its line is #891 and stays out of the generator until it is fixed.
  const tabLine = (): string => `<w:p><w:pPr><w:tabs><w:tab w:val="${pick(['left', 'right', 'decimal'])}" w:pos="7200"/></w:tabs></w:pPr>`
    + `<w:r><w:t>${pick(['Total', 'Fee', 'Balance due'])}</w:t></w:r><w:r><w:tab/><w:t>${Math.floor(next() * 9000) / 100}</w:t></w:r></w:p>`;

  const textBox = (): string => `<w:p><w:r><w:drawing>
    <wp:anchor distT="0" distR="114300" distB="0" distL="114300" simplePos="0" relativeHeight="10"
      behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="${pick(['paragraph', 'page', 'margin'])}"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="2286000" cy="685800"/><wp:wrapSquare wrapText="bothSides"/>
      <wp:docPr id="${Math.floor(next() * 1000) + 1}" name="Text Box"/><wp:cNvGraphicFramePr/>
      <a:graphic><a:graphicData uri="${WPS}"><wps:wsp>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="685800"/></a:xfrm></wps:spPr>
        <wps:txbx><w:txbxContent><w:p><w:r><w:t>${words(6)}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
        <wps:bodyPr/>
      </wps:wsp></a:graphicData></a:graphic>
    </wp:anchor></w:drawing></w:r></w:p>`;

  const table = (): string => {
    const rows = 1 + Math.floor(next() * 4);
    const cell = () => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph(words(3 + Math.floor(next() * 20)))}</w:tc>`;
    return '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4"/>'
      + '<w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
      + Array.from({ length: rows }, () => `<w:tr>${cell()}${cell()}</w:tr>`).join('') + '</w:tbl>';
  };

  let body = '';
  const blocks = 25 + Math.floor(next() * 40);
  for (let i = 0; i < blocks; i++) {
    const roll = next();
    if (roll < 0.08) body += table();
    else if (roll < 0.12) body += textBox();
    else if (roll < 0.17) body += tabLine();
    else if (roll < 0.24) {
      body += paragraph(words(4 + Math.floor(next() * 30)),
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    } else body += paragraph(words(3 + Math.floor(next() * 90)));
  }

  const runningStories = next() < 0.5;
  const section = `<w:sectPr>${runningStories
    ? '<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>' : ''}`
    + `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="${pick([1080, 1440])}" w:right="1440" `
    + `w:bottom="${pick([1080, 1440])}" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
  const fonts = themeOnlyFonts
    ? '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
    : '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>';
  const story = (tag: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} xmlns:w="${W}"><w:p><w:r><w:t>${tag} ${words(5)}</w:t></w:r></w:p></w:${tag}>`;

  const entries = [
    {
      name: '[Content_Types].xml',
      data: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        + '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        + (runningStories
          ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            + '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
          : '')
        + '</Types>'),
    },
    {
      name: '_rels/.rels',
      data: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        + (runningStories
          ? '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            + '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
          : '')
        + '</Relationships>'),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${fonts}<w:sz w:val="22"/></w:rPr></w:rPrDefault>`
        + '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
        + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'),
    },
    {
      name: 'word/numbering.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">`
        + `<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="${pick(['left', 'right', 'center'])}"/>`
        + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}"><w:body>${body}${section}</w:body></w:document>`),
    },
  ];
  if (runningStories) {
    entries.push({ name: 'word/header1.xml', data: xml(story('hdr')) });
    entries.push({ name: 'word/footer1.xml', data: xml(story('ftr')) });
  }
  return storedZip(entries);
}

/** A two-column table whose first cell holds a sentence longer than its line followed by a tab (#891). */
export function tabAfterWrappingTextDocx(): Uint8Array {
  const sentence = 'the parties agree that the supplier shall deliver the services described in the schedule within thirty days';
  const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
  const table = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
    + `<w:tr>${cell(`<w:p><w:r><w:t>${sentence}</w:t></w:r><w:r><w:tab/><w:t>12.50</w:t></w:r></w:p>`)}${cell('<w:p><w:r><w:t>Second cell</w:t></w:r></w:p>')}</w:tr></w:tbl>`;
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    },
    {
      name: '_rels/.rels',
      data: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p>${table}`
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'),
    },
  ]);
}
