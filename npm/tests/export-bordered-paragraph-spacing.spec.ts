import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Issue #1677. Word 2024 draws a paragraph border around the text and w:space only: the space
// before and after lie outside the box and meet the neighbours' spacing like any paragraph's.
// Every expectation below is Word's own print of the same document (16pt exact lines, a 0.5pt
// border with w:space="1", 14pt before and after), measured in November 2026.

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const LINE = 16, SPACE = 14, EDGE = 1.5; // EDGE: border width plus w:space

const box = '<w:pBdr>' + ['top', 'left', 'bottom', 'right'].map(side =>
  `<w:${side} w:val="single" w:sz="4" w:space="1" w:color="000000"/>`).join('') + '</w:pBdr>';

function paragraph(text: string, spaced = false, properties = ''): string {
  const spacing = spaced ? SPACE * 20 : 0;
  return `<w:p><w:pPr>${properties}<w:spacing w:before="${spacing}" w:after="${spacing}" w:line="320" w:lineRule="exact"/></w:pPr>` +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function document(body: string, compat = ''): Uint8Array {
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const parts = [
    ['document', 'document.main', `<w:document xmlns:w="${W_NS}"><w:body>${body}` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
      '</w:sectPr></w:body></w:document>'],
    ['styles', 'styles', `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'],
    ['settings', 'settings', `<w:settings xmlns:w="${W_NS}"><w:compat>${compat}` +
      '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'],
  ];
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      parts.map(([name, type]) => `<Override PartName="/word/${name}.xml" ContentType="${wml}.${type}+xml"/>`).join('') + '</Types>') },
    { name: '_rels/.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      parts.slice(1).map(([name]) => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('') + '</Relationships>') },
    ...parts.map(([name, , data]) => ({ name: `word/${name}.xml`, data: xml(data) })),
  ]);
}

/** Each word's page and distance of its baseline from the top of that page, in points. */
async function baselines(pdf: Uint8Array): Promise<Map<string, { page: number; y: number }>> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const loaded = await task.promise;
    const found = new Map<string, { page: number; y: number }>();
    for (let number = 1; number <= loaded.numPages; number++) {
      const page = await loaded.getPage(number);
      for (const item of (await page.getTextContent()).items)
        if ('str' in item && item.str.trim()) found.set(item.str.trim(), { page: number, y: page.view[3] - item.transform[5] });
    }
    return found;
  } finally { await task.destroy(); }
}

const filler = Array.from({ length: 40 }, (_, index) => paragraph(`Fill${index + 1}`)).join('');

// `steps` are the baseline distances between consecutive named paragraphs.
const cases: { name: string; body: string; compat?: string; words: string[]; steps: number[] }[] = [
  { name: 'a box collapses its spacing with both neighbours',
    body: paragraph('RowanOpen') + paragraph('RowanBoxed', true, box) + paragraph('RowanNext', true) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanBoxed', 'RowanNext', 'RowanClose'],
    steps: [LINE + SPACE + EDGE, LINE + EDGE + SPACE, LINE + SPACE] },
  { name: 'a group keeps its interior spacing and collapses its outer spacing',
    body: paragraph('RowanOpen') + ['RowanOne', 'RowanTwo', 'RowanThree'].map(text => paragraph(text, true, box)).join('') +
      paragraph('RowanNext', true) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanOne', 'RowanTwo', 'RowanThree', 'RowanNext', 'RowanClose'],
    steps: [LINE + SPACE + EDGE, LINE + SPACE, LINE + SPACE, LINE + EDGE + SPACE, LINE + SPACE] },
  { name: 'two different boxes collapse the spacing between them',
    body: paragraph('RowanOpen') + paragraph('RowanBoxed', true, box) +
      paragraph('RowanOther', true, box.replace(/000000/g, 'C00000')) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanBoxed', 'RowanOther', 'RowanClose'],
    steps: [LINE + SPACE + EDGE, LINE + EDGE + SPACE + EDGE, LINE + EDGE + SPACE] },
  // w:doNotUseHTMLParagraphAutoSpacing adds the space after to the next space before (#1553).
  { name: 'the compatibility sum below a box counts its space after once',
    compat: '<w:doNotUseHTMLParagraphAutoSpacing/>',
    body: paragraph('RowanOpen') + paragraph('RowanBoxed', true, box) + paragraph('RowanNext', true) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanBoxed', 'RowanNext', 'RowanClose'],
    steps: [LINE + SPACE + EDGE, LINE + EDGE + 2 * SPACE, LINE + SPACE] },
  { name: 'an unbordered paragraph is unchanged (control)',
    body: paragraph('RowanOpen') + paragraph('RowanBoxed', true) + paragraph('RowanNext', true) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanBoxed', 'RowanNext', 'RowanClose'],
    steps: [LINE + SPACE, LINE + SPACE, LINE + SPACE] },
];

for (const entry of cases) test(`bordered paragraph spacing: ${entry.name}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(document(entry.body, entry.compat), { browser, outputs: ['pdf'],
    reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(1);
  const found = await baselines(result.pdf!);
  const y = entry.words.map(word => found.get(word)!.y);
  // PDF baselines round within a CSS pixel, and CSS lays the 0.5pt border out as 0.75pt.
  entry.steps.forEach((step, index) => expect(Math.abs(y[index + 1] - y[index] - step),
    `${entry.words[index]} to ${entry.words[index + 1]}`).toBeLessThanOrEqual(.75));
});

// A box that overflows to a later page starts at the top of that page: Word drops its space
// before there, as it does an unbordered paragraph's, and keeps it after the first page's text.
for (const [name, properties, edge] of [['box', box, EDGE], ['unbordered control', '', 0]] as const)
  test(`bordered paragraph spacing: no space before a ${name} at the top of a later page`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(
      document(filler + paragraph('RowanBoxed', true, properties) + paragraph('RowanNext', true)),
      { browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(2);
    const found = await baselines(result.pdf!);
    expect(found.get('Fill40')!.page).toBe(1);
    expect(found.get('RowanBoxed')!.page).toBe(2);
    expect(Math.abs(found.get('RowanBoxed')!.y - found.get('Fill1')!.y - edge)).toBeLessThanOrEqual(.75);
    expect(Math.abs(found.get('RowanNext')!.y - found.get('RowanBoxed')!.y - (LINE + edge + SPACE))).toBeLessThanOrEqual(.75);
  });
