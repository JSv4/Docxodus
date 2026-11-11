import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Issue #1677. Word 2024 draws a paragraph border around the text and w:space only: the space
// before and after lie outside the box and meet the neighbours' spacing like any paragraph's.
// Every expectation below is Word's own print of the same document (16pt exact lines, a 0.5pt
// border with w:space="1", 14pt before and after), measured in November 2026.
// Issue #1692 adds what else Word does with that box: the rule each paragraph's w:between draws
// below it, contextual spacing across the box's edge, a page start or a keep on a bordered
// paragraph, and shading that fills the box to its border.

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const LINE = 16, SPACE = 14, EDGE = 1.5; // EDGE: border width plus w:space

const box = '<w:pBdr>' + ['top', 'left', 'bottom', 'right'].map(side =>
  `<w:${side} w:val="single" w:sz="4" w:space="1" w:color="000000"/>`).join('') + '</w:pBdr>';

const ruled = box.replace('</w:pBdr>', '<w:between w:val="single" w:sz="4" w:space="1" w:color="000000"/></w:pBdr>');
const contextual = '<w:contextualSpacing/>';
const shading = '<w:shd w:val="clear" w:color="auto" w:fill="B0C4DE"/>';

/** `properties` precede w:spacing in the schema's order (keeps, borders, shading); `after` follows it. */
function paragraph(text: string, spaced = false, properties = '', after = ''): string {
  const spacing = spaced ? SPACE * 20 : 0;
  return `<w:p><w:pPr>${properties}<w:spacing w:before="${spacing}" w:after="${spacing}" w:line="320" w:lineRule="exact"/>${after}</w:pPr>` +
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
  // w:between (#1692): w:space, the rule and w:space again separate the members, below the spacing.
  { name: 'each w:between rule and its space separate the members of a box',
    body: paragraph('RowanOpen') + ['RowanOne', 'RowanTwo', 'RowanThree'].map(text => paragraph(text, true, ruled)).join('') +
      paragraph('RowanNext', true) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanOne', 'RowanTwo', 'RowanThree', 'RowanNext', 'RowanClose'],
    steps: [LINE + SPACE + EDGE, LINE + SPACE + 1 + EDGE, LINE + SPACE + 1 + EDGE, LINE + EDGE + SPACE, LINE + SPACE] },
  // A paragraph without w:between stays in the box: the rule above it and that rule's space below
  // come from the paragraph above, and only the space above the missing rule from the one below.
  { name: 'paragraphs that differ only in w:between share one box',
    body: paragraph('RowanOpen') + paragraph('RowanOne', true, ruled) + paragraph('RowanTwo', true, box) +
      paragraph('RowanThree', true, ruled) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanOne', 'RowanTwo', 'RowanThree', 'RowanClose'],
    steps: [LINE + SPACE + EDGE, LINE + SPACE + EDGE, LINE + SPACE + 1, LINE + EDGE + SPACE] },
  // w:contextualSpacing (#1692): one style throughout, so no spacing survives on either side of the box.
  { name: 'contextual spacing applies across the edge of a box',
    body: paragraph('RowanOpen', false, '', contextual) + paragraph('RowanBoxed', true, box, contextual) +
      paragraph('RowanNext', true, '', contextual) + paragraph('RowanClose'),
    words: ['RowanOpen', 'RowanBoxed', 'RowanNext', 'RowanClose'],
    steps: [LINE + EDGE, LINE + EDGE, LINE] },
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

// A page start or a keep on a bordered paragraph (#1692). `lead` unspaced lines precede the box;
// 39 of them leave room for the box but not for the line after it. Word starts the box on page 2
// in the first three and leaves it on page 1 in the control. A page start inside a run of bordered
// paragraphs closes the box before it.
const pageBreak = '<w:pageBreakBefore/>', keepNext = '<w:keepNext/>';
for (const [name, lead, body, page] of [
  ['a bordered paragraph with a page start', 1, paragraph('RowanBoxed', false, pageBreak + box), 2],
  ['a page start inside a run of bordered paragraphs', 1,
    paragraph('RowanFirst', false, box) + paragraph('RowanBoxed', false, pageBreak + box), 2],
  ['a bordered paragraph kept with a next paragraph that does not fit', 39, paragraph('RowanBoxed', false, keepNext + box), 2],
  ['a bordered paragraph with no keep (control)', 39, paragraph('RowanBoxed', false, box), 1],
] as const) test(`bordered paragraph pagination: ${name}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const fill = Array.from({ length: lead }, (_, index) => paragraph(`Fill${index + 1}`)).join('');
  const result = await renderDocxArtifacts(document(fill + body + paragraph('RowanNext')),
    { browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(2);
  const found = await baselines(result.pdf!);
  expect(found.get('RowanBoxed')!.page).toBe(page);
  expect(found.get('RowanNext')!.page).toBe(2);
  // On page 2 the box starts at the top margin, one border and w:space above its text.
  if (page === 2) expect(Math.abs(found.get('RowanBoxed')!.y - found.get('Fill1')!.y - EDGE)).toBeLessThanOrEqual(.75);
});

// Where a page divides a box (#1692). Word leaves the box open at the break and starts the next
// page with the continuing text, at the top margin.
test('bordered paragraph pagination: a paragraph divided by the page does not repeat the rule above it', async ({ browser }) => {
  test.setTimeout(120_000);
  // 37 lines leave room for the box's first paragraph, its rule and two of the next paragraph's four lines.
  const fill = Array.from({ length: 37 }, (_, index) => paragraph(`Fill${index + 1}`)).join('');
  const lines = paragraph('RowanL1', false, ruled).replace('</w:r>',
    '</w:r>' + [2, 3, 4].map(n => `<w:r><w:br/><w:t>RowanL${n}</w:t></w:r>`).join(''));
  const result = await renderDocxArtifacts(
    document(fill + paragraph('RowanFirst', false, ruled) + lines + paragraph('RowanLast', false, ruled)),
    { browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(2);
  const found = await baselines(result.pdf!);
  expect([found.get('RowanL2')!.page, found.get('RowanL3')!.page]).toEqual([1, 2]);
  expect(Math.abs(found.get('RowanL3')!.y - found.get('Fill1')!.y)).toBeLessThanOrEqual(.4);
  // The next whole paragraph still has its rule and both spaces above it.
  expect(Math.abs(found.get('RowanLast')!.y - found.get('RowanL4')!.y - (LINE + 1 + EDGE))).toBeLessThanOrEqual(.75);
});

test('bordered paragraph pagination: a box that starts a page and is taller than it continues on the next', async ({ browser }) => {
  test.setTimeout(120_000);
  // 40 lines and the top border fill page 2; five lines remain for page 3.
  const rows = Array.from({ length: 45 }, (_, index) => paragraph(`Row${index + 1}`, false, (index ? '' : pageBreak) + box)).join('');
  const result = await renderDocxArtifacts(document(paragraph('Fill1') + rows),
    { browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(3);
  const found = await baselines(result.pdf!);
  expect([found.get('Row1')!.page, found.get('Row40')!.page, found.get('Row41')!.page]).toEqual([2, 2, 3]);
  expect(Math.abs(found.get('Row1')!.y - found.get('Fill1')!.y - EDGE)).toBeLessThanOrEqual(.75);
  expect(Math.abs(found.get('Row41')!.y - found.get('Fill1')!.y)).toBeLessThanOrEqual(.4);
});

/** The bounds, in points from the page's top left, of everything page 1 fills with `colour` (`#rrggbb`). */
async function filled(pdf: Uint8Array, colour: string): Promise<{ top: number; bottom: number; left: number; right: number }> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const page = await (await task.promise).getPage(1);
    const list = await page.getOperatorList();
    let current = '', bounds: number[] | undefined;
    list.fnArray.forEach((op, index) => {
      if (op === OPS.setFillRGBColor) current = String(list.argsArray[index][0]).toLowerCase();
      // constructPath names the operator that paints the path and ends with the path's bounding
      // box. Chromium fills in CSS pixels from the page's top left; its clip paths are not fills.
      if (op !== OPS.constructPath || list.argsArray[index][0] !== OPS.fill || current !== colour) return;
      const [left, top, right, bottom] = Array.from(list.argsArray[index][2] as ArrayLike<number>);
      bounds = bounds ? [Math.min(bounds[0], left), Math.min(bounds[1], top), Math.max(bounds[2], right), Math.max(bounds[3], bottom)]
        : [left, top, right, bottom];
    });
    expect(bounds, `a ${colour} fill`).toBeDefined();
    const [left, top, right, bottom] = bounds!.map(pixels => pixels * .75);
    return { top, bottom, left, right };
  } finally { await task.destroy(); }
}

// Shading (#1692): Word fills a box to its border, through the w:space band, and only the lines of
// an unbordered paragraph. The body starts 72pt down and in; the shaded paragraph follows one line.
for (const [name, properties, top, bottom, outset] of [
  ['fills a box to its border', box + shading, 72 + LINE + SPACE + .5, 72 + LINE + SPACE + EDGE + LINE + 1, 1],
  ['covers the lines of an unbordered paragraph (control)', shading, 72 + LINE + SPACE, 72 + LINE + SPACE + LINE, 0],
] as const) test(`bordered paragraph shading ${name}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(
    document(paragraph('RowanOpen') + paragraph('RowanBoxed', true, properties) + paragraph('RowanClose')),
    { browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  const fill = await filled(result.pdf!, '#b0c4de');
  expect(Math.abs(fill.top - top), 'top').toBeLessThanOrEqual(.75);
  expect(Math.abs(fill.bottom - bottom), 'bottom').toBeLessThanOrEqual(.75);
  expect(Math.abs(fill.left - (72 - outset)), 'left').toBeLessThanOrEqual(.75);
  expect(Math.abs(fill.right - (540 + outset)), 'right').toBeLessThanOrEqual(.75);
});
