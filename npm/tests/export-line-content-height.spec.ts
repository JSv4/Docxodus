import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

// Issue #1565: a line is as tall as the text on it, not as the paragraph's largest run. Each
// expected step is the baseline-to-baseline distance, in points, that desktop Word 2024 printed
// for the same document set in Times New Roman, Liberation Serif's metric twin (recorded in
// docs/ooxml_corner_cases.md). Word places baselines on a 0.12pt grid and the exported PDF
// rounds glyph origins to 0.75pt, so each step is held to 0.76pt and the whole run to the same.

const single = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const multiple = '<w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/>';
const numbered = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
const run = (text: string, props = '') =>
  `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const size = (points: number) => `<w:sz w:val="${points * 2}"/>`;
const paragraph = (content: string, props = single) => `<w:p><w:pPr>${props}</w:pPr>${content}</w:p>`;
const lineBreak = (points?: number) => `<w:r>${points ? `<w:rPr>${size(points)}</w:rPr>` : ''}<w:br/></w:r>`;

const harbor = 'Every autumn the ferry crews repaint the harbor signs before the storms reach our northern pier '
  + 'while the gulls complain about the noise all afternoon long and the old lighthouse keeper counts '
  + 'the coils of rope stacked beside the slipway and writes the tally into a damp ledger that nobody '
  + 'ever reads because the weather has already decided how the season will turn out for everybody.';
const item = 'Every autumn the ferry crews repaint the harbor signs before the storms reach our northern '
  + 'pier while the gulls complain about the noise all afternoon long and nobody minds at all.';
const leadIn = (props: string) => paragraph(run('Harbor.', size(16)) + run(` ${harbor}`, size(9)), props)
  + paragraph(run('Next paragraph.', size(9)));
const breaks = (points: number | undefined, props: string) => paragraph(run('First line at twelve points.')
  + lineBreak(points) + lineBreak(points) + run('Third line after an empty line.') + lineBreak(points)
  + run('Fourth line.'), props) + paragraph(run('Next paragraph.'));
const list = (props: string) => paragraph(run('Opening line.'))
  + Array.from({ length: 4 }, () => paragraph(run(item), numbered + props)).join('')
  + paragraph(run('Closing line.'));
const numbering = (markerProps: string, format = 'decimal', text = '%1.') => '<w:abstractNum w:abstractNumId="0">'
  + `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/>`
  + '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
  + `<w:rPr>${markerProps}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
const face = (font: string) => `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`;
const tasks = ['Repaint the harbor signs', 'Count the coils of rope', 'Check the slipway winch', 'Log the tide tables',
  'Grease the ferry ramp', 'Sweep the pier planks', 'Trim the lantern wick', 'Store the storm shutters'];
/** Eight one-line Liberation Sans items and a wrapped ninth, as in the #1567 generator. */
const bullets = (font: string, glyph: string, props = '', spacing = single) => docx(
  paragraph(run('Opening line.', face('Liberation Sans')))
  + [...tasks, 'Repaint the harbor signs before the storms reach our northern pier while the gulls complain '
    + 'about the noise all afternoon long and nobody on the slipway minds at all.']
    .map(task => paragraph(run(task, face('Liberation Sans')), numbered + spacing)).join('')
  + paragraph(run('Closing line.', face('Liberation Sans'))),
  numbering(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/>${props}`, 'bullet', glyph));

/** A one-page Letter document in 12pt Liberation Serif, as the issue's generators write it. */
function docx(body: string, numberingXml = ''): Uint8Array {
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const relations = (content: string) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + `<Override PartName="/word/document.xml" ContentType="${type}.document.main+xml"/>`
      + `<Override PartName="/word/styles.xml" ContentType="${type}.styles+xml"/>`
      + `<Override PartName="/word/settings.xml" ContentType="${type}.settings+xml"/>`
      + (numberingXml ? `<Override PartName="/word/numbering.xml" ContentType="${type}.numbering+xml"/>` : '')
      + '</Types>',
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>`
      + `<Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/>`
      + (numberingXml ? `<Relationship Id="numbering" Type="${R_NS}/numbering" Target="numbering.xml"/>` : '')),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>`
      + '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/>'
      + '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" `
      + 'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>',
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
      + 'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
  };
  if (numberingXml) parts['word/numbering.xml'] = `<w:numbering xmlns:w="${W_NS}">${numberingXml}</w:numbering>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

const item4 = (into: number, first: number, second: number) => [into, first, second];
const cases: Array<{ name: string; source: Uint8Array; steps: number[] }> = [
  // An empty line made by w:br takes the break's own size, larger or smaller; a break ending a line of
  // text does not size it, and the text after a break keeps its own height.
  { name: 'an empty 8pt break line', source: docx(breaks(8, single)), steps: [23.04, 13.80, 13.80] },
  { name: 'an empty 20pt break line', source: docx(breaks(20, single)), steps: [36.84, 13.80, 13.80] },
  { name: 'an empty 8pt break line at 1.15', source: docx(breaks(8, multiple)), steps: [26.40, 15.96, 15.84] },
  { name: 'small text after breaks', source: docx(paragraph(run('Harbor.', size(16))
    + run(' Small text ends here.', size(9)) + lineBreak(9) + run('Second line of small text.', size(9))
    + lineBreak(9) + run('Third line of small text.', size(9))) + paragraph(run('Next.', size(9)))),
  steps: [11.88, 10.44, 10.32] },
  // A larger list number raises only its own line's ascent; Word adds neither its descent nor a multiple of it.
  { name: 'an 18pt list number', source: docx(list(single), numbering(size(18))),
    steps: [...item4(19.44, 13.80, 13.80), ...item4(19.32, 13.80, 13.80), ...item4(19.44, 13.80, 13.80),
      ...item4(19.44, 13.83, 13.80), 13.80] },
  { name: 'an 18pt list number at 1.15', source: docx(list(multiple), numbering(size(18))),
    steps: [...item4(19.44, 15.84, 15.84), ...item4(21.48, 15.84, 15.96), ...item4(21.48, 15.84, 15.84),
      ...item4(21.51, 15.84, 15.84), 15.96] },
  { name: 'a line break inside items under an 18pt number', source: docx(['01', '02', '03', '04']
    .map(item => paragraph(run(`Aster${item}a`) + lineBreak() + run(`Aster${item}b`), numbered + single)).join(''),
  numbering(size(18))), steps: [13.80, 19.32, 13.80, 19.44, 13.80, 19.44, 13.80] },
  { name: 'a 16pt bullet in the text face', source: bullets('Liberation Sans', '\u2022', size(16)),
    steps: [17.52, 17.52, 17.64, 17.52, 17.52, 17.52, 17.64, 17.52, 17.52, 13.83, 13.80, 13.80] },
  // Alignment does not change how tall a line is.
  { name: 'a centered 16pt lead word', source: docx(leadIn(`${single}<w:jc w:val="center"/>`)),
    steps: [11.88, 10.44, 10.32, 10.32] },
  { name: 'a justified 16pt lead word', source: docx(leadIn(`${single}<w:jc w:val="both"/>`)),
    steps: [11.88, 10.44, 10.32, 10.32] },
  // Controls: unchanged on main.
  { name: 'control: a left-aligned 16pt lead word', source: docx(leadIn(single)), steps: [11.88, 10.44, 10.32, 10.32] },
  { name: 'control: 12pt breaks', source: docx(breaks(undefined, single)), steps: [27.60, 13.80, 13.80] },
  { name: 'control: a plain list number', source: docx(list(single), numbering('')),
    steps: Array(13).fill(13.80) },
  { name: 'control: uniform 9pt text', source: docx(paragraph(run(`Harbor. ${harbor}`, size(9)))
    + paragraph(run('Next paragraph.', size(9)))), steps: [10.32, 10.32, 10.44, 10.32] },
];

/** Baselines of each visual line on each page, in points from its top edge. */
async function pageBaselines(pdf: Uint8Array): Promise<number[][]> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const document = await task.promise;
    const pages: number[][] = [];
    for (let number = 1; number <= document.numPages; number++) {
      const page = await document.getPage(number);
      const height = page.view[3] - page.view[1];
      const ys = (await page.getTextContent()).items
        .flatMap(item => 'str' in item && item.str.trim() ? [height - item.transform[5]] : [])
        .sort((a, b) => a - b);
      // Chromium rounds differently sized glyph origins on one baseline separately.
      const lines: number[][] = [];
      for (const y of ys) {
        if (lines.length && y - lines.at(-1)!.at(-1)! < .8) lines.at(-1)!.push(y);
        else lines.push([y]);
      }
      pages.push(lines.map(line => Math.max(...line)));
    }
    return pages;
  } finally { await task.destroy(); }
}

/** Each step within 0.76pt of Word's, and the run of them too, so rounding cannot accumulate. */
function expectSteps(lines: number[], steps: number[]) {
  const actual = lines.slice(1).map((y, index) => y - lines[index]);
  expect(actual.slice(0, steps.length)).toHaveLength(steps.length);
  for (const [index, step] of steps.entries())
    expect.soft(Math.abs(actual[index] - step), `step ${index + 1}: ${actual[index].toFixed(2)}pt`).toBeLessThan(.76);
  const total = steps.reduce((sum, step) => sum + step, 0);
  expect(Math.abs(lines[steps.length] - lines[0] - total)).toBeLessThan(.76);
}

for (const { name, source, steps } of cases) {
  test(`line steps follow Word: ${name}`, async ({ browser }, info) => {
    test.setTimeout(120_000);
    writeFileSync(info.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden' });
    expect(result.renderReport.status).toBe('complete');
    writeFileSync(info.outputPath('actual.pdf'), result.pdf!);
    const pages = await pageBaselines(result.pdf!);
    expect(pages).toHaveLength(1);
    expect(pages[0]).toHaveLength(steps.length + 1);
    expectSteps(pages[0], steps);
  });
}

// The per-line boxes stay intact when such a paragraph continues on the next page, and the saved
// page lays out exactly as the live one did: the export re-verifies its PageMap offline, and an
// empty line's height that CSS serialized to fewer digits fell one 1/64px layout unit short.
test('line steps follow Word: break lines continuing on the next page', async ({ browser }, info) => {
  test.setTimeout(120_000);
  let content = run('Opening at sixteen points.', size(16));
  for (let line = 1; line <= 45; line++)
    content += run(`Line ${String(line).padStart(2, '0')} of the harbor ledger in twelve point text.`)
      + lineBreak(8) + (line % 3 ? '' : lineBreak(8));
  const source = docx(paragraph(content) + paragraph(run('After the long paragraph.')));
  writeFileSync(info.outputPath('source.docx'), source);
  const result = await renderDocxArtifacts(source, { browser, outputs: ['pdf'],
    reviewProfile: 'final', commentProfile: 'hidden' });
  expect(result.renderReport.status).toBe('complete');
  writeFileSync(info.outputPath('actual.pdf'), result.pdf!);
  const pages = await pageBaselines(result.pdf!);
  // Word sets lines 1 to 38 on the first page and 39 to 45 on the second.
  expect(pages.map(page => page.length)).toEqual([38, 8]);
  const third = (empty: number) => [13.80, empty, 13.80];
  expectSteps(pages[0], [14.64, ...[23.04, 23.04, 22.92, 23.04, 22.92, 23.04, 23.04, 22.92, 23.04, 23.07, 22.92, 23.04]
    .flatMap(third)]);
  // Word's last step, 36.84, also crosses the line its paragraph mark sits on after the final break.
  expectSteps(pages[1], [23.04, 13.80, 13.80, 22.92, 13.80, 13.80]);
});
