import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const cases = [
  { name: 'ordinary paragraphs override explicit values', lists: [0, 0], gap: 14 },
  { name: 'ordinary paragraphs without explicit values', lists: [0, 0], gap: 14, absent: true },
  { name: 'entering a list', lists: [0, 1], gap: 14 },
  { name: 'leaving a list', lists: [1, 0], gap: 14 },
  { name: 'items of one list', lists: [1, 1], gap: 0 },
  { name: 'different lists sharing a definition', lists: [1, 2], gap: 14 },
  { name: 'disabled automatic spacing control', lists: [1, 1], gap: 5, disabled: true },
  { name: 'neighbor-owned explicit spacing', lists: [1, 1], gap: 30, neighbor: true },
  // doNotUseHTMLParagraphAutoSpacing: Word ignores the flags and adds 5pt after to 5pt before.
  { name: 'compatibility setting adds the explicit values', lists: [1, 1], gap: 10, compat: true },
];

function document(entry: typeof cases[number]): Uint8Array {
  const paragraph = (text: string, index: number) => {
    const list = entry.lists[index] ?? 0;
    const automatic = index < 2 && !entry.disabled;
    return '<w:p><w:pPr>' + (list ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list}"/></w:numPr>` : '') +
      '<w:spacing w:line="320" w:lineRule="exact"' +
      (entry.absent ? '' : ` w:before="${entry.neighbor && index === 1 ? 600 : 100}" w:after="100"`) +
      (automatic ? ` w:beforeAutospacing="${entry.neighbor && index === 1 ? 0 : 1}" w:afterAutospacing="1"` : '') +
      `/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  };
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const parts = [
    ['document', 'document.main', `<w:document xmlns:w="${W_NS}"><w:body>` + paragraph('AshOpening', 2) +
      paragraph('AshFirst', 0) + paragraph('AshSecond', 1) + paragraph('AshClosing', 2) +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200"/>' +
      '</w:sectPr></w:body></w:document>'],
    ['styles', 'styles', `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'],
    ['settings', 'settings', `<w:settings xmlns:w="${W_NS}">` +
      (entry.compat ? '<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>' : '') + '</w:settings>'],
    ['numbering', 'numbering', `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>' +
      '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      [1, 2].map(id => `<w:num w:numId="${id}"><w:abstractNumId w:val="0"/>` +
        (id === 2 ? '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' : '') + '</w:num>').join('') + '</w:numbering>'],
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

for (const entry of cases) test(`automatic paragraph spacing: ${entry.name}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(document(entry), { browser, outputs: ['pdf'],
    reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(1);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const content = await (await (await task.promise).getPage(1)).getTextContent();
    const text = content.items.filter(item => 'str' in item);
    for (const label of ['AshOpening', 'AshFirst', 'AshSecond', 'AshClosing'])
      expect(text.filter(item => item.str === label)).toHaveLength(1);
    const first = text.find(item => item.str === 'AshFirst')!;
    const second = text.find(item => item.str === 'AshSecond')!;
    // PDF baselines round within a CSS pixel; the native tests pin exact CSS margins.
    expect(Math.abs(first.transform[5] - second.transform[5] - (16 + entry.gap))).toBeLessThanOrEqual(.5);
  } finally { await task.destroy(); }
});
