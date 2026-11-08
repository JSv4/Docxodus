import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const owner of ['body', 'cell']) for (const count of [12, 216])
for (const [percent, perLine] of [[80, 4], [100, 3], [150, 2]]) {
  test(`${owner} ${count} words at ${percent}% retain scaled glyphs, complete lines and anchors in PDF and offline HTML`, async ({ browser, context }, testInfo) => {
    test.setTimeout(120_000);
    const words = Array.from({ length:count }, (_, i) => `W${String(i).padStart(4, '0')}`);
    const paragraph = `<w:p><w:pPr><w:widowControl w:val="0"/>
      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:right="6480"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:w w:val="${percent}"/>
      <w:sz w:val="24"/></w:rPr><w:t>${words.join(' ')}</w:t></w:r></w:p>`;
    const source = pageFittingDocx(owner === 'body' ? paragraph : '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
      + '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
      + ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="0" w:type="dxa"/>`).join('')
      + '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc>'
      + '<w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' + paragraph + '</w:tc></w:tr></w:tbl>');
    writeFileSync(testInfo.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, { browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden' });
    expect(result.renderReport.status).toBe('complete');
    const task = getDocument({ data:new Uint8Array(result.pdf!), verbosity:0 });
    try {
      const pdf = await task.promise, rows: string[][] = [];
      expect(pdf.numPages).toBe(Math.ceil(words.length / perLine / 47));
      for (let i = 1; i <= pdf.numPages; i++) {
        const content = await (await pdf.getPage(i)).getTextContent();
        const lines = new Map<number, string[]>();
        for (const item of content.items) {
          if (!('str' in item) || !item.str.trim()) continue;
          const y = Math.round(item.transform[5] * 100);
          const line = lines.get(y) ?? [];
          line.push(...item.str.trim().split(/\s+/)); lines.set(y, line);
          // Liberation Mono: each character advances 1229/2048 em. Its
          // glyph height and baseline size stay at 12pt when width changes.
          expect(item.width / item.str.length).toBeCloseTo(1229 / 2048 * 12 * percent / 100, 2);
          expect(item.height).toBeCloseTo(12, 2);
        }
        rows.push(...lines.values());
      }
      expect(rows).toEqual(Array.from({ length:Math.ceil(words.length / perLine) }, (_, i) =>
        words.slice(i * perLine, (i + 1) * perLine)));
    } finally { await task.destroy(); }
    const path = testInfo.outputPath('width.html');
    writeFileSync(path, result.html!);
    const offline = await context.newPage();
    await offline.goto(pathToFileURL(path).href);
    await offline.evaluate(() => document.fonts.ready);
    const fragments = await offline.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('.page-box p[data-source-anchor-id]')).map(p => {
      const page = p.closest<HTMLElement>('.page-box')!, box = p.getBoundingClientRect(), parent = page.getBoundingClientRect();
      return { anchor:p.dataset.sourceAnchorId, page:Number(page.dataset.pageNumber), text:p.textContent!.trim(),
        geometry:{ x:(box.left - parent.left) * .75, y:(box.top - parent.top) * .75, width:box.width * .75, height:box.height * .75 } };
    }));
    expect(fragments.map(p => p.text).join(' ').split(/\s+/)).toEqual(words);
    expect(new Set(fragments.map(p => p.anchor)).size).toBe(1);
    const mappedParagraphs = result.pageMap.fragments.filter(p => p.anchorId === fragments[0].anchor);
    expect(fragments.length).toBe(mappedParagraphs.length);
    for (const [i, p] of fragments.entries()) {
      const mapped = mappedParagraphs[i];
      expect(p.anchor).toBe(mapped.anchorId); expect(p.page).toBe(mapped.pageNumber);
      for (const dimension of ['x', 'y', 'width', 'height'] as const)
        expect(Math.abs(p.geometry[dimension] - mapped.geometry[dimension])).toBeLessThan(.1);
    }
    await offline.close();
  });
}

test('scaled repeated headers keep precision paint in bordered, padded cells', async ({ browser, page }, testInfo) => {
  test.setTimeout(120_000);
  const p = (text:string) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:w w:val="80"/><w:sz w:val="24"/>'
    + `</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const row = (texts:string[], header=false) => '<w:tr>' + (header ? '<w:trPr><w:tblHeader/></w:trPr>' : '')
    + texts.map(text => '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>' + p(text) + '</w:tc>').join('') + '</w:tr>';
  const labels = Array.from({ length:60 }, (_, i) => [`Left${i}`, `Right${i}`]);
  const source = pageFittingDocx('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>'
    + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="4"/>`).join('')
    + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar>'
    + ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="120" w:type="dxa"/>`).join('')
    + '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
    + row(['HeaderLeft', 'HeaderRight'], true) + labels.map(texts => row(texts)).join('') + '</w:tbl>');
  writeFileSync(testInfo.outputPath('source.docx'), source);
  const result = await renderDocxArtifacts(source, { browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden' });
  expect(result.pageCount).toBeGreaterThan(1);
  await page.setContent(result.html!);
  const headers = await page.locator('.page-box tr').evaluateAll(rows => rows.filter(row => row.textContent?.includes('HeaderLeft'))
    .map(row => ({ repeated:row.hasAttribute('data-docx-repeated-header-row'),
      vectorParagraphs:row.querySelectorAll('p > svg').length, cells:row.querySelectorAll('td').length })));
  expect(headers).toHaveLength(result.pageCount);
  expect(headers.every(h => h.cells === 2 && h.vectorParagraphs === 2)).toBe(true);
  expect(headers.filter(h => h.repeated)).toHaveLength(result.pageCount - 1);
  const task = getDocument({ data:new Uint8Array(result.pdf!), verbosity:0 });
  try {
    const pdf = await task.promise, observed:string[] = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const content = await (await pdf.getPage(n)).getTextContent();
      const items = content.items.filter(i => 'str' in i && i.str.trim());
      const text = items.flatMap(i => 'str' in i ? i.str.trim().split(/\s+/) : []);
      expect(text.filter(t => t === 'HeaderLeft')).toHaveLength(1);
      expect(text.filter(t => t === 'HeaderRight')).toHaveLength(1);
      for (const i of items) if ('str' in i) expect(i.width / i.str.length).toBeCloseTo(1229 / 2048 * 12 * .8, 2);
      observed.push(...text.filter(t => !t.startsWith('Header')));
    }
    expect(observed).toEqual(labels.flat());
  } finally { await task.destroy(); }
});
