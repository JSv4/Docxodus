import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath:process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const width of [20, 100]) for (const split of [false, true]) {
  test(`fitText keeps ${split ? 'two contiguous runs' : 'one run'} within ${width}pt before ordinary text`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const run = (text:string, fit=false, bold=false) => `<w:r><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/>
      ${bold ? '<w:b/>' : ''}${fit ? `<w:fitText w:val="${width * 20}" w:id="1"/>` : ''}
      </w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const properties = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>';
    const name = split ? run('Na', true) + run('me', true, true) : run('Name', true);
    const source = pageFittingDocx(`<w:p>${properties}${name}${run(' tail')}</w:p>
      <w:p>${properties}${run('Name')}${run(' tail')}</w:p>`);
    writeFileSync(testInfo.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, { browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden' });
    expect(result.renderReport.status).toBe('complete');
    const html = testInfo.outputPath('fit.html'); writeFileSync(html, result.html!);
    writeFileSync(testInfo.outputPath('fit.pdf'), result.pdf!);
    await page.goto(pathToFileURL(html).href); await page.evaluate(() => document.fonts.ready);
    const boxes = await page.locator('.page-box p[data-source-anchor-id]').evaluateAll(paragraphs => paragraphs.map(p => {
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      const glyphs:DOMRect[] = [];
      for (let node=walker.nextNode(); node; node=walker.nextNode()) for (let i=0; i<node.textContent!.length; i++) {
        const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i+1);
        glyphs.push(range.getBoundingClientRect());
      }
      return { text:p.textContent, width:(glyphs[3].right - glyphs[0].left) * .75,
        tail:(glyphs[5].left - glyphs[0].left) * .75, y:glyphs[0].top, tailY:glyphs[5].top };
    }));
    expect(boxes.map(p => p.text)).toEqual(['Name tail', 'Name tail']);
    expect(boxes[0].width).toBeCloseTo(width, 1);
    // Liberation Serif hmtx advances, at 12pt (independent of the fitted run).
    expect(boxes[1].width).toBeCloseTo(28.65234375, 1);
    expect(boxes[0].tail).toBeCloseTo(width + 3, 1);
    expect(boxes[0].y).toBeCloseTo(boxes[0].tailY, 1);
    const task = getDocument({ data:new Uint8Array(result.pdf!), verbosity:0 });
    try {
      const pdf = await task.promise; expect(pdf.numPages).toBe(1);
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(i => 'str' in i && i.str.trim());
      const firstY = 'transform' in items[0] ? items[0].transform[5] : 0;
      const first = items.filter(i => 'transform' in i && Math.abs(i.transform[5] - firstY) < .1);
      expect(first.map(i => 'str' in i ? i.str : '').join('').replace(/\s/g, '')).toBe('Nametail');
      const fitted = first.filter(i => 'str' in i && !i.str.includes('tail'));
      expect(fitted.length).toBeGreaterThan(0);
      const last = fitted.at(-1)!;
      if ('transform' in fitted[0] && 'transform' in last)
        expect(last.transform[4] + last.width - fitted[0].transform[4]).toBeCloseTo(width, 1);
    } finally { await task.destroy(); }
  });
}

test('a manual-width name stays whole inside a paragraph continued across pages', async ({ browser, page }, testInfo) => {
  test.setTimeout(120_000);
  const labels = Array.from({ length:300 }, (_, i) => `L${String(i).padStart(4, '0')}`);
  const font = '<w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:sz w:val="24"/>';
  const run = (text:string, fit=false) => `<w:r><w:rPr>${font}${fit ? '<w:fitText w:val="2000" w:id="1"/>' : ''}</w:rPr>`
    + `<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const expected = [...labels.slice(0, 150), 'Name', ...labels.slice(150)];
  const source = pageFittingDocx('<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" '
    + 'w:line="240" w:lineRule="auto"/><w:ind w:right="6480"/></w:pPr>'
    + run(labels.slice(0, 150).join(' ') + ' ') + run('Name', true) + run(' ' + labels.slice(150).join(' ')) + '</w:p>');
  writeFileSync(testInfo.outputPath('source.docx'), source);
  const result = await renderDocxArtifacts(source, { browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden' });
  expect(result.renderReport.status).toBe('complete'); expect(result.pageCount).toBeGreaterThan(1);
  const html = testInfo.outputPath('fit.html'); writeFileSync(html, result.html!);
  writeFileSync(testInfo.outputPath('fit.pdf'), result.pdf!);
  await page.goto(pathToFileURL(html).href); await page.evaluate(() => document.fonts.ready);
  const fragments = await page.locator('.page-box p[data-source-anchor-id]').evaluateAll(paragraphs => paragraphs.map(p => {
    const box = p.getBoundingClientRect(), sheet = p.closest('.page-box')!.getBoundingClientRect();
    return { text:p.textContent, anchor:p.getAttribute('data-source-anchor-id'),
      x:(box.left-sheet.left)*.75, y:(box.top-sheet.top)*.75, width:box.width*.75, height:box.height*.75 };
  }));
  expect(fragments.flatMap(p => p.text!.trim().split(/\s+/))).toEqual(expected);
  expect(new Set(fragments.map(p => p.anchor)).size).toBe(1);
  const mapped = result.pageMap.fragments.filter(f => f.anchorId === fragments[0].anchor);
  expect(mapped).toHaveLength(fragments.length);
  for (const [i, p] of fragments.entries()) for (const dimension of ['x', 'y', 'width', 'height'] as const)
    expect(Math.abs(mapped[i].geometry[dimension] - p[dimension])).toBeLessThan(.1);
  const task = getDocument({ data:new Uint8Array(result.pdf!), verbosity:0 });
  try {
    const pdf = await task.promise, actual:string[] = []; let names = 0;
    for (let n=1; n<=pdf.numPages; n++) for (const item of (await (await pdf.getPage(n)).getTextContent()).items) {
      if (!('str' in item) || !item.str.trim()) continue;
      actual.push(...item.str.trim().split(/\s+/));
      if (item.str.trim() === 'Name') { names++; expect(item.width).toBeCloseTo(100, 1); }
    }
    expect(actual).toEqual(expected); expect(names).toBe(1);
  } finally { await task.destroy(); }
});
