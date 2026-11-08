import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const words = Array.from({ length:216 }, (_, i) => `W${String(i).padStart(4, '0')}`);
for (const [percent, perLine] of [[80, 4], [100, 3], [150, 2]]) {
  test(`${percent}% text retains scaled glyphs, complete lines and anchors in PDF and offline HTML`, async ({ browser, context }, testInfo) => {
    test.setTimeout(120_000);
    const source = pageFittingDocx(`<w:p><w:pPr><w:widowControl w:val="0"/>
      <w:ind w:right="6480"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:sz w:val="24"/>
      <w:w w:val="${percent}"/></w:rPr><w:t>${words.join(' ')}</w:t></w:r></w:p>`);
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
    expect(fragments.length).toBe(result.pageMap.fragments.length);
    for (const [i, p] of fragments.entries()) {
      const mapped = result.pageMap.fragments[i];
      expect(p.anchor).toBe(mapped.anchorId); expect(p.page).toBe(mapped.pageNumber);
      for (const dimension of ['x', 'y', 'width', 'height'] as const)
        expect(Math.abs(p.geometry[dimension] - mapped.geometry[dimension])).toBeLessThan(.1);
    }
    await offline.close();
  });
}
