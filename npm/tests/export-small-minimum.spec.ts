import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const sentence = 'Every autumn the ferry crews repaint the harbor signs before the storms reach our northern pier '
  + 'while the gulls complain about the noise all afternoon long and the old lighthouse keeper counts '
  + 'the coils of rope stacked beside the slipway and writes the tally into a damp ledger. ';

for (const [size, minimum, expected, repeats] of [[10, 13, 13, 2], [10, 13, 13, 32],
  [10, 16, 16, 2], [20, 16, 2355 / 2048 * 20, 2]] as const) {
  test(`${minimum}pt minimum on ${size}pt text preserves pitch and content (${repeats} repeats)`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const text = sentence.repeat(repeats).trim();
    const source = pageFittingDocx('<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
      + `w:line="${minimum * 20}" w:lineRule="atLeast"/></w:pPr><w:r><w:rPr>`
      + '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>'
      + `<w:sz w:val="${size * 2}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`);
    writeFileSync(testInfo.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, { browser, outputs: ['pdf', 'html'],
      reviewProfile: 'final', commentProfile: 'hidden' });
    expect(result.renderReport.status).toBe('complete');
    writeFileSync(testInfo.outputPath('minimum.pdf'), result.pdf!);
    writeFileSync(testInfo.outputPath('minimum.html'), result.html!);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, words: string[] = [], steps: number[] = [];
      expect(pdf.numPages).toBe(repeats === 32 ? 2 : 1);
      for (let number = 1; number <= pdf.numPages; number++) {
        const items = (await (await pdf.getPage(number)).getTextContent()).items
          .filter(item => 'str' in item && item.str.trim());
        const baselines = [...new Set(items.flatMap(item => 'transform' in item ? [item.transform[5]] : []))]
          .sort((a, b) => b - a);
        for (let i = 1; i < baselines.length; i++) steps.push(baselines[i - 1] - baselines[i]);
        words.push(...items.flatMap(item => 'str' in item ? [item.str] : []));
        expect(baselines.at(-1)).toBeGreaterThan(71);
        expect(baselines[0]).toBeLessThan(721);
        // Word's original 16pt-on-10pt PDF starts at 85.780pt from the top:
        // the whole surplus is above the text, not split as CSS leading.
        // Liberation Serif has the same 2355/2048 height and 443/2048 descent.
        const expectedFirst = 792 - 72 - expected + 443 / 2048 * size;
        expect(Math.abs(baselines[0] - expectedFirst)).toBeLessThan(.76);
      }
      expect(words.join(' ').replace(/\s+/g, ' ').trim()).toBe(text);
      expect(steps.length).toBeGreaterThan(3);
      // Each page's two rounded endpoints can change the accumulated pitch by at most 0.75pt.
      expect(Math.abs(steps.reduce((sum, step) => sum + step, 0) - expected * steps.length))
        .toBeLessThan(.76 * pdf.numPages);
      for (const step of steps) expect(Math.abs(step - expected)).toBeLessThan(.76);
    } finally { await task.destroy(); }
  });
}
