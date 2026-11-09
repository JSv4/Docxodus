import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const natural = (size: number) => 2355 / 2048 * size;
const run = (value: string, size: number) => `<w:r><w:rPr>
  <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="${size * 2}"/>
  </w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;

for (const [kind, value, tall, count] of [
  ['auto', 240, 'first', 80], ['auto', 240, 'last', 80], ['auto', 276, 'first', 80],
  ['atLeast', 260, 'first', 80], ['atLeast', 320, 'first', 80], ['atLeast', 320, 'last', 80],
  ['auto', 240, 'none', 80], ['atLeast', 320, 'none', 80],
  ['atLeast', 200, 'first', 80], ['atLeast', 600, 'last', 80],
  ['atLeast', 320, 'first', 800], ['atLeast', 320, 'last', 800],
  ['auto', 276, 'last', 800],
] as const) {
  test(`mixed lines use their own ${kind} ${value} metrics, tall word ${tall}, ${count} words`, async ({ browser, page }, info) => {
    test.setTimeout(120_000);
    const text = Array.from({ length: count }, (_, index) => `Alder${index + 1}x`).join(' ');
    const split = tall === 'first' ? text.indexOf(' ') : tall === 'last' ? text.lastIndexOf(' ') : 0;
    const content = tall === 'none' ? run(text, 10) : tall === 'first'
      ? run(text.slice(0, split), 20) + run(text.slice(split), 10)
      : run(text.slice(0, split + 1), 10) + run(text.slice(split + 1), 20);
    const source = pageFittingDocx(`<w:p><w:pPr><w:spacing w:before="0" w:after="0"
      w:line="${value}" w:lineRule="${kind}"/></w:pPr>${content}</w:p>`);
    writeFileSync(info.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, { browser, outputs: ['pdf', 'html'],
      reviewProfile: 'final', commentProfile: 'hidden' });
    expect(result.renderReport.status).toBe('complete');
    writeFileSync(info.outputPath('actual.pdf'), result.pdf!);
    writeFileSync(info.outputPath('actual.html'), result.html!);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      if (count === 80) expect(pdf.numPages).toBe(1);
      else expect(pdf.numPages).toBeGreaterThan(1);
      await page.setContent(result.html!);
      await page.evaluate(() => document.fonts.ready);
      const allText: string[] = [];
      for (let number = 1; number <= pdf.numPages; number++) {
        const items = (await (await pdf.getPage(number)).getTextContent()).items
          .flatMap(item => 'str' in item && item.str.trim() ? [item] : []);
        allText.push(...items.map(item => item.str));
        // Chromium rounds different-sized glyphs on a shared baseline separately.
        // Count visual lines, retaining a strict one-pixel limit within each line.
        const lines: Array<{ y: number; large: boolean; ys: number[] }> = [];
        for (const item of [...items].sort((a, b) => b.transform[5] - a.transform[5])) {
          const y = item.transform[5];
          let line = lines.at(-1);
          if (!line || line.y - y > .75) lines.push(line = { y, large: false, ys: [] });
          line.large ||= Math.abs(item.transform[0]) > 15;
          line.ys.push(y);
        }
        expect(lines.length).toBeGreaterThan(4);
        for (const line of lines) expect(Math.max(...line.ys) - Math.min(...line.ys)).toBeLessThanOrEqual(.75);
        if (kind === 'atLeast') {
          const size = lines[0].large ? 20 : 10;
          const first = 792 - 72 - Math.max(value / 20, natural(size)) + 443 / 2048 * size;
          expect.soft(Math.abs(lines[0].y - first)).toBeLessThan(.76);
        }
        const ordinary = lines.filter(line => !line.large).map(line => line.y);
        const pitch = kind === 'auto' ? natural(10) * value / 240 : Math.max(natural(10), value / 20);
        for (let index = 1; index < ordinary.length; index++)
          expect.soft(Math.abs(ordinary[index - 1] - ordinary[index] - pitch)).toBeLessThan(.76);
        expect.soft(Math.abs(ordinary[0] - ordinary.at(-1)! - pitch * (ordinary.length - 1))).toBeLessThan(.76);
        for (let index = 1; index < lines.length; index++) {
          const metrics = (large: boolean) => {
            const size = large ? 20 : 10, n = natural(size);
            const h = kind === 'atLeast' ? Math.max(value / 20, n) : n * value / 240;
            const descent = 443 / 2048 * size + (kind === 'atLeast' ? 0 : h - n);
            return { ascent: h - descent, descent };
          };
          const expected = metrics(lines[index - 1].large).descent + metrics(lines[index].large).ascent;
          expect.soft(Math.abs(lines[index - 1].y - lines[index].y - expected)).toBeLessThan(.76);
        }
        {
          const height = await page.locator('.page-content').nth(number - 1).locator('p')
            .evaluate(p => p.getBoundingClientRect().height * .75);
          const largeHeight = kind === 'atLeast' ? Math.max(natural(20), value / 20) : natural(20) * value / 240;
          let expected = pitch * lines.length + lines.filter(line => line.large).length * (largeHeight - pitch);
          if (kind === 'auto' && number < pdf.numPages)
            expected -= natural(lines.at(-1)!.large ? 20 : 10) * (value / 240 - 1);
          // A minimum is applied once to each whole line, with no extra inherited
          // strut and no compression of its larger visible run.
          expect.soft(Math.abs(height - expected)).toBeLessThan(.76);
        }
      }
      expect(allText.join(' ').replace(/\s+/g, ' ').trim()).toBe(text);
    } finally { await task.destroy(); }
  });
}
