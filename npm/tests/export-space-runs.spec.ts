import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const sentence = 'Sorrel records the ferry inventory before the winter crossing. ';
const run = (text: string, size = 20) => `<w:r><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
  <w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

for (const multiple of [1, 2]) for (const position of ['control', 'first', 'middle', 'last']) {
  test(`larger space at ${position} preserves natural ${multiple}x lines and continuation`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const words = sentence.repeat(110).trim();
    const offset = ['control', 'first'].includes(position) ? 0 : position === 'middle' ? words.indexOf('the') : words.length;
    const source = words.slice(0, offset) + ' ' + words.slice(offset);
    const body = `<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0"
      w:line="${240 * multiple}" w:lineRule="auto"/></w:pPr>`
      + run(words.slice(0, offset)) + run(' ', position === 'control' ? 20 : 24) + run(words.slice(offset)) + '</w:p>';
    const result = await renderDocxArtifacts(pageFittingDocx(body), { browser, outputs: ['pdf', 'html'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, text: string[] = [];
      expect(pdf.numPages).toBeGreaterThan(1);
      for (let number = 1; number <= pdf.numPages; number++) {
        const content = await (await pdf.getPage(number)).getTextContent();
        const items = content.items.filter(item => 'str' in item && item.str.trim());
        text.push(...items.map(item => 'str' in item ? item.str : ''));
        const heights = [...new Set(items.map(item => 'transform' in item ? item.transform[5] : 0))].sort((a, b) => b - a);
        expect(heights.length).toBeGreaterThan(3);
        const pitch = 2355 / 2048 * 10 * multiple;
        for (let i = 1; i < heights.length; i++) {
          // Each glyph origin is quantized independently to 0.75pt in Chromium.
          expect(Math.abs(heights[i - 1] - heights[i] - pitch)).toBeLessThanOrEqual(.75);
          expect(Math.abs(heights[0] - heights[i] - pitch * i)).toBeLessThanOrEqual(.75);
        }
      }
      expect(text.join(' ').replace(/\s+/g, ' ').trim()).toBe(source.replace(/\s+/g, ' ').trim());
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const facts = await page.locator('.page-box p[data-source-anchor-id]').evaluateAll(paragraphs => ({
      anchors: paragraphs.map(p => p.getAttribute('data-source-anchor-id')),
      text: paragraphs.map(p => p.textContent).join(''),
      spaces: paragraphs.flatMap(p => Array.from(p.querySelectorAll('span'))).filter(s => s.textContent === ' ')
        .map(span => {
          const range = document.createRange(); range.selectNodeContents(span);
          return { width: range.getBoundingClientRect().width * .75, size: getComputedStyle(span).fontSize };
        }),
    }));
    expect(facts.text).toBe(source);
    expect(new Set(facts.anchors).size).toBe(1);
    expect(facts.anchors.length).toBe(result.pageCount);
    if (position === 'control') return;
    expect(facts.spaces).toHaveLength(1);
    expect(facts.spaces[0].size).toBe('16px');
    expect(facts.spaces[0].width).toBeCloseTo(3, 1);
    expect(result.pageMap.fragments.filter(f => f.anchorId === facts.anchors[0])).toHaveLength(result.pageCount);
  });
}

for (const multiple of [1, 2]) for (const size of [24, 96, 200])
  test(`a ${size / 2}pt space preserves the first baseline and total ${multiple}x line height`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const render = (spaceSize: number) => renderDocxArtifacts(pageFittingDocx(`<w:p><w:pPr>
      <w:spacing w:before="0" w:after="0" w:line="${240 * multiple}" w:lineRule="auto"/></w:pPr>`
      + run('Sorrel') + run(' ', spaceSize) + run('crossing.') + '</w:p>'),
    { browser, outputs: ['pdf', 'html'], reviewProfile: 'final', commentProfile: 'hidden' });
    const facts = async (artifact: Awaited<ReturnType<typeof render>>) => {
      const task = getDocument({ data: new Uint8Array(artifact.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        expect(pdf.numPages).toBe(1);
        const text = await (await pdf.getPage(1)).getTextContent();
        const baselines = text.items.flatMap(item => 'str' in item && item.str.trim() ? [item.transform[5]] : []);
        await page.setContent(artifact.html!);
        const height = await page.locator('.page-box p').evaluate(p => p.getBoundingClientRect().height);
        return { baseline: Math.max(...baselines), height };
      } finally { await task.destroy(); }
    };
    const expected = await facts(await render(20));
    const observed = await facts(await render(size));
    expect(observed.height).toBeCloseTo(expected.height, 3);
    expect(observed.baseline).toBeCloseTo(expected.baseline, 3);
  });
