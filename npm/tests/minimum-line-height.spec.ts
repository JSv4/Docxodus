import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('a minimum below uniform natural height preserves the natural line and repeated preparation', async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async () => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('div');
    root.innerHTML = '<p style="font:20pt/16pt Liberation Serif;--docx-min-line-height:16pt;margin:0">'
      + '<span>Nettle records remain legible.</span></p>';
    document.body.append(root);
    await document.fonts.ready;
    const paragraph = root.querySelector('p')!;
    const measure = () => [paragraph, paragraph.firstElementChild!].map(element =>
      Number.parseFloat(getComputedStyle(element).lineHeight));
    applyUnroundedNormalLineHeights(root);
    const first = measure();
    applyUnroundedNormalLineHeights(root);
    return { first, repeated: measure() };
  });
  for (const height of result.first) expect(height * .75).toBeCloseTo(2355 / 2048 * 20, 3);
  expect(result.repeated).toEqual(result.first);
});

for (const [name, content] of [
  ['mixed sizes', '<span>Large text</span><span style="font-size:10pt">small text</span>'],
  ['mixed families', '<span>Serif text</span><span style="font-family:Liberation Sans">sans text</span>'],
  ['raised text', '<span style="vertical-align:super">Raised text</span>'],
  ['explicit breaks', '<span>First line<br>Second line</span>'],
  ['empty paragraph', '<span>&nbsp;</span>'],
  ['image', '<span>Picture <img width="1" height="1"></span>'],
]) test(`minimum spacing retains the existing ${name} policy`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async markup => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('div');
    root.innerHTML = '<p style="font:20pt/16pt Liberation Serif;--docx-min-line-height:16pt">' + markup + '</p>';
    document.body.append(root);
    const elements = Array.from(root.querySelectorAll('p,span'));
    const before = elements.map(e => getComputedStyle(e).lineHeight);
    applyUnroundedNormalLineHeights(root);
    return { before, after: elements.map(e => getComputedStyle(e).lineHeight) };
  }, content);
  expect(result.after).toEqual(result.before);
});

for (const [name, wrapper] of [
  ['footnote', '<div data-footnote-id="1"><span class="footnote-content">CONTENT</span></div>'],
  ['comment', '<div data-comment-id="1">CONTENT</div>'],
  ['header registry', '<div id="pagination-hf-registry" style="display:none"><div data-hf-type="header-default">CONTENT</div></div>'],
  ['footer', '<div data-hf-type="footer-default">CONTENT</div>'],
  ['frame', '<div data-docx-text-frame="true">CONTENT</div>'],
  ['cell', '<table><tr><td>CONTENT</td></tr></table>'],
]) test(`minimum spacing retains the ${name} owner policy`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async wrapper => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('div');
    root.innerHTML = wrapper.replace('CONTENT', '<p style="font:20pt/16pt Liberation Serif;--docx-min-line-height:16pt">'
      + '<span>Nettle owner control.</span></p>');
    document.body.append(root);
    const paragraph = root.querySelector('p')!;
    const before = getComputedStyle(paragraph).lineHeight;
    applyUnroundedNormalLineHeights(root);
    return { before, after: getComputedStyle(paragraph).lineHeight };
  }, wrapper);
  expect(result.after).toBe(result.before);
});

for (const [name, minimum, size, height] of [
  ['already natural', 10, 20, 'normal'], ['unresolved small minimum', 13, 10, 'normal'],
  ['larger minimum', 16, 10, '16pt'], ['explicit height override', 16, 20, '18pt'],
] as const) test(`minimum spacing preserves the ${name} control`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async ({ minimum, size, height }) => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('div');
    root.innerHTML = `<p style="font:${size}pt/${height} Liberation Serif"><span>Nettle control.</span></p>`;
    const control = root.firstElementChild!, actual = control.cloneNode(true) as HTMLElement;
    actual.style.setProperty('--docx-min-line-height', `${minimum}pt`);
    root.append(actual); document.body.append(root);
    applyUnroundedNormalLineHeights(root);
    const heights = (p: Element) => [p, p.firstElementChild!].map(e => getComputedStyle(e).lineHeight);
    return { control: heights(control), actual: heights(actual) };
  }, { minimum, size, height });
  expect(result.actual).toEqual(result.control);
});

for (const repeats of [8, 40]) test(`a 16pt minimum cannot compress 20pt exported text (${repeats} repetitions)`, async ({ browser }) => {
  test.setTimeout(120_000);
  const words = 'Nettle records the harbor inventory before the winter crossing. '.repeat(repeats);
  const paragraph = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
    + 'w:line="320" w:lineRule="atLeast"/></w:pPr><w:r><w:rPr><w:sz w:val="40"/>'
    + '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/></w:rPr>'
    + `<w:t xml:space="preserve">${words}</w:t></w:r></w:p>`;
  const docx = pageFittingDocx(paragraph);
  const control = await renderDocxArtifacts(pageFittingDocx(paragraph.replace('w:line="320" w:lineRule="atLeast"',
    'w:line="240" w:lineRule="auto"')), { browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden' });
  const result = await renderDocxArtifacts(docx, { browser, outputs: ['pdf'], reviewProfile: 'final',
    commentProfile: 'hidden', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  const controlTask = getDocument({ data: new Uint8Array(control.pdf!), verbosity: 0 });
  try {
    const controlText = await (await (await controlTask.promise).getPage(1)).getTextContent();
    const firstBaseline = Math.max(...controlText.items.flatMap(item => 'str' in item && item.str.trim() ? [item.transform[5]] : []));
    const pdf = await task.promise, text: string[] = [];
    for (let number = 1; number <= pdf.numPages; number++) {
      const content = await (await pdf.getPage(number)).getTextContent();
      const items = content.items.filter(item => 'str' in item && item.str.trim());
      const heights = [...new Set(items.map(item => 'transform' in item ? item.transform[5] : 0))].sort((a, b) => b - a);
      expect(heights.length).toBeGreaterThan(3);
      if (number === 1) expect(heights[0]).toBeCloseTo(firstBaseline, 0);
      for (let i = 1; i < heights.length; i++) expect(heights[i - 1] - heights[i]).toBeCloseTo(2355 / 2048 * 20, 0);
      text.push(...items.map(item => 'str' in item ? item.str : ''));
    }
    expect(text.join(' ').replace(/\s+/g, ' ').trim()).toBe(words.trim());
    expect(pdf.numPages).toBe(control.pageCount);
  } finally { await task.destroy(); await controlTask.destroy(); }
});
