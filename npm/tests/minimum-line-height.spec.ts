import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const [owner, wrapper] of [
  ['body', 'CONTENT'], ['cell', '<table><tr><td>CONTENT</td></tr></table>'],
  ['footnote', '<div data-footnote-id="1"><span class="footnote-content">CONTENT</span></div>'],
  ['endnote', '<section class="endnotes"><ol><li>CONTENT</li></ol></section>'],
  ['header registry', '<div id="pagination-hf-registry" style="display:none"><div data-hf-type="header-default">CONTENT</div></div>'],
  ['footer', '<div data-hf-type="footer-default">CONTENT</div>'],
]) test(`a ${owner} minimum below uniform natural height preserves repeated preparation`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async wrapper => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('div');
    root.innerHTML = '<p style="font:20pt/16pt Liberation Serif;--docx-min-line-height:16pt;margin:0">'
      + '<span>Nettle records remain legible.</span></p>';
    root.innerHTML = wrapper.replace('CONTENT', root.innerHTML);
    const note = root.querySelector('section.endnotes p');
    if (note) note.innerHTML = '<span data-source-anchor-id="en:en:0123456789abcdef0123456789abcdef">' + note.innerHTML + '</span>';
    document.body.append(root);
    await document.fonts.ready;
    const paragraph = root.querySelector('p')!;
    const measure = () => [paragraph, ...Array.from(paragraph.querySelectorAll('span'))].map(element =>
      Number.parseFloat(getComputedStyle(element).lineHeight));
    applyUnroundedNormalLineHeights(root);
    const first = measure();
    applyUnroundedNormalLineHeights(root);
    const clone = root.cloneNode(true) as HTMLElement;
    document.body.append(clone); applyUnroundedNormalLineHeights(clone);
    const copied = clone.querySelector('p')!;
    return { first, repeated: measure(), cloned:[copied, ...Array.from(copied.querySelectorAll('span'))]
      .map(element => Number.parseFloat(getComputedStyle(element).lineHeight)),
      auto:copied.style.getPropertyValue('--docx-auto-line-spacing'), minimum:copied.style.getPropertyValue('--docx-min-line-height') };
  }, wrapper);
  for (const height of result.first) expect(height * .75).toBeCloseTo(2355 / 2048 * 20, 3);
  expect(result.repeated).toEqual(result.first);
  expect(result.cloned).toEqual(result.first);
  expect(result.auto).toBe('1'); expect(result.minimum).toBe('16pt');
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
    return { before, after: elements.map(e => getComputedStyle(e).lineHeight),
      automatic:elements.map(e => getComputedStyle(e).getPropertyValue('--docx-auto-line-spacing')) };
  }, content);
  expect(result.after).toEqual(result.before);
  expect(result.automatic.every(value => !value)).toBe(true);
});

for (const [name, wrapper] of [
  ['comment', '<div data-comment-id="1">CONTENT</div>'],
  ['frame', '<div data-docx-text-frame="true">CONTENT</div>'],
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
  ['already natural', 10, 20, 'normal'],
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
    return { control: heights(control), actual: heights(actual), automatic:actual.style.getPropertyValue('--docx-auto-line-spacing') };
  }, { minimum, size, height });
  expect(result.actual).toEqual(result.control);
  expect(result.automatic).toBe('');
});

for (const [name, minimum, height, markup, recovers] of [
  ['12pt minimum', 12, 'normal', '<span>Small minimum.</span>', true],
  ['13pt minimum', 13, 'normal', '<span>Small minimum.</span>', true],
  ['fractional minimum', 13.5, 'normal', '<span>Small minimum.</span>', true],
  ['explicit override', 13, '12pt', '<span>Small minimum.</span>', false],
  ['explicit natural height', 13, `${2355 / 2048 * 10}pt`, '<span>Small minimum.</span>', false],
  ['explicit run height', 13, 'normal', `<span style="line-height:${2355 / 2048 * 10}pt">Small minimum.</span>`, false],
  ['explicit paragraph height', 13, `${2355 / 2048 * 10}pt`, '<span style="line-height:normal">Small minimum.</span>', false],
  ['mixed sizes', 13, 'normal', '<span>Text</span><span style="font-size:9pt">small</span>', false],
  ['mixed fonts', 13, 'normal', '<span>Text</span><span style="font-family:Liberation Sans">sans</span>', false],
  ['raised run', 13, 'normal', '<span style="vertical-align:super">Raised text.</span>', false],
] as const) test(`small minimum respects ${name}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async ({ minimum, height, markup, recovers }) => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const actual = document.createElement('p');
    actual.style.cssText = `font:10pt/${height} Liberation Serif;--docx-min-line-height:${minimum}pt`;
    actual.innerHTML = markup;
    const control = actual.cloneNode(true) as HTMLElement;
    control.style.removeProperty('--docx-min-line-height');
    if (recovers) control.style.lineHeight = `${minimum}pt`;
    document.body.append(actual, control);
    const measure = (p: HTMLElement) => [p, ...Array.from(p.querySelectorAll('span'))].map(e => ({
      height: getComputedStyle(e).lineHeight, top: getComputedStyle(e).top,
      auto: (e as HTMLElement).style.getPropertyValue('--docx-auto-line-spacing'),
    }));
    applyUnroundedNormalLineHeights(control);
    const count = applyUnroundedNormalLineHeights(actual), first = measure(actual);
    const repeatCount = applyUnroundedNormalLineHeights(actual), repeated = measure(actual);
    const clone = actual.cloneNode(true) as HTMLElement; document.body.append(clone);
    const cloneCount = applyUnroundedNormalLineHeights(clone);
    return { count, repeatCount, cloneCount, first, repeated, clone: measure(clone), expected: measure(control) };
  }, { minimum, height, markup, recovers });
  expect(result.first).toEqual(result.expected);
  expect(result.repeated).toEqual(result.first);
  expect(result.clone).toEqual(result.first);
  expect(result.repeatCount).toBe(0); expect(result.cloneCount).toBe(0);
  if (recovers) expect(result.count).toBe(2);
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

const uniformMarker = '<span data-list-marker="true" style="display:inline-flex;align-items:baseline;width:24px">'
  + '<span><span>1.</span></span><span style="display:inline-block;flex:1">'
  + '<span data-docx-tab="left" style="display:inline-block;width:100%;height:0"></span></span></span>';

for (const [name, marker, restored] of [
  ['flat number', '<span data-list-marker="true">1.</span>', true],
  ['uniform number', uniformMarker, true],
  ['larger number', uniformMarker.replace('<span>1.', '<span style="font-size:30pt">1.'), false],
  ['raised number', uniformMarker.replace('<span>1.', '<span style="vertical-align:super">1.'), false],
  ['positioned number', uniformMarker.replace('<span>1.', '<span style="position:relative;top:2px">1.'), false],
  ['nested link', uniformMarker.replace('<span>1.</span>', '<a href="#">1.</a>'), false],
] as const) test(`a minimum preserves the ${name} metric contract`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async marker => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('div');
    root.innerHTML = '<p style="font:20pt/16pt Liberation Serif;--docx-min-line-height:16pt">'
      + marker + '<span>Teasel records remain legible.</span></p>';
    document.body.append(root);
    const elements = Array.from(root.querySelectorAll('p,span,a'));
    const measure = () => elements.map(e => getComputedStyle(e).lineHeight);
    const before = measure();
    applyUnroundedNormalLineHeights(root);
    const after = measure();
    applyUnroundedNormalLineHeights(root);
    return { before, after, repeated: measure() };
  }, marker);
  if (restored) for (const height of result.after) expect(Number.parseFloat(height) * .75).toBeCloseTo(2355 / 2048 * 20, 3);
  else expect(result.after).toEqual(result.before);
  expect(result.repeated).toEqual(result.after);
});

for (const [name, content, owned, restored] of [
  ['plain leaves', '<span>Verbena note text.</span>', true, true],
  ['superscript reference beside text', '<span style="vertical-align:super">i</span><span>Verbena note text.</span>', true, false],
  ['linked reference beside text', '<a href="#ref">i</a><span>Verbena note text.</span>', true, false],
  ['mixed note fonts', '<span style="font-size:10pt">Small</span><span>Verbena note text.</span>', true, false],
  ['unowned nested spans', '<span>Verbena note text.</span>', false, false],
] as const) test(`endnote minimum validates ${name}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  const result = await page.evaluate(async ({ content, owned }) => {
    const url = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(url);
    const root = document.createElement('section'); root.className = 'endnotes';
    const noteId = 'en:en:0123456789abcdef0123456789abcdef';
    root.innerHTML = '<ol><li><p data-source-anchor-id="p:en:abcdef0123456789abcdef0123456789" '
      + 'style="font:20pt/16pt Liberation Serif;--docx-min-line-height:16pt"><span '
      + (owned ? `data-source-anchor-id="${noteId}"` : '') + '>' + content + '</span></p></li></ol>';
    document.body.append(root);
    const elements = Array.from(root.querySelectorAll('p,span,a'));
    const before = elements.map(e => getComputedStyle(e).lineHeight);
    const identities = () => elements.map(e => e.getAttribute('data-source-anchor-id'));
    const originalIds = identities();
    applyUnroundedNormalLineHeights(root);
    return { before, after:elements.map(e => getComputedStyle(e).lineHeight), originalIds, ids:identities() };
  }, { content, owned });
  expect(result.ids).toEqual(result.originalIds);
  if (restored) for (const height of result.after) expect(Number.parseFloat(height) * .75).toBeCloseTo(2355 / 2048 * 20, 3);
  else expect(result.after).toEqual(result.before);
});
