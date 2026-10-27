import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import {
  anchorlessCommentsFixture, balloonDocx, commented, crowdedCommentsFixture, formatted, markupBalloonsFixture, plain,
  reformatted,
} from './docx-markup-balloons-fixture.js';

async function pdfText(pdf: Uint8Array): Promise<string> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const document = await task.promise;
    const pages = await Promise.all(Array.from({ length: document.numPages }, (_, i) => document.getPage(i + 1)));
    const contents = await Promise.all(pages.map(page => page.getTextContent()));
    return contents.flatMap(content => content.items.filter(item => 'str' in item).map(item => item.str)).join(' ').replace(/\s+/g, ' ');
  } finally { await task.destroy(); }
}

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('the PDF prints Word\'s All Markup balloons on the source paper size', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture(), {
    browser, outputs: ['pdf'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000,
  });
  await writeFile(testInfo.outputPath('markup-balloons.pdf'), result.pdf!);
  expect(result.renderReport.options.commentColumnPlacement).toBe('balloons');
  expect(result.pageMap.pages.map(p => [p.width, p.height])).toEqual([[612, 792]]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const page = await pdf.getPage(1), viewport = page.getViewport({ scale: 1 });
    expect(viewport.width).toBeCloseTo(612, 0); expect(viewport.height).toBeCloseTo(792, 0);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ')
      .replace(/\s+/g, ' ');
    for (const expected of ['Formatted: Font: Bold', 'Formatted: Indent: Left: 0.5"', 'Commented [AA1]:',
      'Commented [BB2R1]:', 'The buyer pays on delivery.']) {
      expect(text).toContain(expected);
    }
  } finally { await task.destroy(); }
});

test('the final profile prints comment balloons and no formatting balloons', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture('a4'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', commentColumnPlacement: 'balloons',
    timeoutMs: 120_000,
  });
  expect(result.pageMap.pages.map(p => [Math.round(p.width), Math.round(p.height)])).toEqual([[595, 842]]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const page = await (await task.promise).getPage(1);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ')
      .replace(/\s+/g, ' ');
    expect(text).toContain('Commented [AA1]:');
    expect(text).not.toContain('Formatted');
  } finally { await task.destroy(); }
});

test('point, empty-range and ranged comments all print, with no truncation warning', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(anchorlessCommentsFixture(), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', commentColumnPlacement: 'balloons',
    timeoutMs: 120_000,
  });
  const text = await pdfText(result.pdf!);
  for (const expected of ['Commented [AA1]: PointNote', 'Commented [BB2]: EmptyNote', 'Commented [AA3]: RangeNote']) {
    expect(text).toContain(expected);
  }
  expect(result.warnings.map(w => w.code)).not.toContain('comment_balloon_truncated');
});

test('formatting changes in a header print their balloons on every page', async ({ browser }) => {
  test.setTimeout(120_000);
  const header = `<w:p>${formatted(21, 'Header run')}</w:p>${reformatted(22, 'Header paragraph', '<w:jc w:val="center"/>')}`;
  const body = Array.from({ length: 60 }, (_, i) => plain(`Body line ${i + 1}.`)).join('');
  const result = await renderDocxArtifacts(balloonDocx({ body, header }), {
    browser, outputs: ['pdf'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'hidden',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000,
  });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(2);
    for (let number = 1; number <= 2; number++) {
      const content = await (await pdf.getPage(number)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ').replace(/\s+/g, ' ');
      expect(text).toContain('Formatted: Font: Bold');
      expect(text).toContain('Formatted: Centered');
    }
  } finally { await task.destroy(); }
});

test('balloons cut to fit the page are reported', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(crowdedCommentsFixture(25, 1, 'a long remark that wraps over several balloon lines. '.repeat(6)), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', commentColumnPlacement: 'balloons',
    timeoutMs: 120_000,
  });
  const warning = result.warnings.find(w => w.code === 'comment_balloon_truncated');
  expect(warning).toBeDefined();
  expect(warning!.resource).toMatch(/^\d+(,\d+)*$/);
});

test('the default export names no comment column placement', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture(), {
    browser, outputs: ['html'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    timeoutMs: 120_000,
  });
  expect('commentColumnPlacement' in result.renderReport.options).toBe(false);
  expect(result.html).not.toContain('page-balloon');
});

for (const [name, options] of [
  ['inline comments', { reviewProfile: 'final', commentProfile: 'inline' }],
  ['the Docxodus presentation', { reviewProfile: 'markup', commentProfile: 'margin' }],
] as const) {
  test(`balloons with ${name} are an invalid argument`, async ({ browser }) => {
    await expect(renderDocxArtifacts(markupBalloonsFixture(), {
      browser, outputs: ['html'], ...options, commentColumnPlacement: 'balloons', timeoutMs: 60_000,
    })).rejects.toHaveProperty('code', 'invalid_argument');
  });
}

test('balloons cut to fit keep the first whole, cut the rest to one line, and report only the cut comments', async ({ browser }) => {
  test.setTimeout(120_000);
  const remark = 'Please confirm the delivery window with the carrier before this clause is finalised. '.repeat(5);
  const result = await renderDocxArtifacts(crowdedCommentsFixture(25, 1, remark), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', commentColumnPlacement: 'balloons',
    timeoutMs: 120_000,
  });
  const warning = result.warnings.find(w => w.code === 'comment_balloon_truncated');
  expect(warning).toBeDefined();
  const cut = warning!.resource!.split(',');
  expect(cut).not.toContain('0');
  expect(cut.length).toBeLessThan(25);
  expect(result.pageMap.pages.map(p => [p.width, p.height])).toEqual([[612, 792]]);
  const text = await pdfText(result.pdf!);
  // The first comment prints whole; a cut comment prints its first line and the boxed ellipsis, nothing more.
  expect(text).toContain(`0: ${remark.trim()}`);
  expect(text).toContain(`${cut[0]}: Please confirm the delivery`);
  expect(text).not.toContain(`${cut[0]}: ${remark.trim()}`);
  expect(text.match(/…/g)).toHaveLength(cut.length);
});

test('pages after a page of cut balloons print at the markup layout\'s own scale', async ({ browser }) => {
  test.setTimeout(120_000);
  const remark = 'Please confirm the delivery window with the carrier before this clause is finalised. '.repeat(5);
  const ids = Array.from({ length: 25 }, (_, id) => id);
  const body = ids.map(id => `<w:p>${commented(id, `w${id} `)}</w:p>`).join('') +
    Array.from({ length: 90 }, (_, i) => plain(`Later line ${i + 1}.`)).join('');
  const comments = ids.map(id => ({ id, author: 'Tomas Lindqvist', initials: 'TL', text: `${id}: ${remark}` }));
  const result = await renderDocxArtifacts(balloonDocx({ body, comments }), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', commentColumnPlacement: 'balloons',
    timeoutMs: 120_000,
  });
  expect(result.warnings.map(w => w.code)).toContain('comment_balloon_truncated');
  // Every page starts its text column at the 1in margin shrunk by the markup factor. Chromium shrinks
  // the whole print instead when anything inside a page box escapes it.
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBeGreaterThan(1);
    for (let number = 1; number <= pdf.numPages; number++) {
      const content = await (await pdf.getPage(number)).getTextContent();
      const lefts = content.items.flatMap(item => 'transform' in item && /^(w\d|Later)/.test(item.str) ? [item.transform[4]] : []);
      expect(lefts.length).toBeGreaterThan(0);
      expect(Math.min(...lefts)).toBeCloseTo(72 * 612 / (612 - 72 + 266.4), 0);
    }
  } finally { await task.destroy(); }
});

test('the endnotes comment profile prints formatting balloons and lists the comments after the text', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture(), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'endnotes',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000,
  });
  const text = await pdfText(result.pdf!);
  expect(text).toContain('Formatted: Font: Bold');
  expect(text).not.toContain('Commented [');
  expect(text).toContain('Is delivery the right trigger?');
  // The page container carries the balloon scope without Word's comment rules, so the marks print.
  expect(result.html).toMatch(/<[^>]* data-docx-markup-balloons=""/);
  expect(result.html).not.toMatch(/<[^>]* data-docx-markup-balloons="comments"/);
});
