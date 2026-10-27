import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import {
  anchorlessCommentsFixture, balloonDocx, commented, crowdedCommentsFixture, formatted, markupBalloonsFixture, plain, reformatted,
} from './docx-markup-balloons-fixture.js';
import { printOverflowFixture } from './docx-print-overflow-fixture.js';

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

/** Each PDF page's text, whitespace collapsed, and the left edge of its text, in points. */
async function pdfPages(pdf: Uint8Array): Promise<Array<{ text: string; left: number }>> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const document = await task.promise;
    const pages = await Promise.all(Array.from({ length: document.numPages }, (_, i) => document.getPage(i + 1)));
    return await Promise.all(pages.map(async (page) => {
      const { width, height } = page.getViewport({ scale: 1 });
      expect([Math.round(width), Math.round(height)]).toEqual([612, 792]);
      const items = (await page.getTextContent()).items.filter(item => 'str' in item && item.str.trim());
      return {
        text: items.map(item => 'str' in item ? item.str : '').join(' ').replace(/\s+/g, ' '),
        left: Math.min(...items.map(item => 'transform' in item ? item.transform[4] : Infinity)),
      };
    }));
  } finally { await task.destroy(); }
}

const overflowOptions = {
  outputs: ['html', 'pdf'] as Array<'html' | 'pdf'>, reviewProfile: 'final' as const, commentProfile: 'margin' as const,
  commentColumnPlacement: 'balloons' as const, timeoutMs: 120_000,
};

test('by default the rest of a cut comment is printed nowhere, as Word\'s Save as PDF does', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(printOverflowFixture(), { browser, ...overflowOptions });
  expect('balloonOverflow' in result.renderReport.options).toBe(false);
  const pages = await pdfPages(result.pdf!);
  expect(pages).toHaveLength(1);
  expect(result.pageMap.pages).toHaveLength(1);
  expect(pages[0].text).toContain('…');
  expect(pages[0].text).not.toMatch(/… \[\d+\]/);
});

test('the appendix overflow prints each cut comment in full after the last page, as Word\'s Print to PDF does', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(printOverflowFixture(), { browser, ...overflowOptions, balloonOverflow: 'appendix' });
  await writeFile(testInfo.outputPath('balloon-appendix.pdf'), result.pdf!);
  expect(result.renderReport.options.balloonOverflow).toBe('appendix');
  const warning = result.warnings.find(w => w.code === 'comment_balloon_truncated')!;
  expect(warning.message).toMatch(/in full after the last page/);
  const cut = warning.resource!.split(',');
  const pages = await pdfPages(result.pdf!);
  expect(pages.length).toBeGreaterThan(1);
  expect(result.pageMap.pages).toHaveLength(pages.length);
  // The boxes are numbered through the document, and each number heads its comment's full text.
  expect(pages[0].text).toContain('… [1]');
  expect(pages[0].text).toContain(`… [${cut.length}]`);
  const appendix = pages.slice(1).map(page => page.text).join(' ');
  expect(appendix.match(/Page 1: \[\d+\] Commented \[(IM|KA)\d+\]/g)).toHaveLength(cut.length);
  expect(appendix).toContain('Page 1: [1] Commented [');
  expect(appendix).toMatch(/Ines Moreau|Kwame Asante/);
  // The date as its author's clock read it, not the UTC instant the comment parts record.
  expect(appendix).toContain('6/11/2025 4:05:00 PM');
  expect(appendix).not.toContain('9:05:00 PM');
  for (const id of cut.filter(id => id !== '12')) {
    expect(appendix).toContain(`Note ${id}: the warranty period in this section should run from acceptance, not delivery.`);
  }
  // The shrunk page and the full-size appended pages each print at their own scale.
  expect(pages[0].left).toBeCloseTo(72 * 612 / (612 - 72 + 266.4), 0);
  for (const page of pages.slice(1)) expect(page.left).toBeCloseTo(72, 0);
  // The comment of several paragraphs keeps its blank lines.
  const paragraphs = [...result.html!.matchAll(/class="page-balloon-appendix-text"[^>]*>([^<]*)</g)].map(match => match[1]);
  const start = paragraphs.indexOf('First, the cap.');
  expect(paragraphs.slice(start, start + 5)).toEqual(['First, the cap.', '', 'Second, the basket.', '', 'Third, the survival period.']);
});

test('a long cut comment runs on across the appended pages with every word printed once, in order', async ({ browser }) => {
  test.setTimeout(120_000);
  const words = Array.from({ length: 900 }, (_, i) => `w${i + 1}`);
  const ids = Array.from({ length: 25 }, (_, id) => id);
  const body = ids.map(id => `<w:p>${commented(id, `Line ${id} `)}</w:p>`).join('');
  const filler = 'The schedule of deliverables needs a second look before the closing date. '.repeat(6);
  const comments = ids.map(id => ({ id, author: 'Nadia Okafor', initials: 'NO', text: id === 12 ? words.join(' ') : filler }));
  const result = await renderDocxArtifacts(balloonDocx({ body, comments }), { browser, ...overflowOptions, balloonOverflow: 'appendix' });
  const appendix = (await pdfPages(result.pdf!)).slice(1);
  expect(appendix.length).toBeGreaterThan(1);
  expect(appendix.map(page => page.text).join(' ').match(/\bw\d+\b/g)).toEqual(words);
  // The comment itself crosses a page break.
  expect(appendix.filter(page => /\bw\d+\b/.test(page.text)).length).toBeGreaterThan(1);
});

test('the appendix overflow without the balloon layout is an invalid argument', async ({ browser }) => {
  await expect(renderDocxArtifacts(markupBalloonsFixture(), {
    browser, outputs: ['html'], reviewProfile: 'final', commentProfile: 'margin', balloonOverflow: 'appendix', timeoutMs: 60_000,
  })).rejects.toHaveProperty('code', 'invalid_argument');
});
