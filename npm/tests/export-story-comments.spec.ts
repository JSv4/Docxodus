import { test, expect, type Browser, type Page } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { STORY_COMMENTS, storyCommentsFixture } from './docx-story-comments-fixture.js';

// Issue #1310: a comment anchored in a header, footer, footnote or endnote renders like a body
// comment in every comment profile and placement, and the export completes without a warning,
// even under the strict unsupported-content policy, because no story is left unsupported.

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

/** Text of each PDF page, whitespace collapsed. */
async function pdfPages(pdf: Uint8Array): Promise<string[]> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const document = await task.promise;
    const pages = await Promise.all(Array.from({ length: document.numPages }, (_, i) => document.getPage(i + 1)));
    const contents = await Promise.all(pages.map(page => page.getTextContent()));
    return contents.map(content => content.items.filter(item => 'str' in item).map(item => item.str).join(' ')
      .replace(/\s+/g, ' '));
  } finally { await task.destroy(); }
}

type Profile = 'endnotes' | 'inline' | 'margin';

async function exportStories(browser: Browser, commentProfile: Profile, balloons = false, bodyOnly = false) {
  const result = await renderDocxArtifacts(storyCommentsFixture(bodyOnly), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile,
    ...(balloons ? { commentColumnPlacement: 'balloons' as const } : {}),
    unsupportedContent: 'strict', timeoutMs: 120_000,
  });
  // Only the host's font environment may warn; nothing about the comments or their stories.
  expect(result.renderReport.warnings.filter(w => !w.code.startsWith('font_'))).toEqual([]);
  expect(result.pageCount).toBe(2);
  return result;
}

/** The comment texts each page's margin or balloon column holds, in the delivered HTML. */
async function columnTexts(browser: Browser, html: string, selector: string): Promise<string[][]> {
  const page = await browser.newPage();
  try {
    await page.setContent(html);
    return await page.evaluate((selector) => Array.from(document.querySelectorAll('.page-box')).map(box =>
      Array.from(box.querySelectorAll(selector)).map(column => column.textContent ?? '').join(' ')
        .match(/[A-Z]+COMMENT/g) ?? []), selector);
  } finally { await page.close(); }
}

test('the endnotes profile lists every story\'s comment once, body first', async ({ browser }, testInfo) => {
  test.setTimeout(150_000);
  const result = await exportStories(browser, 'endnotes');
  await writeFile(testInfo.outputPath('story-comments-endnotes.pdf'), result.pdf!);
  const text = (await pdfPages(result.pdf!)).join(' ');
  const listed = text.match(/[A-Z]+COMMENT/g);
  expect(listed).toEqual(STORY_COMMENTS.map(({ comment }) => comment));
  for (const { text: storyText } of STORY_COMMENTS) expect(text).toContain(storyText);
});

test('the inline profile carries every story\'s comment on its highlight', async ({ browser }) => {
  test.setTimeout(150_000);
  const result = await exportStories(browser, 'inline');
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    for (const { id, text, comment } of STORY_COMMENTS) {
      const highlight = page.locator(`.page-box span.comment-highlight[data-comment-id="${id}"]`, { hasText: text });
      await expect(highlight.first()).toHaveAttribute('data-comment', comment);
    }
  } finally { await page.close(); }
});

test('margin notes for a header or footer comment print on every page that story prints on', async ({ browser }, testInfo) => {
  test.setTimeout(150_000);
  const result = await exportStories(browser, 'margin');
  await writeFile(testInfo.outputPath('story-comments-margin.pdf'), result.pdf!);
  expect(await columnTexts(browser, result.html!, '.page-comment-margin')).toEqual([
    // Top to bottom: the header, the body, the footnote, the footer.
    ['HEADERCOMMENT', 'BODYCOMMENT', 'NOTECOMMENT', 'FOOTERCOMMENT'],
    ['HEADERCOMMENT', 'ENDCOMMENT', 'FOOTERCOMMENT'],
  ]);
  // The narrow margin column wraps a long token in the PDF text layer.
  const pages = (await pdfPages(result.pdf!)).map(page => page.replace(/\s/g, ''));
  for (const comment of ['HEADERCOMMENT', 'FOOTERCOMMENT']) for (const page of pages) expect(page).toContain(comment);
  expect(pages[0]).toContain('NOTECOMMENT');
  expect(pages[1]).toContain('ENDCOMMENT');
});

test('balloons for a header or footer comment print on every page', async ({ browser }, testInfo) => {
  test.setTimeout(150_000);
  const result = await exportStories(browser, 'margin', true);
  await writeFile(testInfo.outputPath('story-comments-balloons.pdf'), result.pdf!);
  const pages = await pdfPages(result.pdf!);
  for (const comment of ['HEADERCOMMENT', 'FOOTERCOMMENT']) for (const page of pages) expect(page).toContain(comment);
  expect(pages[0]).toContain('BODYCOMMENT');
  expect(pages[0]).toContain('NOTECOMMENT');
  expect(pages[1]).toContain('ENDCOMMENT');
});

test('control: a body-only comment prints alone in every profile', async ({ browser }) => {
  test.setTimeout(150_000);
  const endnotes = await exportStories(browser, 'endnotes', false, true);
  expect((await pdfPages(endnotes.pdf!)).join(' ').match(/[A-Z]+COMMENT/g)).toEqual(['BODYCOMMENT']);
  // The margin column is narrow enough to wrap the token in the PDF, so read the delivered HTML.
  const margin = await exportStories(browser, 'margin', false, true);
  expect(await columnTexts(browser, margin.html!, '.page-comment-margin')).toEqual([['BODYCOMMENT'], []]);
  const balloons = await exportStories(browser, 'margin', true, true);
  const pages = await pdfPages(balloons.pdf!);
  expect(pages[0].match(/[A-Z]+COMMENT/g)).toEqual(['BODYCOMMENT']);
  expect(pages[1]).not.toContain('COMMENT');
});

/** The paginated view, without the export: each page's comment column, for one placement. */
async function viewColumns(page: Page, placement: 'margin' | 'outside'): Promise<string[][]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  return page.evaluate(({ bytes, placement }) => {
    // Margin comments (2), paginated (1), notes and running stories rendered (the two trues after
    // the annotation prefix), tracked changes off.
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Story comments', 'docx-', false, '', 2, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', true, true, false, true, true, false, null, false, 0);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.insertAdjacentHTML('beforeend', parsed.head.innerHTML);
    document.body.innerHTML = '<div id="host"></div>';
    const host = document.querySelector<HTMLElement>('#host')!;
    (window as any).DocxodusPagination.paginateHtml(parsed.body.innerHTML, host, {
      showPageNumbers: false, commentColumnPlacement: placement,
    });
    const pages = Array.from(host.querySelectorAll<HTMLElement>('.page-box')).map(box => box.dataset.pageNumber);
    const columns = Array.from(host.querySelectorAll<HTMLElement>('.page-comment-margin'));
    return pages.map(number => columns.filter(column => column.dataset.pageNumber === number)
      .flatMap(column => (column.textContent ?? '').match(/[A-Z]+COMMENT/g) ?? []));
  }, { bytes: Array.from(storyCommentsFixture()), placement });
}

for (const placement of ['margin', 'outside'] as const) {
  test(`the paginated view places running-story notes on every page (${placement})`, async ({ page }) => {
    test.setTimeout(120_000);
    expect(await viewColumns(page, placement)).toEqual([
      ['HEADERCOMMENT', 'BODYCOMMENT', 'NOTECOMMENT', 'FOOTERCOMMENT'],
      ['HEADERCOMMENT', 'ENDCOMMENT', 'FOOTERCOMMENT'],
    ]);
  });
}
