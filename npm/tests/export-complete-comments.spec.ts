import { test, expect } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { longCommentFixture, LONG_COMMENT_TOKEN } from './docx-long-comment-fixture.js';
import { commentThreadFixture } from './docx-comment-thread-fixture.js';
import { storyCommentsFixture, STORY_COMMENTS } from './docx-story-comments-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('the appendix pointer fits a short page band', async ({ browser }) => {
  const result = await renderDocxArtifacts(longCommentFixture(3, false, false, false, 3360), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin',
  });
  expect(await printedText(result.pdf!)).toContain('CedarNote002');
});

async function printedText(bytes: Uint8Array): Promise<string> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    let text = '';
    for (let i = 1; i <= pdf.numPages; i++)
      text += (await (await pdf.getPage(i)).getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join('');
    return text.replace(/\s+/g, '');
  } finally { await task.destroy(); }
}

for (const hideHeader of [false, true]) test(`comment owner covers every continuation with metadata hidden=${hideHeader}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(longCommentFixture(900), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
    additionalCss: hideHeader ? 'div.comment-margin-note-header { display: none }' : '',
  });
  const paragraphPages = new Set(result.pageMap.fragments.filter(f => f.anchorId.startsWith('p:cmt:')).map(f => f.pageNumber));
  const ownerPages = new Set(result.pageMap.fragments.filter(f => f.anchorId.startsWith('cmt:')).map(f => f.pageNumber));
  expect(paragraphPages.size).toBeGreaterThan(1);
  for (const page of paragraphPages) expect(ownerPages, `comment owner missing on page ${page}`).toContain(page);
  expect(await printedText(result.pdf!)).toContain('CedarNote899');
});

for (const multipleParagraphs of [false, true]) {
  test(`margin export prints every word of a long ${multipleParagraphs ? 'multi-paragraph' : 'single-paragraph'} comment`, async ({ browser }) => {
    test.setTimeout(120_000);
    const count = 240;
    const result = await renderDocxArtifacts(longCommentFixture(count, true, true, multipleParagraphs), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      let text = '';
      for (let number = 1; number <= pdf.numPages; number++) {
        const page = await pdf.getPage(number), viewport = page.getViewport({ scale: 1 });
        expect(viewport.width).toBeCloseTo(612, 0);
        expect(viewport.height).toBeCloseTo(792, 0);
        text += (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join('');
      }
      const compact = text.replace(/\s+/g, '');
      expect(compact).toContain('Cedarinventorynoteanchor.');
      expect(compact).toContain('Cedarsecondpagecontrol.');
      for (let i = 0; i < count; i++) expect(compact, `missing comment word ${i}`).toContain(`CedarNote${String(i).padStart(3, '0')}`);
      expect(compact).toContain(LONG_COMMENT_TOKEN);
      expect(pdf.numPages).toBe(result.pageCount);
    } finally { await task.destroy(); }
  });
}

for (const threaded of [false, true]) {
  test(`overflow prints all ${threaded ? 'nested replies' : 'short independent comments'} and keeps return links`, async ({ browser }) => {
    test.setTimeout(120_000);
    const count = threaded ? 3 : 60, words = threaded ? 120 : 1;
    const result = await renderDocxArtifacts(commentThreadFixture(threaded, false, false, { count, words }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
    });
    const text = await printedText(result.pdf!);
    for (let i = 0; i < count; i++) for (let j = 0; j < words; j++) expect(text).toContain(`CedarThread${i}Word${j}`);
    const page = await browser.newPage();
    try {
      await page.setContent(result.html!);
      await expect(page.locator('[data-comment-appendix]')).not.toHaveCount(0);
      const links = await page.locator('[data-comment-appendix] a.comment-margin-backref').evaluateAll(nodes =>
        nodes.map(node => ({ href: node.getAttribute('href'), disabled: node.getAttribute('aria-disabled') })));
      expect(links).toHaveLength(count);
      for (let i = 0; i < count; i++) expect(links[i]).toEqual({ href: `#comment-ref-${i}`, disabled: null });
    } finally { await page.close(); }
  });
}

test('overflow retains comments selected by body, header, footer, footnote and endnote markers', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(storyCommentsFixture(false, 120), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
  });
  const text = await printedText(result.pdf!);
  for (const { comment, text: anchor } of STORY_COMMENTS) {
    expect(text).toContain(anchor);
    for (let i = 0; i < 120; i++) expect(text).toContain(`${comment}${i}`);
  }
});

test('comment appendix pages are charged against the export page limit', async ({ browser }) => {
  await expect(renderDocxArtifacts(longCommentFixture(240, false, false, true), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', limits: { finalPages: 1 },
  })).rejects.toMatchObject({ code: 'resource_limit' });
});

for (const mode of ['hidden', 'balloons'] as const) test(`${mode} retains its existing long-comment policy`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(longCommentFixture(240), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: mode === 'hidden' ? 'hidden' : 'margin',
    ...(mode === 'balloons' ? { commentColumnPlacement: 'balloons' as const, balloonOverflow: 'appendix' as const } : {}),
    timeoutMs: 120_000,
  });
  const text = await printedText(result.pdf!);
  expect(text).toContain('Cedarinventorynoteanchor.');
  if (mode === 'hidden') {
    expect(result.pageCount).toBe(1);
    expect(text).not.toContain('CedarNote');
  } else {
    for (let i = 0; i < 240; i++) expect(text).toContain(`CedarNote${String(i).padStart(3, '0')}`);
  }
  expect(result.html).not.toContain('data-comment-appendix');
});
