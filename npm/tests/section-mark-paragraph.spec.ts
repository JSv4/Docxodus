import { test, expect, Page } from '@playwright/test';
import {
  sectionMarkDocx, line, mark, table, SPACED, PAGE_BREAK, NEXT_PAGE,
} from './docx-section-mark-fixture.js';

// Word sets an empty paragraph that only carries a section break on the preceding paragraph's last
// line (#1572): it takes no line and no spacing, and after a page break it adds no page. Word's
// values are in docs/ooxml_corner_cases.md; each case here is checked against a no-mark control.

interface Layout { pages: number; lines: { text: string; page: number; top: number }[] }

/** Converts and paginates one document; the engine boots once per test page. */
async function layout(page: Page, bytes: Uint8Array): Promise<Layout> {
  if (!await page.evaluate(() => (window as any).DocxodusPagination !== undefined)) {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
  }
  return page.evaluate((bytes) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Section marks', 'docx-', false, '', -1, 'comment-', 1, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
    if (html.startsWith('{')) throw new Error(html);
    const host = document.createElement('div'); document.body.replaceChildren(host);
    (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
    const pages = Array.from(host.querySelectorAll<HTMLElement>('.page-content'));
    const lines = pages.flatMap((content, index) => {
      const origin = content.getBoundingClientRect().top;
      const walker = document.createTreeWalker(content, NodeFilter.SHOW_TEXT);
      const found: { text: string; page: number; top: number }[] = [];
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = node.textContent?.trim() ?? '';
        if (!/[A-Za-z]/.test(text)) continue;
        const range = document.createRange(); range.selectNodeContents(node);
        found.push({ text, page: index + 1, top: (range.getBoundingClientRect().top - origin) * 0.75 });
      }
      return found;
    });
    return { pages: host.querySelectorAll('.page-box').length, lines };
  }, Array.from(bytes));
}

const top = (result: Layout, text: string) => {
  const found = result.lines.find(l => l.text.startsWith(text));
  if (!found) throw new Error(`Missing ${text} in ${JSON.stringify(result.lines)}`);
  return found;
};
const step = (result: Layout, from: string, to: string) => top(result, to).top - top(result, from).top;

const A = line('Line above.');
const B = line('Line below.');
const BEFORE = '<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>';
const BIG = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="56"/></w:rPr>';
const BOOKMARK = '<w:bookmarkStart w:id="0" w:name="here"/><w:bookmarkEnd w:id="0"/>';
const CONTINUOUS_TWO_COLUMNS = NEXT_PAGE.replace('<w:pgSz', '<w:type w:val="continuous"/><w:pgSz')
  .replace('</w:sectPr>', '<w:cols w:num="2" w:space="720"/></w:sectPr>');

test('an empty section mark adds no line, spacing or run height before a continuous section', async ({ page }) => {
  const none = step(await layout(page, sectionMarkDocx(A + B)), 'Line above', 'Line below');
  expect(none).toBeGreaterThan(10);
  for (const [name, body, final] of [
    ['plain', A + mark() + B, undefined],
    ['12 pt before', A + mark(BEFORE) + B, undefined],
    ['28 pt mark', A + mark(BIG) + B, undefined],
    ['bookmarked', A + mark().replace('</w:pPr>', `</w:pPr>${BOOKMARK}`) + B, undefined],
    ['two-column next section', A + mark() + B, CONTINUOUS_TWO_COLUMNS],
  ] as const) {
    const result = await layout(page, sectionMarkDocx(body, final));
    expect(Math.abs(step(result, 'Line above', 'Line below') - none), name).toBeLessThan(0.3);
    expect(result.pages, name).toBe(1);
  }
});

test('an empty section mark drops its own space after but keeps the preceding paragraph\'s', async ({ page }) => {
  const spaced = (text: string) => line(text, SPACED);
  const none = step(await layout(page, sectionMarkDocx(spaced('Line above.') + spaced('Line below.'))), 'Line above', 'Line below');
  const marked = await layout(page, sectionMarkDocx(spaced('Line above.') + mark(SPACED) + spaced('Line below.')));
  expect(Math.abs(step(marked, 'Line above', 'Line below') - none)).toBeLessThan(0.3);
});

test('a mark with no paragraph line before it keeps its own line', async ({ page }) => {
  const control = await layout(page, sectionMarkDocx(A + B));
  const none = step(control, 'Line above', 'Line below');
  const secondLine = top(control, 'Line below').top;
  const plain = await layout(page, sectionMarkDocx(A + line('') + B));
  expect(Math.abs(step(plain, 'Line above', 'Line below') - 2 * none)).toBeLessThan(0.3);
  const text = await layout(page, sectionMarkDocx(A + mark(undefined, 'Text carrying the break.') + B));
  expect(Math.abs(step(text, 'Text carrying', 'Line below') - none)).toBeLessThan(0.3);
  const afterTable = await layout(page, sectionMarkDocx(A + table('Cell text.') + mark() + B));
  expect(step(afterTable, 'Cell text', 'Line below')).toBeGreaterThan(1.8 * none);
  const first = await layout(page, sectionMarkDocx(mark() + B));
  expect(Math.abs(top(first, 'Line below').top - secondLine)).toBeLessThan(0.3);
  // The second mark opens a new-page section, so it is the first line on page 2.
  const second = await layout(page, sectionMarkDocx(A + mark() + mark() + B));
  expect(second.pages).toBe(2);
  expect(top(second, 'Line below').page).toBe(2);
  expect(Math.abs(top(second, 'Line below').top - secondLine)).toBeLessThan(0.3);
});

const part = (n: number) => line(`Section ${n} opening line.`) + line(`Section ${n} second line.`);

test('a page break before an empty next-page section mark adds no page', async ({ page }) => {
  for (const compat of [15, 12]) {
    const result = await layout(page, sectionMarkDocx(
      part(1) + PAGE_BREAK + mark() + part(2) + PAGE_BREAK + mark() + part(3), NEXT_PAGE, compat));
    expect(result.pages, `mode ${compat}`).toBe(3);
    expect([1, 2, 3].map(n => top(result, `Section ${n} opening`).page)).toEqual([1, 2, 3]);
  }
  const markOnly = await layout(page, sectionMarkDocx(part(1) + mark() + part(2) + mark() + part(3), NEXT_PAGE));
  expect(markOnly.pages).toBe(3);
  const textMark = mark(undefined, 'Closing line.');
  const text = await layout(page, sectionMarkDocx(
    part(1) + PAGE_BREAK + textMark + part(2) + PAGE_BREAK + textMark + part(3), NEXT_PAGE));
  expect(text.pages).toBe(5);
});

test('a page break still applies when the empty mark opens a continuous section', async ({ page }) => {
  const firstLine = top(await layout(page, sectionMarkDocx(A + B)), 'Line above').top;
  const result = await layout(page, sectionMarkDocx(A + PAGE_BREAK + mark() + B));
  expect(result.pages).toBe(2);
  expect(top(result, 'Line below').page).toBe(2);
  expect(Math.abs(top(result, 'Line below').top - firstLine)).toBeLessThan(0.3);
});
