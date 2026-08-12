import { expect, test, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { generateTabDocx, type TabAlignment } from './docx-tab-fixture.js';

const TAB_TARGET_PX = 4 * 96;
const TEST_FILES = join(dirname(fileURLToPath(import.meta.url)), '../../TestFiles');

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined, {
    timeout: 30_000,
  });
}

async function renderGeneratedTab(
  page: Page,
  alignment: TabAlignment,
  before: string,
  after: string,
  leader?: 'dot',
) {
  const bytes = generateTabDocx(alignment, before, after, leader);
  const result = await page.evaluate((values) =>
    (window as any).DocxodusTests.convertToHtml(new Uint8Array(values)), Array.from(bytes));
  expect(result.error).toBeUndefined();
  expect(result.html).toBeDefined();

  await page.evaluate((html: string) => {
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.innerHTML = parsed.body.innerHTML;
  }, result.html!);

  return page.evaluate(({ expectedAlignment, followingText }) => {
    const tab = document.querySelector<HTMLElement>(`[data-docx-tab="${expectedAlignment}"]`)!;
    const paragraph = tab.closest('p')!;
    const following = Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
      .find((span) => span.children.length === 0 && span.textContent === followingText)!;
    const paragraphBox = paragraph.getBoundingClientRect();
    const followingBox = following.getBoundingClientRect();
    const tabBox = tab.getBoundingClientRect();
    const decimalOffset = followingText.indexOf('.');
    let decimal = Number.NaN;
    if (decimalOffset >= 0) {
      const textNode = following.firstChild!;
      const range = document.createRange();
      range.setStart(textNode, decimalOffset);
      range.collapse(true);
      decimal = range.getBoundingClientRect().left - paragraphBox.left;
    }
    return {
      followingLeft: followingBox.left - paragraphBox.left,
      followingRight: followingBox.right - paragraphBox.left,
      followingCenter: (followingBox.left + followingBox.right) / 2 - paragraphBox.left,
      followingWidth: followingBox.width,
      decimal,
      tabWidth: Number(tab.dataset.docxTabWidth) * 96,
      leaderWidth: tabBox.width,
      leaderText: tab.textContent ?? '',
      borderBottomStyle: getComputedStyle(tab).borderBottomStyle,
    };
  }, { expectedAlignment: alignment, followingText: after });
}

test.describe('generated tab-stop geometry', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('right tabs retain their target across current and following-run widths', async ({ page }) => {
    const narrowCurrent = await renderGeneratedTab(page, 'right', 'iiii', '7');
    const wideCurrent = await renderGeneratedTab(page, 'right', 'iiiiiiii', '7');
    const wideFollowing = await renderGeneratedTab(page, 'right', 'iiii', '12345');

    expect(narrowCurrent.followingRight).toBeCloseTo(TAB_TARGET_PX, 0);
    expect(wideCurrent.followingRight).toBeCloseTo(TAB_TARGET_PX, 0);
    expect(wideFollowing.followingRight).toBeCloseTo(TAB_TARGET_PX, 0);
    // Four extra 'i' characters at the estimator's 0.25 em/char and 12 pt: 4 × 4px.
    expect(narrowCurrent.tabWidth - wideCurrent.tabWidth).toBeCloseTo(16, 0);
    expect(narrowCurrent.followingLeft - wideFollowing.followingLeft)
      .toBeCloseTo(wideFollowing.followingWidth - narrowCurrent.followingWidth, 0);
  });

  test('center tab pins the following run midpoint', async ({ page }) => {
    const geometry = await renderGeneratedTab(page, 'center', 'iiii', 'CENTER');
    expect(geometry.followingCenter).toBeCloseTo(TAB_TARGET_PX, 0);
  });

  test('decimal tab pins the decimal separator', async ({ page }) => {
    const geometry = await renderGeneratedTab(page, 'decimal', 'iiii', '123.45');
    expect(geometry.decimal).toBeCloseTo(TAB_TARGET_PX, 0);
  });

  test('right dot leader fills its advance without entering the following run', async ({ page }) => {
    const geometry = await renderGeneratedTab(page, 'right', 'iiii', '7', 'dot');
    expect(geometry.followingRight).toBeCloseTo(TAB_TARGET_PX, 0);
    expect(geometry.leaderWidth).toBeCloseTo(geometry.tabWidth, 0);
    expect(geometry.leaderText).toBe('');
    expect(geometry.borderBottomStyle).toBe('dotted');
  });

  test('paginated cached TOC content retains its right-tab target', async ({ page }) => {
    const bytes = new Uint8Array(readFileSync(join(TEST_FILES, 'HC022-Table-Of-Contents.docx')));
    const result = await page.evaluate((values) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(values),
        'Document', 'docx-', true, '', -1, 'comment-',
        1, 1, 'page-',
        false, 0, 'annot-',
        true, true,
        false, false, false,
      );
      return html.startsWith('{') && html.includes('"Error"') ? { error: html } : { html };
    }, Array.from(bytes));
    expect(result.error).toBeUndefined();

    const geometry = await page.evaluate((html: string) => {
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = parsed.body.innerHTML;
      const tab = document.querySelector<HTMLElement>('[data-docx-tab="right"]')!;
      const paragraph = tab.closest('p')!;
      const pageNumber = Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
        .find((span) => span.children.length === 0 && span.textContent === '1')!;
      const paragraphBox = paragraph.getBoundingClientRect();
      const pageNumberBox = pageNumber.getBoundingClientRect();
      return {
        tabWidth: tab.getBoundingClientRect().width,
        tabDataWidth: Number(tab.dataset.docxTabWidth) * 96,
        pageNumberRight: pageNumberBox.right - paragraphBox.left,
      };
    }, result.html!);

    const tocTargetPx = 9350 / 1440 * 96;
    expect(geometry.tabWidth).toBeGreaterThan(200);
    expect(geometry.tabWidth).toBeGreaterThan(geometry.tabDataWidth);
    expect(Math.abs(geometry.pageNumberRight - tocTargetPx)).toBeLessThanOrEqual(2);
  });
});
