import { expect, test, type Page } from '@playwright/test';
import { outOfFlowDocx } from './docx-out-of-flow-fixture.js';

/**
 * Issue #1285: tab-rendering gaps in leader lines and signature lines, laid out by Chromium.
 *
 * Every entry sits on a Letter page with one-inch margins, so the column is 624px wide and a
 * right stop at 9360 twips is its right edge. Each entry has a left stop at 1296 twips (86.4px)
 * and `w:ind w:left="1296" w:right="864" w:hanging="648"`, so its line ends at 566.4px.
 */

const COLUMN_PX = 624;
const LABEL_PX = 43.2;
const TITLE_PX = 86.4;

const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const tab = (rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:tab/></w:r>`;

function entry(content: string, stops = '<w:tab w:val="left" w:pos="1296"/><w:tab w:val="right" w:leader="dot" w:pos="9360"/>') {
  return `<w:p><w:pPr><w:tabs>${stops}</w:tabs><w:ind w:left="1296" w:right="864" w:hanging="648"/></w:pPr>${content}</w:p>`;
}

interface Word { text: string; left: number; right: number; top: number }

/** Convert, parse with the browser's HTML parser, and lay out in a 6.5in column. */
async function layOut(page: Page, bodyXml: string) {
  const bytes = Array.from(outOfFlowDocx(bodyXml));
  const result = await page.evaluate((values) =>
    (window as any).DocxodusTests.convertToHtml(new Uint8Array(values)), bytes);
  expect(result.error).toBeUndefined();
  await page.evaluate((html: string) => {
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.innerHTML = parsed.body.innerHTML;
    document.body.style.cssText = 'margin:0;padding:0;width:6.5in';
  }, result.html!);
}

/** Every word of every paragraph, in px from the column's left edge and top. */
function words(page: Page): Promise<Word[][]> {
  return page.evaluate(() => {
    const column = document.body.getBoundingClientRect();
    return Array.from(document.querySelectorAll('p')).map((p) => {
      const found: { text: string; left: number; right: number; top: number }[] = [];
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        for (const match of (node as Text).data.matchAll(/\S+/g)) {
          const range = document.createRange();
          range.setStart(node, match.index!);
          range.setEnd(node, match.index! + match[0].length);
          const box = range.getBoundingClientRect();
          found.push({ text: match[0], left: box.left - column.left, right: box.right - column.left, top: box.top - column.top });
        }
      }
      return found;
    });
  });
}

const find = (line: Word[], text: string) => line.find((w) => w.text === text)!;

test.describe('tab leader and signature lines (#1285)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined, { timeout: 30_000 });
  });

  test('an underlined or struck tab draws its line across the tab, without growing the line', async ({ page }) => {
    await layOut(page,
      `<w:p>${run('By:')}${tab('<w:u w:val="single"/>')}</w:p>` +
      `<w:p>${run('By:')}${tab('<w:strike/>')}</w:p>` +
      `<w:p>${run('By:')}${tab()}</w:p>`);
    const geometry = await page.evaluate(() => Array.from(document.querySelectorAll('p')).map((p) => {
      const box = (e: Element) => e.getBoundingClientRect();
      const tabBox = p.querySelector('[data-docx-tab]');
      return {
        height: box(p).height,
        tab: tabBox ? { left: box(tabBox).left, right: box(tabBox).right } : null,
        lines: tabBox ? Array.from(tabBox.children).map((line) => ({
          left: box(line).left, right: box(line).right, top: box(line).top - box(p).top,
          style: getComputedStyle(line).borderTopStyle,
        })) : [],
      };
    }));
    const [underlined, struck, plain] = geometry;
    for (const decorated of [underlined, struck]) {
      expect(decorated.lines).toHaveLength(1);
      expect(decorated.lines[0].style).toBe('solid');
      expect(decorated.lines[0].left).toBeCloseTo(decorated.tab!.left, 1);
      expect(decorated.lines[0].right).toBeCloseTo(decorated.tab!.right, 1);
      expect(decorated.tab!.right - decorated.tab!.left).toBeGreaterThan(20);
      expect(decorated.height).toBeCloseTo(plain.height, 1);
    }
    expect(plain.lines).toHaveLength(0);
    // Both lines lie inside the line box, the underline well below the strike.
    for (const decorated of [underlined, struck]) {
      expect(decorated.lines[0].top).toBeGreaterThan(0);
      expect(decorated.lines[0].top).toBeLessThan(decorated.height);
    }
    expect(underlined.lines[0].top).toBeGreaterThan(struck.lines[0].top + 4);
  });

  test('a bookmark inside a link keeps the whole entry inside that link and on its stops', async ({ page }) => {
    const linked = (bookmark: string) => entry(`<w:hyperlink w:anchor="_Toc9">${bookmark}${run('2.1.')}${tab()}${run('Spillway')}<w:r><w:tab/><w:t>12</w:t></w:r></w:hyperlink>`);
    await layOut(page, linked('<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkEnd w:id="1"/>') + linked(''));
    const links = await page.evaluate(() => Array.from(document.querySelectorAll('p')).map((p) => ({
      anchors: p.querySelectorAll('a').length,
      linked: p.querySelector('a[href]')?.textContent,
      target: document.getElementById('_Toc1')?.closest('a[href]') != null,
    })));
    expect(links[0]).toEqual({ anchors: 1, linked: '2.1.Spillway12', target: true });
    const [withBookmark, control] = await words(page);
    for (const line of [withBookmark, control]) {
      expect(find(line, '2.1.').left).toBeCloseTo(LABEL_PX, 0);
      expect(find(line, 'Spillway').left).toBeCloseTo(TITLE_PX, 0);
      expect(find(line, '12').right).toBeCloseTo(COLUMN_PX, 0);
      expect(new Set(line.map((w) => w.top)).size).toBe(1);
    }
  });
});
