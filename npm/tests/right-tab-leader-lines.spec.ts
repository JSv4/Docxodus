import { test, expect, type Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Leader-and-page-number lines: a label, a left tab, a title, a right dot-leader tab at the
// column's right edge (9360 twips = 624px), and a page number. Each variant adds one inert item
// that used to push the right tab back to a fixed server-measured width (issue #1275), and the
// right indent ends the line 864 twips before the stop, which used to break every entry.
const entries = [
  ['2.1.', 'Quarry Access Roads', '3'],
  ['2.2.', 'Spillway', '12'],
  ['2.3.', 'Culvert Inspection and Annual Load Ratings', '27'],
  ['2.4.', 'Riprap', '140'],
] as const;
const variants = ['plain', 'bookmark', 'hyphen', 'underline', 'strike'] as const;

function leaderLineFixture(variant: typeof variants[number], rightIndent: boolean, stop = 9360): Uint8Array {
  const tabRPr = variant === 'underline' ? '<w:rPr><w:u w:val="single"/></w:rPr>'
    : variant === 'strike' ? '<w:rPr><w:strike/></w:rPr>' : '';
  const body = entries.map(([label, title, number], i) =>
    `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1296"/><w:tab w:val="right" w:leader="dot" w:pos="${stop}"/></w:tabs>` +
    `<w:ind w:left="1296"${rightIndent ? ' w:right="864"' : ''} w:hanging="648"/></w:pPr>` +
    (variant === 'bookmark' ? `<w:bookmarkStart w:id="${i}" w:name="_Toc${i}"/><w:bookmarkEnd w:id="${i}"/>` : '') +
    `<w:r><w:t>${label}</w:t></w:r><w:r>${tabRPr}<w:tab/></w:r><w:r><w:t>${title}</w:t></w:r>` +
    (variant === 'hyphen' ? '<w:r><w:noBreakHyphen/><w:t>West</w:t></w:r>' : '') +
    `<w:r>${tabRPr}<w:tab/><w:t>${number}</w:t></w:r></w:p>`).join('');
  return leaderDocx(body);
}

/** A Letter page with one-inch margins, in 12pt Liberation Serif, holding `body`. */
function leaderDocx(body: string): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

async function layOut(page: Page, variant: typeof variants[number], rightIndent: boolean, stop: number) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  const results = await page.evaluate(({ bytes, entries }) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Leader lines', 'docx-', false, '', -1, 'comment-',
      0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.style.cssText = 'margin:0;padding:0';
    document.body.innerHTML = `<div id="host" style="width:6.5in;margin:0;padding:0">${parsed.body.innerHTML}</div>`;
    const host = document.querySelector<HTMLElement>('#host')!;
    const origin = host.getBoundingClientRect().left;
    return entries.map(([label, , number]) => {
      const paragraph = Array.from(host.querySelectorAll('p')).find(p => p.textContent?.includes(label))!;
      const leaf = (text: string) => Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
        .find(e => e.childElementCount === 0 && e.textContent === text)!.getBoundingClientRect();
      const labelRect = leaf(label), numberRect = leaf(number);
      return { labelTop: labelRect.top, numberTop: numberRect.top, numberRight: numberRect.right - origin };
    });
  }, { bytes: Array.from(leaderLineFixture(variant, rightIndent, stop)), entries });
  for (const entry of results) {
    expect(entry.numberTop).toBeCloseTo(entry.labelTop, 0);
    expect(entry.numberRight).toBeCloseTo(stop / 15, 0);
  }
}

for (const variant of variants) for (const rightIndent of [true, false]) {
  test(`leader lines keep one line and end the page number on the right stop (${variant}, right indent ${rightIndent})`, async ({ page }) => {
    await layOut(page, variant, rightIndent, 9360);
  });
}

// A stop past the column edge (10000 twips = 666.67px in a 624px column): LibreOffice 25.8 keeps
// each entry on one line and ends the number on the stop, in the page margin.
for (const rightIndent of [true, false]) {
  test(`a right stop past the column edge keeps one line with the number on the stop (right indent ${rightIndent})`, async ({ page }) => {
    await layOut(page, 'plain', rightIndent, 10000);
  });
}

// Issue #1285: more tab gaps in leader and signature lines. Each entry below has a left stop at
// 1296 twips (86.4px) and w:ind w:left="1296" w:right="864" w:hanging="648", so its line ends at
// 566.4px, 864 twips before the column's right edge.
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
async function layOutBody(page: Page, bodyXml: string) {
  const bytes = Array.from(leaderDocx(bodyXml));
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
    await layOutBody(page,
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

  test('a leading decorated tab keeps its run color and font like one after text', async ({ page }) => {
    const styledTab = tab('<w:color w:val="123456"/><w:sz w:val="48"/><w:u w:val="single"/>');
    await layOutBody(page, `<w:p>${styledTab}${run('Name')}</w:p>`
      + `<w:p>${run('By:')}${styledTab}${run('Name')}</w:p>`);
    const lines = await page.locator('[data-docx-tab] > span').evaluateAll(elements => elements.map(line => {
      const style = getComputedStyle(line);
      return { color: style.borderTopColor, size: style.fontSize };
    }));
    expect(lines).toEqual([
      { color: 'rgb(18, 52, 86)', size: '32px' },
      { color: 'rgb(18, 52, 86)', size: '32px' },
    ]);
  });

  test('a bookmark inside a link keeps the whole entry inside that link and on its stops', async ({ page }) => {
    const linked = (bookmark: string) => entry(`<w:hyperlink w:anchor="_Toc9">${bookmark}${run('2.1.')}${tab()}${run('Spillway')}<w:r><w:tab/><w:t>12</w:t></w:r></w:hyperlink>`);
    await layOutBody(page, linked('<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkEnd w:id="1"/>') + linked(''));
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

  for (const closesBeforeTab of [false, true]) test(`a complex HYPERLINK ${closesBeforeTab ? 'before' : 'across'} a tab keeps its label and bookmark inside the link`, async ({ page }) => {
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + '<w:r><w:instrText> HYPERLINK "https://example.com/entry" </w:instrText></w:r>'
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    await layOutBody(page, entry(begin + '<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkEnd w:id="1"/>'
      + run('2.1.') + (closesBeforeTab ? end : '') + tab() + run('Spillway')
      + '<w:r><w:tab/><w:t>12</w:t></w:r>' + (closesBeforeTab ? '' : end)));
    const link = page.locator('a[href="https://example.com/entry"]');
    await expect(link).toHaveCount(1);
    await expect(link).toHaveText(closesBeforeTab ? '2.1.' : '2.1.Spillway12');
    await expect(link.locator('span[id="_Toc1"]')).toHaveCount(1);
    await expect(link.locator('a')).toHaveCount(0);
  });
});
