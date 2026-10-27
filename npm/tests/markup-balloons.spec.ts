import { test, expect, type Page } from '@playwright/test';
import {
  anchorlessCommentsFixture, balloonDocx, commented, crowdedCommentsFixture, formatted, markedParagraph,
  markupBalloonsFixture, plain, reformatted,
} from './docx-markup-balloons-fixture.js';

/** Inert data the Word presentation records for balloons; nothing reads it under other placements. */
const BALLOON_DATA = ['data-format-description', 'data-format-key', 'data-author-color', 'data-initials',
  'data-font-family', 'data-balloon-font-family', 'data-balloon-font-size'];

async function open(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
}

/** Converts in the Word presentation and paginates into a fresh host, in the page's current document. */
async function paginate(page: Page, bytes: Uint8Array, placement: 'balloons' | 'margin', stripBalloonData = false) {
  return page.evaluate(({ bytes, placement, stripBalloonData, balloonData }) => {
    // Word presentation (last argument), margin comments (2), paginated, headers rendered.
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Cedar markup', 'docx-', false, '', 2, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', false, true, true, true, true, false, null, false, 1);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    if (stripBalloonData) {
      for (const name of balloonData) for (const element of Array.from(parsed.querySelectorAll(`[${name}]`))) {
        element.removeAttribute(name);
      }
    }
    // Stylesheets accumulate in the head, as they do when a viewer re-paginates.
    document.head.insertAdjacentHTML('beforeend', parsed.head.innerHTML);
    document.body.innerHTML = '<div id="host"></div>';
    const host = document.querySelector<HTMLElement>('#host')!;
    (window as any).DocxodusPagination.paginateHtml(parsed.body.innerHTML, host, {
      showPageNumbers: false, commentColumnPlacement: placement,
    });
    const pt = (px: number) => px * 0.75;
    const pages = Array.from(host.querySelectorAll<HTMLElement>('.page-box'));
    const pageOf = (element: Element) => pages.indexOf(element.closest<HTMLElement>('.page-box')!);
    const balloons = Array.from(host.querySelectorAll<HTMLElement>('.page-balloon')).map(balloon => {
      const r = balloon.getBoundingClientRect();
      const box = balloon.closest<HTMLElement>('.page-box')!.getBoundingClientRect();
      return { text: balloon.textContent!.replace(/\s+/g, ' ').trim(), border: getComputedStyle(balloon).borderTopColor,
        page: pageOf(balloon), truncated: balloon.dataset.balloonTruncated === 'true',
        top: pt(r.top - box.top), bottom: pt(r.bottom - box.top), left: pt(r.left - box.left), right: pt(r.right - box.left) };
    });
    const spanTop = (text: string) => {
      const span = Array.from(host.querySelectorAll<HTMLElement>('.page-box span')).filter(s => s.textContent === text).pop()!;
      return { page: pageOf(span), top: pt(span.getBoundingClientRect().top - span.closest('.page-box')!.getBoundingClientRect().top) };
    };
    return {
      pages: pages.map(p => ({ width: pt(p.getBoundingClientRect().width), height: pt(p.getBoundingClientRect().height),
        scale: Number(p.dataset.docxMarkupScale ?? '1') })),
      balloons,
      leaders: host.querySelectorAll('.page-box svg path[stroke-dasharray]').length,
      markerWidths: Array.from(host.querySelectorAll<HTMLElement>('.page-box a.comment-marker'))
        .map(a => a.getBoundingClientRect().width),
      highlights: Array.from(host.querySelectorAll<HTMLElement>('.page-box span.comment-highlight'))
        .map(span => getComputedStyle(span).backgroundColor),
      buyerTop: placement === 'balloons' && host.textContent!.includes('buyer') ? spanTop('buyer').top : 0,
      rangeEnd: host.textContent!.includes('range ends') ? spanTop('range ends') : undefined,
      pageTree: (() => {
        const clone = host.cloneNode(true) as HTMLElement;
        for (const name of balloonData) for (const element of Array.from(clone.querySelectorAll(`[${name}]`))) {
          element.removeAttribute(name);
        }
        return clone.innerHTML;
      })(),
    };
  }, { bytes: Array.from(bytes), placement, stripBalloonData, balloonData: BALLOON_DATA });
}

function expectNoOverlap(balloons: Array<{ page: number; top: number; bottom: number }>) {
  for (let i = 1; i < balloons.length; i++) {
    if (balloons[i].page === balloons[i - 1].page) expect(balloons[i].top).toBeGreaterThanOrEqual(balloons[i - 1].bottom);
  }
}

for (const [paper, width, height] of [['letter', 612, 792], ['a4', 595.3, 841.9]] as const) {
  test(`balloons keep the ${paper} paper size and shrink the page beside the markup area`, async ({ page }) => {
    await open(page);
    const result = await paginate(page, markupBalloonsFixture(paper), 'balloons');
    expect(result.pages).toHaveLength(1);
    expect(result.pages[0].width).toBeCloseTo(width, 0);
    expect(result.pages[0].height).toBeCloseTo(height, 0);
    // The page as far as its text column's right edge (1in margins) plus the 3.7in area fits the paper.
    expect(result.pages[0].scale).toBeCloseTo(width / (width - 72 + 266.4), 4);
    for (const balloon of result.balloons) {
      expect(balloon.left).toBeGreaterThan((width - 72) * result.pages[0].scale);
      expect(balloon.right).toBeLessThanOrEqual(width);
    }
  });
}

test('balloons hold comments, replies and formatting changes in Word\'s wording, one leader each', async ({ page }) => {
  await open(page);
  const result = await paginate(page, markupBalloonsFixture(), 'balloons');
  expect(result.balloons.map(b => b.text)).toEqual([
    'Formatted: Font: Bold',
    'Formatted: Indent: Left: 0.5"',
    'Commented [AA1]: Is delivery the right trigger?',
    'Commented [BB2R1]: Yes, it matches the risk clause.',
    'Commented [BB3]: One.',
    'Commented [AA4]: Two.',
    'Commented [BB5]: Three.',
  ]);
  // Revision and comment authors are numbered on their own; Alice comes first in both.
  expect(result.balloons.map(b => b.border)).toEqual([
    'rgb(0, 120, 212)', 'rgb(0, 120, 212)', 'rgb(0, 120, 212)', 'rgb(209, 52, 56)',
    'rgb(209, 52, 56)', 'rgb(0, 120, 212)', 'rgb(209, 52, 56)',
  ]);
  expect(result.leaders).toBe(result.balloons.length);
  // Word prints no comment reference marks.
  expect(result.markerWidths.length).toBeGreaterThan(0);
  for (const width of result.markerWidths) expect(width).toBe(0);
});

test('point, empty-range and ranged comments all get balloons', async ({ page }) => {
  await open(page);
  const { balloons } = await paginate(page, anchorlessCommentsFixture(), 'balloons');
  expect(balloons.map(b => b.text)).toEqual(
    ['Commented [AA1]: PointNote', 'Commented [BB2]: EmptyNote', 'Commented [AA3]: RangeNote']);
});

test('balloons sit level with their line where room allows', async ({ page }) => {
  await open(page);
  const { balloons, buyerTop, pages } = await paginate(page, markupBalloonsFixture(), 'balloons');
  expect(balloons[0].top).toBeCloseTo(buyerTop, 0);
  expectNoOverlap(balloons);
  expect(balloons[balloons.length - 1].bottom).toBeLessThanOrEqual(pages[0].height);
});

for (const [name, count, perLine] of [['sixty one-line comments on one line', 60, 60], ['twenty-five long comments', 25, 1]] as const) {
  test(`balloons never overlap, even with ${name}`, async ({ page }) => {
    await open(page);
    const text = perLine === 1 ? 'a long remark that wraps over several balloon lines. '.repeat(6) : 'short';
    const { balloons } = await paginate(page, crowdedCommentsFixture(count, perLine, text), 'balloons');
    expect(balloons).toHaveLength(count);
    expectNoOverlap(balloons);
  });
}

test('a comment whose range crosses a page break is anchored where the range ends', async ({ page }) => {
  await open(page);
  const filler = Array.from({ length: 44 }, (_, i) => plain(`Filler line ${i + 1}.`)).join('');
  const body = `${filler}<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">the range starts</w:t></w:r></w:p>` +
    `${plain('in the middle').repeat(4)}<w:p><w:r><w:t>range ends</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
    '<w:r><w:commentReference w:id="0"/></w:r></w:p>';
  const result = await paginate(page, balloonDocx({ body, comments: [{ id: 0, author: 'Alice Archer', initials: 'AA', text: 'Spans.' }] }), 'balloons');
  expect(result.pages.length).toBeGreaterThan(1);
  expect(result.balloons).toHaveLength(1);
  expect(result.balloons[0].page).toBe(result.rangeEnd!.page);
  expect(result.balloons[0].page).toBeGreaterThan(0);
  expect(result.balloons[0].top).toBeCloseTo(result.rangeEnd!.top, 0);
});

test('formatting changes in a header get balloons on every page the header prints on', async ({ page }) => {
  await open(page);
  const header = `<w:p>${formatted(21, 'Header run')}</w:p>${reformatted(22, 'Header paragraph', '<w:jc w:val="center"/>')}`;
  const body = Array.from({ length: 60 }, (_, i) => plain(`Body line ${i + 1}.`)).join('');
  const result = await paginate(page, balloonDocx({ body, header }), 'balloons');
  expect(result.pages).toHaveLength(2);
  for (const pageIndex of [0, 1]) {
    expect(result.balloons.filter(b => b.page === pageIndex).map(b => b.text).sort())
      .toEqual(['Formatted: Centered', 'Formatted: Font: Bold']);
  }
});

test('a balloon pagination leaves later paginations in the same document unchanged', async ({ page }) => {
  await open(page);
  const bytes = markupBalloonsFixture();
  const before = await paginate(page, bytes, 'margin');
  await paginate(page, bytes, 'balloons');
  const after = await paginate(page, bytes, 'margin');
  expect(after.markerWidths).toEqual(before.markerWidths);
  expect(after.markerWidths.some(width => width > 0)).toBe(true);
  expect(after.highlights).toEqual(before.highlights);
  expect(after.pageTree).toBe(before.pageTree);
});

test('the margin placement ignores the balloon data the Word presentation records', async ({ page }) => {
  await open(page);
  const recorded = await paginate(page, markupBalloonsFixture(), 'margin');
  const stripped = await paginate(page, markupBalloonsFixture(), 'margin', true);
  expect(recorded.pages[0].scale).toBe(1);
  expect(recorded.balloons).toHaveLength(0);
  expect(recorded.pageTree).toBe(stripped.pageTree);
});

/** The top of the line holding the span whose text is `text`, in points from its page box's top. */
async function lineTopOf(page: Page, text: string): Promise<number> {
  return page.evaluate((text) => {
    const span = Array.from(document.querySelectorAll<HTMLElement>('#host .page-box span')).find(s => s.textContent === text)!;
    return (span.getBoundingClientRect().top - span.closest('.page-box')!.getBoundingClientRect().top) * 0.75;
  }, text);
}

test('one formatting revision runs on across paragraphs until a tracked paragraph mark is crossed', async ({ page }) => {
  await open(page);
  // The same bold change in three paragraphs. Only the third paragraph's own mark is tracked, and it
  // comes after the third change: nothing tracked lies between the changes, so they are one revision.
  const three = `<w:p>${formatted(41, 'alpha')}</w:p><w:p>${formatted(42, 'beta')}</w:p>` +
    markedParagraph(90, 'ins', formatted(43, 'gamma'));
  const merged = await paginate(page, balloonDocx({ body: three }), 'balloons');
  expect(merged.balloons.map(b => b.text)).toEqual(['Formatted: Font: Bold']);
  // A fourth paragraph's change lies beyond the third's tracked mark: a second revision, at its own line.
  const split = await paginate(page, balloonDocx({ body: three + `<w:p>${formatted(44, 'delta')}</w:p>` }), 'balloons');
  expect(split.balloons.map(b => b.text)).toEqual(['Formatted: Font: Bold', 'Formatted: Font: Bold']);
  expect(split.balloons[0].top).toBeCloseTo(await lineTopOf(page, 'alpha'), 0);
  expect(split.balloons[1].top).toBeCloseTo(await lineTopOf(page, 'delta'), 0);
});

test('a tracked mark on the first paragraph ends the formatting revision there', async ({ page }) => {
  await open(page);
  const body = markedParagraph(91, 'del', formatted(45, 'alpha')) + `<w:p>${formatted(46, 'beta')}</w:p>`;
  const { balloons, pages } = await paginate(page, balloonDocx({ body }), 'balloons');
  expect(balloons.map(b => b.text)).toEqual(['Formatted: Font: Bold', 'Formatted: Font: Bold']);
  // Anchored at its own line, but pushed just below the first balloon, which is taller than a text line.
  expect(balloons[1].top).toBeCloseTo(balloons[0].bottom + 1 * pages[0].scale, 0);
});

test('a formatting change in the header and one in the body never share a balloon', async ({ page }) => {
  await open(page);
  const header = `<w:p>${formatted(21, 'Header run')}</w:p>`;
  const body = `<w:p>${formatted(22, 'Body run')}</w:p>${plain('Body line.')}`;
  const { balloons } = await paginate(page, balloonDocx({ body, header }), 'balloons');
  expect(balloons.map(b => b.text)).toEqual(['Formatted: Font: Bold', 'Formatted: Font: Bold']);
});
