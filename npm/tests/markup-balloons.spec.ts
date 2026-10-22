import { test, expect, type Page } from '@playwright/test';
import { markupBalloonsFixture } from './docx-markup-balloons-fixture.js';

/** Inert data the Word presentation records for balloons; nothing reads it under other placements. */
const BALLOON_DATA = ['data-format-description', 'data-format-key', 'data-author-color', 'data-initials',
  'data-font-family', 'data-balloon-font-family', 'data-balloon-font-size'];

async function paginate(page: Page, paper: 'letter' | 'a4', placement: 'balloons' | 'margin', stripBalloonData = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  return page.evaluate(({ bytes, placement, stripBalloonData, balloonData }) => {
    // Word presentation (last argument), margin comments (2), paginated.
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Cedar markup', 'docx-', false, '', 2, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', false, false, true, true, true, false, null, false, 1);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    if (stripBalloonData) {
      for (const name of balloonData) for (const element of Array.from(parsed.querySelectorAll(`[${name}]`))) {
        element.removeAttribute(name);
      }
    }
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.innerHTML = '<div id="host"></div>';
    const host = document.querySelector<HTMLElement>('#host')!;
    (window as any).DocxodusPagination.paginateHtml(parsed.body.innerHTML, host, {
      showPageNumbers: false, commentColumnPlacement: placement,
    });
    const pt = (px: number) => px * 0.75;
    const pages = Array.from(host.querySelectorAll<HTMLElement>('.page-box'));
    const balloons = Array.from(host.querySelectorAll<HTMLElement>('.page-balloon')).map(balloon => {
      const r = balloon.getBoundingClientRect();
      const box = balloon.closest<HTMLElement>('.page-box')!.getBoundingClientRect();
      return { text: balloon.textContent!.replace(/\s+/g, ' ').trim(), border: getComputedStyle(balloon).borderTopColor,
        top: pt(r.top - box.top), bottom: pt(r.bottom - box.top), left: pt(r.left - box.left), right: pt(r.right - box.left) };
    });
    const anchor = (text: string) => {
      const element = Array.from(host.querySelectorAll<HTMLElement>('.page-box span'))
        .find(span => span.textContent === text)!;
      return pt(element.getBoundingClientRect().top - element.closest('.page-box')!.getBoundingClientRect().top);
    };
    return {
      pages: pages.map(p => ({ width: pt(p.getBoundingClientRect().width), height: pt(p.getBoundingClientRect().height),
        scale: Number(p.dataset.docxMarkupScale ?? '1') })),
      balloons,
      leaders: host.querySelectorAll('.page-box svg path[stroke-dasharray]').length,
      markerWidths: Array.from(host.querySelectorAll<HTMLElement>('.page-box a.comment-marker'))
        .map(a => a.getBoundingClientRect().width),
      buyerTop: placement === 'balloons' ? anchor('buyer') : 0,
      pageTree: (() => {
        const clone = host.cloneNode(true) as HTMLElement;
        for (const name of balloonData) for (const element of Array.from(clone.querySelectorAll(`[${name}]`))) {
          element.removeAttribute(name);
        }
        return clone.innerHTML;
      })(),
    };
  }, { bytes: Array.from(markupBalloonsFixture(paper)), placement, stripBalloonData, balloonData: BALLOON_DATA });
}

for (const [paper, width, height] of [['letter', 612, 792], ['a4', 595.3, 841.9]] as const) {
  test(`balloons keep the ${paper} paper size and shrink the page beside the markup area`, async ({ page }) => {
    const result = await paginate(page, paper, 'balloons');
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
  const result = await paginate(page, 'letter', 'balloons');
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

test('balloons sit level with their line where room allows and never overlap', async ({ page }) => {
  const { balloons, buyerTop, pages } = await paginate(page, 'letter', 'balloons');
  expect(balloons[0].top).toBeCloseTo(buyerTop, 0);
  for (let i = 1; i < balloons.length; i++) expect(balloons[i].top).toBeGreaterThanOrEqual(balloons[i - 1].bottom);
  expect(balloons[balloons.length - 1].bottom).toBeLessThanOrEqual(pages[0].height);
});

test('the margin placement ignores the balloon data the Word presentation records', async ({ page }) => {
  const recorded = await paginate(page, 'letter', 'margin');
  const stripped = await paginate(page, 'letter', 'margin', true);
  expect(recorded.pages[0].scale).toBe(1);
  expect(recorded.balloons).toHaveLength(0);
  expect(recorded.pageTree).toBe(stripped.pageTree);
});
