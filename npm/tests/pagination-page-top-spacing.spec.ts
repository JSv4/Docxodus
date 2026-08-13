import { expect, test, type Page } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  PAGE_TOP_LABELS,
  PAGE_TOP_SPACE_BEFORE_TWIPS,
  generatePageTopSpacingDocx,
} from './docx-page-top-spacing-fixture.js';

/**
 * Issue #428 — Word suppresses paragraph space-before at the top of later pages in a section.
 *
 * The suppression belongs to page placement, not OOXML conversion: the authored margin remains
 * available while blocks are measured, then the clone placed first on a later page drops it.
 * Natural overflow and `w:pageBreakBefore` therefore share one rule. A document/section start is
 * deliberately different and retains the authored spacing, as does an ordinary same-page gap.
 */

const PT_TO_PX = 4 / 3;
const SPACE_BEFORE_PX = PAGE_TOP_SPACE_BEFORE_TWIPS / 20 * PT_TO_PX;
const TOLERANCE_PX = 0.75;
const __dirname = dirname(fileURLToPath(import.meta.url));

interface ParagraphGeometry {
  label: string;
  page: number;
  section: number;
  topFromBody: number;
  gapFromPrevious: number | null;
  marginTop: number;
}

async function renderFixture(page: Page): Promise<ParagraphGeometry[]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined,
    { timeout: 30_000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });

  return page.evaluate(async ({ bytes, labels }) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', true, true, false, false, false,
    ) as string;
    if (html.startsWith('{')) throw new Error(`conversion failed: ${html.slice(0, 300)}`);

    document.body.innerHTML = '<main id="fixture"></main>';
    (window as any).DocxodusPagination.paginateHtml(
      html,
      document.getElementById('fixture'),
      { scale: 1, showPageNumbers: false, pageGap: 0, fragmentParagraphs: false },
    );
    await document.fonts.ready;
    await new Promise<void>(resolve => requestAnimationFrame(() =>
      requestAnimationFrame(() => resolve())));

    const wanted = new Set<string>(Object.values(labels));
    const result: ParagraphGeometry[] = [];
    const pages = Array.from(document.querySelectorAll<HTMLElement>('#fixture .page-box'));
    for (const [pageIndex, pageBox] of pages.entries()) {
      const body = pageBox.querySelector<HTMLElement>('.page-content')!;
      const paragraphs = Array.from(body.querySelectorAll<HTMLElement>(':scope > p'));
      for (const [index, paragraph] of paragraphs.entries()) {
        const label = (paragraph.textContent || '').trim();
        if (!wanted.has(label)) continue;
        const rect = paragraph.getBoundingClientRect();
        const bodyRect = body.getBoundingClientRect();
        const previous = index > 0 ? paragraphs[index - 1].getBoundingClientRect() : null;
        result.push({
          label,
          page: pageIndex + 1,
          section: Number(pageBox.dataset.sectionIndex),
          topFromBody: rect.top - bodyRect.top,
          gapFromPrevious: previous ? rect.top - previous.bottom : null,
          marginTop: parseFloat(getComputedStyle(paragraph).marginTop),
        });
      }
    }
    return result;
  }, {
    bytes: Array.from(generatePageTopSpacingDocx()),
    labels: PAGE_TOP_LABELS,
  });
}

test.describe('paragraph space-before at page boundaries', () => {
  let paragraphs: ParagraphGeometry[];

  test.beforeAll(async ({ browser }) => {
    const page = await browser.newPage();
    try {
      paragraphs = await renderFixture(page);
    } finally {
      await page.close();
    }
  });

  const find = (label: string) => paragraphs.find(paragraph => paragraph.label === label)!;

  test('suppresses space-before after natural pagination', () => {
    const paragraph = find(PAGE_TOP_LABELS.natural);
    expect(paragraph.page).toBe(2);
    expect(paragraph.topFromBody).toBeCloseTo(0, 0);
    expect(paragraph.marginTop).toBeCloseTo(0, 0);
  });

  test('suppresses the same spacing after w:pageBreakBefore', () => {
    const paragraph = find(PAGE_TOP_LABELS.pageBreakBefore);
    expect(paragraph.page).toBe(3);
    expect(paragraph.topFromBody).toBeCloseTo(0, 0);
    expect(paragraph.marginTop).toBeCloseTo(0, 0);
  });

  test('retains space-before at a document or section start', () => {
    const documentStart = find(PAGE_TOP_LABELS.sectionStart);
    const sectionStart = find(PAGE_TOP_LABELS.nextSection);

    expect(documentStart.page).toBe(1);
    expect(sectionStart.page).toBe(4);
    expect(sectionStart.section).toBe(1);
    expect(documentStart.topFromBody).toBeCloseTo(SPACE_BEFORE_PX, 0);
    expect(sectionStart.topFromBody).toBeCloseTo(SPACE_BEFORE_PX, 0);
    expect(documentStart.marginTop).toBeCloseTo(SPACE_BEFORE_PX, 0);
    expect(sectionStart.marginTop).toBeCloseTo(SPACE_BEFORE_PX, 0);
  });

  test('retains ordinary inter-paragraph spacing', () => {
    const paragraph = find(PAGE_TOP_LABELS.samePage);
    expect(paragraph.page).toBe(1);
    expect(paragraph.gapFromPrevious).not.toBeNull();
    expect(Math.abs(paragraph.gapFromPrevious! - SPACE_BEFORE_PX)).toBeLessThanOrEqual(
      TOLERANCE_PX,
    );
    expect(paragraph.marginTop).toBeCloseTo(SPACE_BEFORE_PX, 0);
  });
});

async function paginateMarkupAfterHardBreak(page: Page, tag: 'h1' | 'div') {
  await page.setContent(`
    <style>#staging > div > * { box-sizing: border-box; }</style>
    <div id="staging">
      <div data-section-index="0"
           data-page-width="102" data-page-height="102"
           data-content-width="100" data-content-height="100"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p style="height:20pt;margin:0">PAGE 1</p>
        <div data-page-break="true"></div>
        <${tag} id="target" style="height:20pt;margin:18pt 0 0">TARGET</${tag}>
      </div>
    </div>
    <div id="container"></div>`);
  await page.addScriptTag({ path: join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
    const pages = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    const target = pages[1].querySelector<HTMLElement>('#target')!;
    const body = pages[1].querySelector<HTMLElement>('.page-content')!;
    return {
      pages: pages.length,
      topFromBody: target.getBoundingClientRect().top - body.getBoundingClientRect().top,
      marginTop: parseFloat(getComputedStyle(target).marginTop),
    };
  });
}

test.describe('page-top placement scope', () => {
  test('uses the paragraph rule for an outlined heading after an explicit break', async ({ page }) => {
    const result = await paginateMarkupAfterHardBreak(page, 'h1');
    expect(result.pages).toBe(2);
    expect(result.topFromBody).toBeCloseTo(0, 0);
    expect(result.marginTop).toBeCloseTo(0, 0);
  });

  test('does not erase a non-paragraph block margin', async ({ page }) => {
    const result = await paginateMarkupAfterHardBreak(page, 'div');
    expect(result.pages).toBe(2);
    expect(result.topFromBody).toBeCloseTo(SPACE_BEFORE_PX, 0);
    expect(result.marginTop).toBeCloseTo(SPACE_BEFORE_PX, 0);
  });
});
