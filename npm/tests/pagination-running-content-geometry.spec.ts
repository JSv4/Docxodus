import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import {
  DEFAULT_RUNNING_CONTENT_GEOMETRY,
  PAGE_HEIGHT_PT,
  STORY_TEXT,
  generateRunningContentDocx,
  twipsToPt,
} from './docx-running-content-fixture.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Issue #377 — where headers, footers, and body text sit vertically on a paginated page.
 *
 * `w:pgMar` declares FOUR independent distances from the paper edge: `w:header` to the TOP of
 * the header story, `w:footer` to the BOTTOM of the footer story, and `w:top`/`w:bottom` to the
 * body text. The engine used to ignore the first two and anchor the running content to the
 * MARGINS instead — header bottom-aligned above `w:top`, footer top-aligned below `w:bottom` —
 * which pulled both bands toward the body by exactly `margin − distance` (25 px on a Word
 * default page) and left the top and bottom of the sheet blank.
 *
 * These assertions restate the OOXML contract, not the implementation: each band edge is
 * checked against the paper edge the spec says owns it, and the bands are required to stay
 * disjoint so a tall story can never draw over body text.
 */

const PT_TO_PX = 4 / 3;
/** Sub-pixel tolerance: band offsets are set in `pt` and read back from layout in `px`. */
const TOL = 0.75;

interface BandBox {
  top: number;
  bottom: number;
  height: number;
  text: string;
}

interface PageGeometry {
  page: number;
  sectionIndex: number;
  boxTop: number;
  boxBottom: number;
  header: BandBox | null;
  body: BandBox | null;
  footer: BandBox | null;
}

const MEASURE = () => {
  const box = (el: Element | null): unknown => {
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return {
      top: r.top,
      bottom: r.bottom,
      height: r.height,
      text: (el.textContent || '').replace(/\s+/g, ' ').trim(),
    };
  };
  const container = document.getElementById('pagination-container') as HTMLElement;
  return Array.from(container.querySelectorAll('.page-box')).map((b, i) => {
    const r = b.getBoundingClientRect();
    return {
      page: i + 1,
      sectionIndex: Number((b as HTMLElement).dataset.sectionIndex ?? -1),
      boxTop: r.top,
      boxBottom: r.bottom,
      header: box(b.querySelector('.page-header')),
      body: box(b.querySelector('.page-content')),
      footer: box(b.querySelector('.page-footer')),
    };
  });
};

async function paginateGeneratedDocument(page: Page, docx: Uint8Array): Promise<PageGeometry[]> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });

  const html = await page.evaluate((bytes) => {
    const converter = (window as any).Docxodus.DocumentConverter;
    return converter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
      /* paginationMode */ 1, /* paginationScale */ 1, 'page-',
      /* renderAnnotations */ false, 0, 'annot-',
      /* footnotes */ false, /* headersAndFooters */ true,
      /* trackedChanges */ false, false, false,
    ) as string;
  }, Array.from(docx));

  expect(html.startsWith('{'), `conversion failed: ${html.slice(0, 300)}`).toBe(false);

  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ measure }: { measure: string }) => {
    const staging = document.getElementById('pagination-staging') as HTMLElement;
    const container = document.getElementById('pagination-container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { scale: 1, showPageNumbers: false }).paginate();
    // eslint-disable-next-line no-eval
    return (0, eval)(`(${measure})`)();
  }, { measure: MEASURE.toString() }) as Promise<PageGeometry[]>;
}

/** The section geometry, in CSS pixels, that a page belonging to `sectionIndex` must obey. */
function expectedPx(sectionIndex: number, geometry = DEFAULT_RUNNING_CONTENT_GEOMETRY) {
  const g = sectionIndex === 0 ? geometry.first : geometry.second;
  return {
    marginTop: twipsToPt(g.marginTop) * PT_TO_PX,
    marginBottom: twipsToPt(g.marginBottom) * PT_TO_PX,
    headerDistance: twipsToPt(g.headerDistance) * PT_TO_PX,
    footerDistance: twipsToPt(g.footerDistance) * PT_TO_PX,
  };
}

test.describe('Paginated running-content geometry', () => {
  let pages: PageGeometry[];

  test.beforeAll(async ({ browser }) => {
    const page = await browser.newPage();
    try {
      pages = await paginateGeneratedDocument(page, generateRunningContentDocx());
    } finally {
      await page.close();
    }
  });

  test('every page carries both running stories, across the inheriting section', async () => {
    expect(pages.length).toBeGreaterThanOrEqual(4);
    // Both sections must be represented, or the inheritance half of this fixture is untested.
    expect(new Set(pages.map((p) => p.sectionIndex))).toEqual(new Set([0, 1]));

    const missing = pages.filter((p) => !p.header?.text || !p.footer?.text);
    expect(missing.map((p) => p.page), 'pages missing a header or footer story').toEqual([]);

    const headers = pages.map((p) => p.header!.text);
    expect(headers[0]).toBe(STORY_TEXT.headerFirst);
    expect(pages[0].footer!.text).toBe(STORY_TEXT.footerFirst);
    expect(headers).toContain(STORY_TEXT.headerEven);
    expect(headers).toContain(STORY_TEXT.headerDefault);

    // Section 2 declares no references of its own; it must still show the inherited stories.
    const inherited = pages.filter((p) => p.sectionIndex === 1);
    expect(inherited.length).toBeGreaterThan(0);
    expect(inherited.every((p) => !!p.header!.text && !!p.footer!.text)).toBe(true);
  });

  test('the header story starts at w:header from the top of the paper', async () => {
    for (const p of pages) {
      const expected = expectedPx(p.sectionIndex);
      expect(
        p.header!.top - p.boxTop,
        `page ${p.page} (section ${p.sectionIndex}) header top`,
      ).toBeCloseTo(expected.headerDistance, 0);
    }
  });

  test('the footer story ends at w:footer from the bottom of the paper', async () => {
    for (const p of pages) {
      const expected = expectedPx(p.sectionIndex);
      expect(
        p.boxBottom - p.footer!.bottom,
        `page ${p.page} (section ${p.sectionIndex}) footer bottom`,
      ).toBeCloseTo(expected.footerDistance, 0);
    }
  });

  test('body text starts at the top margin unless the header has already passed it', async () => {
    for (const p of pages) {
      const expected = expectedPx(p.sectionIndex);
      const wanted = Math.max(expected.marginTop, expected.headerDistance + p.header!.height);
      expect(
        p.body!.top - p.boxTop,
        `page ${p.page} (section ${p.sectionIndex}) body top`,
      ).toBeCloseTo(wanted, 0);
    }
  });

  test('body text ends at the bottom margin unless the footer has already passed it', async () => {
    for (const p of pages) {
      const expected = expectedPx(p.sectionIndex);
      const wanted = Math.max(expected.marginBottom, expected.footerDistance + p.footer!.height);
      expect(
        p.boxBottom - p.body!.bottom,
        `page ${p.page} (section ${p.sectionIndex}) body bottom`,
      ).toBeCloseTo(wanted, 0);
    }
  });

  test('the three bands never overlap', async () => {
    const overlaps = pages.filter(
      (p) => p.header!.bottom > p.body!.top + TOL || p.body!.bottom > p.footer!.top + TOL,
    );
    expect(
      overlaps.map((p) => p.page),
      'pages where a running story overlaps the body band',
    ).toEqual([]);
  });

  test('the second section uses its OWN distances, not the first section\'s', async () => {
    const s1 = pages.find((p) => p.sectionIndex === 0)!;
    const s2 = pages.find((p) => p.sectionIndex === 1)!;
    // The fixture deliberately differs in every distance, so equal offsets mean the renderer
    // carried section 1's page setup across the transition.
    expect(s2.header!.top - s2.boxTop).not.toBeCloseTo(s1.header!.top - s1.boxTop, 0);
    expect(s2.boxBottom - s2.footer!.bottom).not.toBeCloseTo(s1.boxBottom - s1.footer!.bottom, 0);
    expect(s2.boxBottom - s2.boxTop).toBeCloseTo(PAGE_HEIGHT_PT * PT_TO_PX, 0);
  });
});

/** Margins too small to contain the stories, so both must push the body inward. */
const TIGHT_GEOMETRY = {
  first: { marginTop: 720, marginBottom: 720, headerDistance: 360, footerDistance: 360 },
  second: { marginTop: 720, marginBottom: 720, headerDistance: 360, footerDistance: 360 },
};

test.describe('Running content that outgrows its margin', () => {
  test('a tall story pushes the body instead of drawing over it', async ({ page }) => {
    const pages = await paginateGeneratedDocument(
      page,
      generateRunningContentDocx(TIGHT_GEOMETRY, 120, 4),
    );

    expect(pages.length).toBeGreaterThan(0);
    for (const p of pages) {
      const expected = expectedPx(p.sectionIndex, TIGHT_GEOMETRY);
      // The fixture is built so the stories really do overflow; if they stopped doing so the
      // rest of this test would pass vacuously.
      expect(
        expected.headerDistance + p.header!.height,
        `page ${p.page} header should reach past the top margin`,
      ).toBeGreaterThan(expected.marginTop);

      expect(p.body!.top - p.boxTop, `page ${p.page} body top`)
        .toBeCloseTo(expected.headerDistance + p.header!.height, 0);
      expect(p.boxBottom - p.body!.bottom, `page ${p.page} body bottom`)
        .toBeCloseTo(expected.footerDistance + p.footer!.height, 0);
      expect(p.header!.bottom, `page ${p.page} header/body overlap`)
        .toBeLessThanOrEqual(p.body!.top + TOL);
      expect(p.body!.bottom, `page ${p.page} body/footer overlap`)
        .toBeLessThanOrEqual(p.footer!.top + TOL);
    }
  });
});

/**
 * A page with no running story at all. `w:header`/`w:footer` are still declared — every real
 * `w:pgMar` carries them — but they reserve nothing when there is no story to place, so the body
 * must sit on its margins. Anchoring the body at `headerDistance` unconditionally moved every
 * header-less document's text down by the distance instead.
 */
const NO_STORY_STAGING = `
  <style>#staging { font: 12px/12px Arial; } #staging p { margin: 0; height: 12pt; }</style>
  <div id="pagination-staging">
    <div data-section-index="0"
         data-page-width="300" data-page-height="300"
         data-content-width="276" data-content-height="276"
         data-margin-top="12" data-margin-right="12"
         data-margin-bottom="12" data-margin-left="12"
         data-header-height="36" data-footer-height="36">
      ${Array.from({ length: 40 }, (_, i) => `<p>line ${i}</p>`).join('')}
    </div>
  </div>
  <div id="pagination-container"></div>`;

test.describe('Pages with no running content', () => {
  test('a declared header distance reserves nothing when the section has no story', async ({
    page,
  }) => {
    await page.setContent(NO_STORY_STAGING);
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const measured = await page.evaluate(({ measure }: { measure: string }) => {
      const staging = document.getElementById('pagination-staging') as HTMLElement;
      const container = document.getElementById('pagination-container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      new PaginationEngine(staging, container, { scale: 1, showPageNumbers: false }).paginate();
      // eslint-disable-next-line no-eval
      return (0, eval)(`(${measure})`)();
    }, { measure: MEASURE.toString() }) as PageGeometry[];

    expect(measured.length).toBeGreaterThan(1);
    for (const p of measured) {
      expect(p.header, `page ${p.page} should have no header band`).toBeNull();
      expect(p.footer, `page ${p.page} should have no footer band`).toBeNull();
      // 12 pt margins on a 300 pt page, with a 36 pt header distance that must not bind.
      expect(p.body!.top - p.boxTop, `page ${p.page} body top`).toBeCloseTo(12 * PT_TO_PX, 0);
      expect(p.boxBottom - p.body!.bottom, `page ${p.page} body bottom`)
        .toBeCloseTo(12 * PT_TO_PX, 0);
    }
  });
});
