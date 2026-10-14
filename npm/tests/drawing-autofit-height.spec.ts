import { expect, test, type Page } from '@playwright/test';
import {
  generateDrawingAnchorDocx,
  type DrawingAnchorFixtureOptions,
} from './docx-drawing-anchor-fixture.js';

/**
 * Issue #396 — how tall a DrawingML textbox is.
 *
 * `wps:bodyPr/a:spAutoFit` means Word sizes the SHAPE to its laid-out text plus the body insets,
 * and treats the stored `wp:extent`/`a:ext cy` as a cache of the last layout rather than as the
 * height. Without it, the stored extent IS the height and content must not move it.
 *
 * The renderer already dropped the CSS height for an auto-fit box, so it already grew to its
 * content; what was wrong was the CONTENT height. OOXML automatic line spacing (`w:lineRule="auto"`)
 * is a multiple of the font's own line box, not a percentage of its em square, so every line of an
 * auto-fit box was short by the ratio between the two — and because the box is content-driven,
 * that text error surfaced as a BOX-SIZE error against the visual-parity oracle.
 *
 * These assertions pin the two halves separately, so a regression in one cannot hide behind the
 * other: an auto-fit box must track its content (line count AND line spacing), and a fixed-extent
 * box must ignore both.
 */

const EMU_PER_PX = 9525;
/** Sub-pixel tolerance: geometry is authored in EMU, emitted in pt, and read back in px. */
const TOL = 1.0;

/** Deliberately long enough to wrap to several lines in the generated 2in-wide box. */
const WRAPPING_TEXT =
  'Auto fit height is driven by the laid out text of the shape rather than by its stored extent.';

const BASE: DrawingAnchorFixtureOptions = {
  horizontal: { relativeFrom: 'column', offsetEmu: 0 },
  vertical: { relativeFrom: 'paragraph', offsetEmu: 0 },
};

interface BoxGeometry {
  height: number;
  width: number;
  textHeight: number;
  lines: number;
  paddingTop: number;
  paddingBottom: number;
  borderTop: number;
  borderBottom: number;
  cssHeight: string;
  autoFitAttr: string | null;
}

async function measure(page: Page, options: DrawingAnchorFixtureOptions): Promise<BoxGeometry> {
  const bytes = generateDrawingAnchorDocx(options);
  return page.evaluate(async (input) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', -1, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', false, false, false, false, false);
    document.body.innerHTML = '<main id="autofit-root"></main>';
    const root = document.getElementById('autofit-root')!;
    (window as any).DocxodusPagination.paginateHtml(html, root, {
      scale: 1, showPageNumbers: false, fragmentParagraphs: false,
    });
    await document.fonts.ready;
    await new Promise<void>(r => requestAnimationFrame(() => requestAnimationFrame(() => r())));

    // Query inside the RENDERED page box: pagination also keeps a hidden staging copy of the
    // content, whose client rects are all zero.
    const pageBox = root.querySelector<HTMLElement>('.page-box')!;
    const anchor = pageBox.querySelector<HTMLElement>('[data-docx-drawing-anchor="true"]')!;
    const text = anchor.firstElementChild as HTMLElement;
    const rect = anchor.getBoundingClientRect();
    const style = getComputedStyle(anchor);

    // Count rendered line boxes by distinct client-rect tops of the text block's contents.
    const range = document.createRange();
    range.selectNodeContents(text);
    const tops = new Set(Array.from(range.getClientRects())
      .filter(r => r.width > 1 && r.height > 1)
      .map(r => Math.round(r.top)));

    return {
      height: rect.height,
      width: rect.width,
      textHeight: text.getBoundingClientRect().height,
      lines: tops.size,
      paddingTop: parseFloat(style.paddingTop),
      paddingBottom: parseFloat(style.paddingBottom),
      borderTop: parseFloat(style.borderTopWidth) || 0,
      borderBottom: parseFloat(style.borderBottomWidth) || 0,
      cssHeight: style.height,
      autoFitAttr: anchor.getAttribute('data-docx-anchor-autofit'),
    };
  }, bytes);
}

test.beforeEach(async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined,
    { timeout: 30_000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });
});

test.describe('DrawingML textbox auto-fit height', () => {
  test('an auto-fit box is exactly its laid-out text plus the bodyPr insets', async ({ page }) => {
    // A stored extent far taller than the content: Word would ignore it under spAutoFit.
    const box = await measure(page, {
      ...BASE, autoFit: true, text: WRAPPING_TEXT, extentEmu: { cx: 1828800, cy: 3657600 },
    });

    expect(box.autoFitAttr).toBe('true');
    expect(box.lines).toBeGreaterThan(1);
    // tIns=76200 EMU and bIns=12700 EMU, plus whatever border the shape draws.
    const expected = box.textHeight + box.paddingTop + box.paddingBottom +
      box.borderTop + box.borderBottom;
    expect(box.height).toBeCloseTo(expected, 0);
    // The stale stored extent (3657600 EMU = 384px) must not be the height.
    expect(box.height).toBeLessThan(3657600 / EMU_PER_PX - 1);
  });

  test('an auto-fit box grows with line COUNT', async ({ page }) => {
    const short = await measure(page, { ...BASE, autoFit: true, text: 'One line' });
    const long = await measure(page, { ...BASE, autoFit: true, text: WRAPPING_TEXT });

    expect(short.lines).toBe(1);
    expect(long.lines).toBeGreaterThan(short.lines);
    // Each extra line adds exactly one line box, so the growth is proportional.
    const lineHeight = (long.height - short.height) / (long.lines - short.lines);
    expect(lineHeight).toBeGreaterThan(0);
    expect(long.height).toBeCloseTo(short.height + lineHeight * (long.lines - short.lines), 0);
  });

  /**
   * The regression that made the tracked `shape` case severe. `w:line="360"` is 1.5 lines of
   * AUTOMATIC spacing — a multiple of the font's line box. Emitting it as a percentage of
   * font-size under-measured every line, so the content-driven box came out short.
   */
  test('an auto-fit box grows with automatic line SPACING, measured against the font line box',
    async ({ page }) => {
      const single = await measure(page, {
        ...BASE, autoFit: true, text: WRAPPING_TEXT, lineTwips: 240,
      });
      const oneAndAHalf = await measure(page, {
        ...BASE, autoFit: true, text: WRAPPING_TEXT, lineTwips: 360,
      });

      expect(oneAndAHalf.lines).toBe(single.lines);
      expect(oneAndAHalf.height).toBeGreaterThan(single.height);

      // 360/240 = 1.5x the single-spaced line box, applied to every line.
      const singleLineBox = single.textHeight / single.lines;
      expect(oneAndAHalf.textHeight / oneAndAHalf.lines).toBeCloseTo(singleLineBox * 1.5, 0);

      // The single-spaced line box is the FONT's, not the em square: Arial at 12pt (24 half-points)
      // is 16px per em and its natural line box is materially taller. A percentage-of-font-size
      // model would collapse this to 16px and shrink the box.
      expect(singleLineBox).toBeGreaterThan(16 * 1.1);
    });

  test('a fixed-extent box keeps its stored height regardless of content', async ({ page }) => {
    const short = await measure(page, { ...BASE, text: 'One line' });
    const long = await measure(page, { ...BASE, text: WRAPPING_TEXT });

    expect(short.autoFitAttr).toBeNull();
    expect(long.lines).toBeGreaterThan(short.lines);
    // 914400 EMU = 96px, the fixture's default stored extent.
    expect(short.height).toBeCloseTo(914400 / EMU_PER_PX, 0);
    expect(long.height).toBeCloseTo(914400 / EMU_PER_PX, 0);
    expect(long.height).toBeCloseTo(short.height, 0);
  });

  test('a fixed-extent box ignores automatic line spacing too', async ({ page }) => {
    const single = await measure(page, { ...BASE, text: WRAPPING_TEXT, lineTwips: 240 });
    const oneAndAHalf = await measure(page, { ...BASE, text: WRAPPING_TEXT, lineTwips: 360 });

    expect(oneAndAHalf.height).toBeCloseTo(single.height, 0);
    expect(single.height).toBeCloseTo(914400 / EMU_PER_PX, 0);
  });

  test('auto-fit constrains only the height; the stored width still applies', async ({ page }) => {
    const box = await measure(page, {
      ...BASE, autoFit: true, text: WRAPPING_TEXT, extentEmu: { cx: 1828800, cy: 3657600 },
    });
    // 1828800 EMU = 192px.
    expect(box.width).toBeCloseTo(1828800 / EMU_PER_PX, 0);
    expect(Math.abs(box.width - 1828800 / EMU_PER_PX)).toBeLessThanOrEqual(TOL);
  });
});
