import { expect, test } from '@playwright/test';
import {
  TOC_AFTER_TWIPS,
  TOC_ENTRIES,
  TOC_HYPERLINK_RGB,
  TOC_LINE_TWIPS,
  generateTocDocx,
} from './docx-toc-fixture.js';

/**
 * Issue #397 — TOC entry line geometry and hyperlink appearance in print layout.
 *
 * The issue named two things, and they turned out to have different answers, so they are pinned
 * SEPARATELY here and neither can hide behind the other:
 *
 *  - **Line-box height was a renderer bug.** TOC entries drifted further down the page with every
 *    entry because automatic line spacing was measured against font-size instead of the font's own
 *    line box (issue #396). Fixed; on the tracked HC022 fixture entries now land within 0.12px of
 *    LibreOffice, where they were off by up to 14.7px.
 *  - **Hyperlink appearance is NOT a renderer bug.** The entry run carries
 *    `w:rStyle w:val="Hyperlink"`, and that character style declares `w:color` and `w:u`. Docxodus
 *    paints exactly the declared color; LibreOffice paints the entries black, ignoring the style it
 *    was given. Docxodus follows the file, so the difference is a reference deviation and the
 *    output must not be changed to match the comparison implementation.
 *
 * The second point only means anything if the renderer is proven to be READING the style rather
 * than decorating every `w:hyperlink` it sees. The fixture therefore also contains an entry with
 * identical markup minus `w:rStyle`, and the test requires that one to be undecorated.
 */

const PT_TO_PX = 4 / 3;
const TWIPS_PER_PT = 20;
/** Sub-pixel tolerance: spacing is authored in twips, emitted in pt, and read back in px. */
const TOL = 0.75;

interface EntryGeometry {
  text: string;
  top: number;
  bottom: number;
  lineBox: number;
  color: string;
  textDecorationLine: string;
  fontSize: number;
}

test.beforeEach(async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined,
    { timeout: 30_000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });
});

async function renderToc(page: import('@playwright/test').Page) {
  const bytes = Array.from(generateTocDocx());
  return page.evaluate(async (input) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', true, '', -1, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', true, true, false, false, false);
    document.body.style.cssText = 'margin:0;padding:0;background:white;';
    document.body.innerHTML = '<main id="toc-root"></main>';
    (window as any).DocxodusPagination.paginateHtml(html, document.getElementById('toc-root'), {
      scale: 1, showPageNumbers: false, pageGap: 0, fragmentParagraphs: false,
    });
    await document.fonts.ready;
    await new Promise<void>(res => requestAnimationFrame(() => requestAnimationFrame(() => res())));

    // The RENDERED page box only: pagination keeps a hidden staging copy with zero client rects.
    const pageBox = document.querySelector('#toc-root .page-box') as HTMLElement;
    const paragraphs = Array.from(pageBox.querySelectorAll('p'));
    const entries = paragraphs
      .filter(p => (p.textContent || '').includes('heading of the generated document') ||
        (p.textContent || '').includes('no character style'))
      .map(p => {
        const rect = p.getBoundingClientRect();
        // The entry's own text run — NOT the leader/page-number runs, whose appearance is a
        // separate question. Runs nest (the converter wraps them), and querySelectorAll is in
        // document order, so the LAST span carrying the entry text is the innermost one: the run
        // that actually holds the character style.
        const carriers = Array.from(p.querySelectorAll('span'))
          .filter(span => (span.textContent || '').includes('generated document') ||
            (span.textContent || '').includes('no character style'));
        const run = carriers[carriers.length - 1] as HTMLElement;
        const runStyle = getComputedStyle(run);
        return {
          text: (p.textContent || '').slice(0, 30),
          top: rect.top, bottom: rect.bottom,
          // Every entry is one line and `w:after` becomes a margin, which is outside the border
          // box — so the paragraph's own height IS the line box. (The run's client rect is the
          // glyph box, which is a different and smaller thing.)
          lineBox: rect.height,
          color: runStyle.color,
          textDecorationLine: runStyle.textDecorationLine,
          fontSize: parseFloat(runStyle.fontSize),
        };
      });

    // A one-line paragraph with no auto-spacing multiplier gives the font's NATIVE line box, which
    // is the base the OOXML multiple is applied to. The heading is a different size, so measure the
    // base from a probe at the entry's own size instead of hard-coding a font metric.
    const probe = document.createElement('p');
    probe.style.cssText = 'line-height: normal; margin: 0; position: absolute; visibility: hidden;';
    probe.textContent = 'Probe';
    const entryRun = pageBox.querySelector('p:nth-of-type(2) span') as HTMLElement;
    const runStyle = getComputedStyle(entryRun);
    // Copy the font WITHOUT the `font` shorthand: that shorthand carries line-height, which would
    // hand the probe the very multiplied value it is supposed to provide the base for.
    probe.style.fontFamily = runStyle.fontFamily;
    probe.style.fontSize = runStyle.fontSize;
    probe.style.fontWeight = runStyle.fontWeight;
    probe.style.fontStyle = runStyle.fontStyle;
    probe.style.lineHeight = 'normal';
    pageBox.appendChild(probe);
    const nativeLineBox = probe.getBoundingClientRect().height;
    probe.remove();

    return { entries: entries as EntryGeometry[], nativeLineBox };
  }, bytes);
}

test.describe('TOC entry line geometry', () => {
  test('an entry line box is the OOXML multiple of the FONT line box, not of font-size',
    async ({ page }) => {
      const { entries, nativeLineBox } = await renderToc(page);
      expect(entries.length).toBe(TOC_ENTRIES.length);

      const multiple = TOC_LINE_TWIPS / 240;
      const expected = nativeLineBox * multiple;
      for (const entry of entries) {
        expect(entry.lineBox, `${entry.text}: line box`).toBeCloseTo(expected, 0);
        // The failure mode this replaces: a percentage of the em square, which is materially
        // smaller because the font's natural line box is taller than 1 em.
        expect(entry.lineBox).toBeGreaterThan(entry.fontSize * multiple + 1);
      }
    });

  test('entries are evenly spaced, so displacement cannot accumulate down the page',
    async ({ page }) => {
      const { entries, nativeLineBox } = await renderToc(page);
      const step = nativeLineBox * (TOC_LINE_TWIPS / 240) +
        (TOC_AFTER_TWIPS / TWIPS_PER_PT) * PT_TO_PX;

      const steps = entries.slice(1).map((entry, index) => entry.top - entries[index].top);
      for (const [index, measured] of steps.entries()) {
        expect(measured, `entry ${index + 1} to ${index + 2}`).toBeCloseTo(step, 0);
      }
      // Any per-entry error compounds; the last entry is where it is loudest.
      const drift = Math.abs((entries[entries.length - 1].top - entries[0].top) -
        step * (entries.length - 1));
      expect(drift, 'accumulated drift across the whole TOC').toBeLessThan(TOL);
    });
});

test.describe('TOC hyperlink appearance', () => {
  /**
   * The reference deviation, pinned as a positive statement about our own output: the declared
   * character style is applied. LibreOffice renders these entries black — see BASELINE.md — but it
   * is a comparison implementation, not the correctness oracle, and the file says otherwise.
   */
  test('an entry run styled `Hyperlink` gets the character style\'s declared color and underline',
    async ({ page }) => {
      const { entries } = await renderToc(page);
      const styled = entries.filter((_, index) => TOC_ENTRIES[index].styled);
      expect(styled.length).toBeGreaterThan(0);

      for (const entry of styled) {
        expect(entry.color, `${entry.text}: color`).toBe(TOC_HYPERLINK_RGB);
        expect(entry.textDecorationLine, `${entry.text}: decoration`).toContain('underline');
      }
    });

  /**
   * The control that makes the assertion above meaningful. `w:hyperlink` is a link, not a style: a
   * renderer that decorated every hyperlink would pass the previous test while actually ignoring
   * the style, and LibreOffice would then be the one following the file.
   */
  test('an identical entry WITHOUT the character style is not decorated', async ({ page }) => {
    const { entries } = await renderToc(page);
    const unstyledIndex = TOC_ENTRIES.findIndex(entry => !entry.styled);
    expect(unstyledIndex).toBeGreaterThanOrEqual(0);
    const unstyled = entries[unstyledIndex];

    expect(unstyled.color, 'an unstyled hyperlink run must not be painted hyperlink blue')
      .not.toBe(TOC_HYPERLINK_RGB);
    expect(unstyled.textDecorationLine,
      'an unstyled hyperlink run must not be underlined by the renderer')
      .not.toContain('underline');
  });

  test('hyperlink appearance is independent of line geometry', async ({ page }) => {
    const { entries } = await renderToc(page);
    const boxes = new Set(entries.map(entry => Math.round(entry.lineBox * 100)));
    // Styled and unstyled entries share one line box: colour and underline must not change height.
    expect(boxes.size, 'the character style must not alter the line box').toBe(1);
  });
});
