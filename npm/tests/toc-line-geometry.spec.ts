import { expect, test } from '@playwright/test';
import {
  TOC_AFTER_TWIPS,
  TOC_ENTRIES,
  TOC_HYPERLINK_RGB,
  TOC_LINE_TWIPS,
  ORDINARY_HYPERLINK_TEXT,
  generateTocDocx,
} from './docx-toc-fixture.js';

/**
 * Issues #397/#427 — TOC entry line geometry and field-result appearance in print layout.
 *
 * The issue named two things, and they turned out to have different answers, so they are pinned
 * SEPARATELY here and neither can hide behind the other:
 *
 *  - **Line-box height was a renderer bug.** TOC entries drifted further down the page with every
 *    entry because automatic line spacing was measured against font-size instead of the font's own
 *    line box (issue #396). Fixed; on the tracked HC022 fixture entries now land within 0.12px of
 *    LibreOffice, where they were off by up to 14.7px.
 *  - **Cached TOC hyperlink appearance is field semantics.** Microsoft Word renders the result
 *    entries black despite their `Hyperlink` character style, while an ordinary hyperlink using
 *    the same style remains blue and underlined. The outer TOC field is the deciding context.
 *
 * The generated fixture uses the same `TOC1` paragraph style and `Hyperlink` character style for
 * its ordinary-link control, so neither a CSS selector nor a style-name heuristic can pass.
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
      .filter(p => (p.textContent || '').includes('heading of the generated document'))
      .map(p => {
        const rect = p.getBoundingClientRect();
        // The entry's own text run — NOT the leader/page-number runs, whose appearance is a
        // separate question. Runs nest (the converter wraps them), and querySelectorAll is in
        // document order, so the LAST span carrying the entry text is the innermost one: the run
        // that actually holds the character style.
        const carriers = Array.from(p.querySelectorAll('span'))
          .filter(span => (span.textContent || '').includes('generated document'));
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

    const ordinaryParagraph = paragraphs.find(p =>
      (p.textContent || '').includes('ordinary hyperlink')) as HTMLElement;
    const ordinaryCarriers = Array.from(ordinaryParagraph.querySelectorAll('span'))
      .filter(span => (span.textContent || '').includes('ordinary hyperlink'));
    const ordinaryRun = ordinaryCarriers[ordinaryCarriers.length - 1] as HTMLElement;
    const ordinaryStyle = getComputedStyle(ordinaryRun);

    return {
      entries: entries as EntryGeometry[],
      nativeLineBox,
      ordinary: {
        text: ordinaryParagraph.textContent || '',
        color: ordinaryStyle.color,
        textDecorationLine: ordinaryStyle.textDecorationLine,
      },
    };
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
  test('cached TOC result links suppress the Hyperlink character style presentation',
    async ({ page }) => {
      const { entries } = await renderToc(page);
      for (const entry of entries) {
        expect(entry.color, `${entry.text}: color`).toBe('rgb(0, 0, 0)');
        expect(entry.textDecorationLine, `${entry.text}: decoration`).not.toContain('underline');
      }
    });

  test('an ordinary link using the same paragraph and character styles remains decorated',
    async ({ page }) => {
      const { ordinary } = await renderToc(page);
      expect(ordinary.text).toBe(ORDINARY_HYPERLINK_TEXT);
      expect(ordinary.color).toBe(TOC_HYPERLINK_RGB);
      expect(ordinary.textDecorationLine).toContain('underline');
    });

  test('hyperlink appearance is independent of line geometry', async ({ page }) => {
    const { entries } = await renderToc(page);
    const boxes = new Set(entries.map(entry => Math.round(entry.lineBox * 100)));
    // Field-suppressed and ordinary presentation share one line box: appearance does not change
    // geometry.
    expect(boxes.size, 'the character style must not alter the line box').toBe(1);
  });
});
