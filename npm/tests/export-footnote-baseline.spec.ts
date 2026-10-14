import { expect, test, type Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { FOOTNOTE_PAGE, generateFootnoteDocx, twipsToPt } from './docx-footnote-fixture.js';

/**
 * The exported footnote text sits where Word puts it (issue #955).
 *
 * Word anchors the footnote area to the bottom of the text column, and a note's lines are its
 * FootnoteText paragraph's own lines, so the last note line's baseline is the bottom margin less the
 * note font's descent. The export laid the note out inside wrappers set in the page's default font,
 * whose strut made every note line taller than the paragraph's own and lifted the text about a pixel
 * above Word's.
 *
 * The note font is the repository's test font, served through the export's font resolver for every
 * requested family, so the expectation comes from its `hhea` table, not from the code under test:
 * descent 483 of 2048 units per em, and no line gap.
 */
const fontBytes = readFileSync(join(dirname(fileURLToPath(import.meta.url)),
  '..', '..', 'docs', 'demo', 'fonts', 'docxodus-canvas-mono.woff2'));
const fontPlan = {
  mode: 'exact',
  format: 'woff2',
  mediaType: 'font/woff2',
  byteLength: fontBytes.byteLength,
  sha256: createHash('sha256').update(fontBytes).digest('hex'),
  bytesBase64: fontBytes.toString('base64'),
  licenseIdentity: 'a'.repeat(64),
};
const DESCENT_EM = 483 / 2048;
const NOTE_SIZE_PT = 10;
/**
 * Half a CSS pixel. Chromium rounds a font's ascent and descent to whole pixels per inline box, which
 * leaves this font's note baseline 0.2 pt off the unrounded model; the defect put it 1.36 pt high.
 */
const TOLERANCE_PT = 0.375;

/** The baseline of the last line of the last note on page 1, in pt from the page top. */
async function lastNoteBaseline(page: Page, source: Uint8Array): Promise<number> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(async ({ bytes, plan }) =>
    (window as any).DocxodusStandalone.convertWithFontResolver(
      bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, plan),
  { bytes: Array.from(source), plan: fontPlan });
  expect(result.renderReport.status).toBe('complete');
  await page.setContent(result.html);
  return page.evaluate(() => {
    const pageBox = document.querySelector('.page-box') as HTMLElement;
    const notes = Array.from(pageBox.querySelectorAll('.page-footnotes p'));
    const last = notes[notes.length - 1];
    // A zero-height inline-block sits on the baseline of the line it ends.
    const probe = document.createElement('span');
    probe.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
    (last.lastElementChild ?? last).appendChild(probe);
    const baseline = probe.getBoundingClientRect().top - pageBox.getBoundingClientRect().top;
    probe.remove();
    return baseline * 0.75;
  });
}

test.describe('exported footnote baseline (#955)', () => {
  const marginBottomPt = twipsToPt(FOOTNOTE_PAGE.heightTwips - FOOTNOTE_PAGE.marginTwips);
  const expected = marginBottomPt - NOTE_SIZE_PT * DESCENT_EM;

  test('a one-line note ends on the bottom margin, less its descent', async ({ page }) => {
    const baseline = await lastNoteBaseline(page, generateFootnoteDocx(1));
    expect(Math.abs(baseline - expected), `${baseline} pt vs ${expected} pt`).toBeLessThanOrEqual(TOLERANCE_PT);
  });

  test('so does the last line of a note that wraps', async ({ page }) => {
    const baseline = await lastNoteBaseline(page, generateFootnoteDocx(1, 5, 1, 40));
    expect(Math.abs(baseline - expected), `${baseline} pt vs ${expected} pt`).toBeLessThanOrEqual(TOLERANCE_PT);
  });
});
