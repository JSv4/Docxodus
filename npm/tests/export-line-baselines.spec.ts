import { expect, test, type Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Exported baselines sit where Word puts them (issue #908). For `w:lineRule="auto"` multiples Word places a
 * line's baseline where single spacing would, and adds the multiple's extra height below the text; CSS splits
 * extra leading half above and half below the glyphs, so 1.15 and double spacing used to sit every line lower
 * than Word by half the extra.
 *
 * The expected values are Word's own, recorded from a Word-exported PDF of the committed fixture
 * (fixtures/line-baselines.word.json, captured with fixtures/capture-line-baselines.mjs). Each fixture page
 * holds one CASE paragraph starting at the top margin. Word's baselines are its PDF's text origins. The
 * fixture's Calibri and Arial are laid out with subsets of
 * Carlito and Liberation Sans, which share their metrics, served through the export's own font resolver
 * (fixtures/fonts), so the result does not depend on the fonts a machine has installed.
 */
const here = dirname(fileURLToPath(import.meta.url));
const fixture = readFileSync(join(here, 'fixtures', 'line-baselines.docx'));
const word = JSON.parse(readFileSync(join(here, 'fixtures', 'line-baselines.word.json'), 'utf8')) as {
  fixtureSha256: string;
  cases: Array<{ case: number; baselinesPt: number[] }>;
};

/** Cases in the fixture: 0-2 are 11 pt Calibri at single, 1.15 and double; 3-5 are 12 pt Arial. */
const CALIBRI = [0, 1, 2];
const ARIAL = [3, 4, 5];

/** How far an exported line may sit from Word's, measured from the same paragraph font's single-spaced first
 * baseline. Word writes baselines on a coarser grid than Chromium lays them out (its single-spaced Calibri
 * pitch reads 13.25, 13.50 and 13.52 pt for one 13.43 pt line height), so one line can read up to ~0.15 pt
 * apart with the model right. Before the fix, 1.15 spacing put every line 1.0 pt low and double 6.7 pt. */
const LINE_TOLERANCE_PT = 0.3;

/** How far the single-spaced first baseline itself may sit from Word's. Chromium puts a line's baseline on a
 * whole CSS pixel, and the exported first line sits one pixel (0.75 pt) above Word's for both fonts whatever
 * the spacing (issue #942). That offset is shared by every line, so the comparison above measures from it. */
const FIRST_LINE_TOLERANCE_PT = 1.0;

function plan(file: string) {
  const bytes = readFileSync(join(here, 'fixtures', 'fonts', file));
  return {
    mode: 'exact',
    format: 'woff2',
    mediaType: 'font/woff2',
    byteLength: bytes.byteLength,
    sha256: createHash('sha256').update(bytes).digest('hex'),
    bytesBase64: bytes.toString('base64'),
    licenseIdentity: 'b'.repeat(64),
  };
}

/**
 * Export the fixture with every font family served by one test font and read each page's baselines, in pt
 * from the page top. Measured in the DOM, not from a printed PDF: Chromium's PDF output snaps text origins
 * to whole CSS pixels (0.75 pt). A fragment's box top is a fixed distance above its baseline for a given
 * font; that distance is calibrated per paragraph with a zero-size inline-block, whose bottom edge sits on
 * the baseline, placed at the start of the paragraph's first run.
 */
async function exportedBaselines(page: Page, fontFile: string): Promise<number[][]> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(async ({ bytes, fontPlan }) => {
    const api = (window as any).DocxodusStandalone;
    return api.convertWithFontResolver(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, fontPlan);
  }, { bytes: Array.from(fixture), fontPlan: plan(fontFile) });
  expect(result.renderReport.status).toBe('complete');
  await page.setContent(result.html);
  await page.evaluate(() => document.fonts.ready);
  return page.evaluate(() => Array.from(document.querySelectorAll('.page-box')).map((box) => {
    const pageTop = box.getBoundingClientRect().top;
    const fragments: Array<{ top: number; node: Text }> = [];
    const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
      if (!(node.textContent ?? '').trim()) continue;
      const range = document.createRange();
      range.selectNodeContents(node);
      for (const rect of Array.from(range.getClientRects()))
        if (rect.width > 0) fragments.push({ top: rect.top, node });
    }
    if (fragments.length === 0) return [];
    // Calibrate on the first fragment: its own baseline, from a marker in the same inline box.
    const first = fragments[0];
    const marker = document.createElement('span');
    marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0;';
    first.node.parentNode!.insertBefore(marker, first.node);
    const ascent = marker.getBoundingClientRect().bottom - first.top;
    marker.remove();
    const baselines = new Set(fragments.map((f) => Math.round((f.top + ascent - pageTop) * 0.75 * 1000) / 1000));
    return [...baselines].sort((a, b) => a - b);
  }));
}

test.describe('exported baselines follow Word (#908)', () => {
  test('the fixture is the one Word was measured on', () => {
    expect(createHash('sha256').update(fixture).digest('hex')).toBe(word.fixtureSha256);
  });

  for (const [font, file, cases] of [
    ['Calibri', 'baseline-test-a.woff2', CALIBRI],
    ['Arial', 'baseline-test-b.woff2', ARIAL],
  ] as const) {
    test(`${font} at single, 1.15 and double: every line where Word puts it`, async ({ page }) => {
      const exported = await exportedBaselines(page, file);
      const wordOrigin = word.cases.find((c) => c.case === cases[0])!.baselinesPt[0];
      const origin = exported[cases[0]][0];
      expect(Math.abs(origin - wordOrigin), `single-spaced first baseline ${origin} pt vs Word ${wordOrigin} pt`)
        .toBeLessThanOrEqual(FIRST_LINE_TOLERANCE_PT);
      for (const id of cases) {
        const expected = word.cases.find((c) => c.case === id)!.baselinesPt;
        const actual = exported[id];
        expect(actual.length, `CASE${id} line count`).toBeGreaterThanOrEqual(expected.length);
        expected.forEach((y, line) => {
          const want = y - wordOrigin;
          const got = actual[line] - origin;
          expect(Math.abs(got - want), `CASE${id} line ${line + 1}: ${got} pt below the single-spaced first ` +
            `baseline, Word ${want} pt`).toBeLessThanOrEqual(LINE_TOLERANCE_PT);
        });
      }
    });

    test(`${font}: the first baseline does not move with the multiple`, async ({ page }) => {
      // Word's own record has the same first baseline at single, 1.15 and double. Half-leading put the
      // double-spaced first line half a line lower than the single-spaced one.
      const exported = await exportedBaselines(page, file);
      const firsts = cases.map((id) => exported[id][0]);
      expect(Math.max(...firsts) - Math.min(...firsts), `first baselines ${firsts.join(', ')}`).toBeLessThan(0.05);
    });
  }

  test('a raised run at 1.15 still sits its w:position above the line', async ({ page }) => {
    // CASE6 is 11 pt Calibri at 1.15 whose first line holds "raised" at w:position 6 (3 pt up). Word draws it
    // 3 pt above the line's baseline; the relative offset that raises it composes with the new placement.
    const [raised, line] = (await exportedBaselines(page, 'baseline-test-a.woff2'))[6];
    const [wordRaised, wordLine] = word.cases.find((c) => c.case === 6)!.baselinesPt;
    expect(Math.abs((line - raised) - (wordLine - wordRaised)), `raised ${raised}, line ${line}`)
      .toBeLessThanOrEqual(0.05);
  });
});
