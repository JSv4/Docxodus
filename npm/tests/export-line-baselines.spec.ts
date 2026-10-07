import { expect, test, type Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

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
 * Carlito and Liberation Sans, which share their metrics, served through the export's font resolver
 * (fixtures/fonts), so the result does not depend on the fonts a machine has installed.
 */
const here = dirname(fileURLToPath(import.meta.url));
const fixture = readFileSync(join(here, 'fixtures', 'line-baselines.docx'));
/** The same cases with paragraph marks that carry no run properties, so each keeps the default 11 pt Calibri
 * (issue #940). Word's baselines do not depend on the mark: it gave the same numbers for a plain-mark version of
 * the fixture as for the current one, so the fixture's record applies. */
const plainMarks = readFileSync(join(here, 'fixtures', 'line-baselines-plain-marks.docx'));
const word = JSON.parse(readFileSync(join(here, 'fixtures', 'line-baselines.word.json'), 'utf8')) as {
  fixtureSha256: string;
  cases: Array<{ case: number; baselinesPt: number[] }>;
};

/** Cases in the fixture: 0-2 are 11 pt Calibri at single, 1.15 and double; 3-5 are 12 pt Arial. */
const CALIBRI = [0, 1, 2];
const ARIAL = [3, 4, 5];

/**
 * How far an exported line may sit from Word's, measured from the same font's single-spaced first baseline.
 * Word writes baselines on its own coarse grid (its single-spaced Calibri pitch reads 13.25, 13.50 and 13.52 pt
 * for one 13.43 pt line height), and before #942 Chromium also put each baseline on a whole pixel. Before the
 * #908 fix the 1.15-spaced lines read 1.5 pt low and the double-spaced ones 6.75 pt.
 */
const LINE_TOLERANCE_PT = 0.8;

/** How far the single-spaced first baseline itself may sit from Word's. Before #942 Chromium's whole-pixel
 * rounding put it one pixel (0.75 pt) above Word's for both fonts; the #942 test below holds it to 0.1 pt. */
const FIRST_LINE_TOLERANCE_PT = 1.0;

/** How far any first baseline may sit from Word's once whole-pixel rounding is taken back (issue #942). */
const WORD_BASELINE_TOLERANCE_PT = 0.1;

/** Each CASE paragraph's line spacing multiple, by case number. */
const MULTIPLE = [1, 276 / 240, 2, 1, 276 / 240, 2];

/** A font-resolver plan serving one test font for every family the export requests. */
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
 * Export the fixture with every font family served by one test font, through the export's own font resolver
 * (the export lays out in an isolated frame, so this is how fonts reach it), and read each page's baselines in
 * pt from the page top. Measured in the DOM, not from a printed PDF: Chromium's PDF output snaps text origins
 * to whole CSS pixels (0.75 pt). A fragment's box top is a fixed distance above its baseline for a given font;
 * that distance is calibrated per page with a zero-size inline-block, whose bottom edge sits on the baseline,
 * placed at the start of the first run.
 */
interface ExportedPage { baselines: number[]; heightPx: number; strutPx: number }

async function exportedBaselines(page: Page, fontFile: string | ReturnType<typeof plan>, docx = fixture): Promise<ExportedPage[]> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(async ({ bytes, fontPlan }) => {
    const api = (window as any).DocxodusStandalone;
    return api.convertWithFontResolver(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, fontPlan);
  }, { bytes: Array.from(docx), fontPlan: typeof fontFile === 'string' ? plan(fontFile) : fontFile });
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
    const paragraph = box.querySelector('p')!;
    const lineBox = {
      heightPx: paragraph.getBoundingClientRect().height,
      strutPx: Number.parseFloat(getComputedStyle(paragraph).lineHeight),
    };
    if (fragments.length === 0) return { baselines: [], ...lineBox };
    // Calibrate on the first fragment: its own baseline, from a marker in the same inline box.
    const first = fragments[0];
    const marker = document.createElement('span');
    marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0;';
    first.node.parentNode!.insertBefore(marker, first.node);
    const ascent = marker.getBoundingClientRect().bottom - first.top;
    marker.remove();
    const baselines = new Set(fragments.map((f) => Math.round((f.top + ascent - pageTop) * 0.75 * 1000) / 1000));
    return { baselines: [...baselines].sort((a, b) => a - b), ...lineBox };
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
      const origin = exported[cases[0]].baselines[0];
      expect(Math.abs(origin - wordOrigin), `single-spaced first baseline ${origin} pt vs Word ${wordOrigin} pt`)
        .toBeLessThanOrEqual(FIRST_LINE_TOLERANCE_PT);
      for (const id of cases) {
        const expected = word.cases.find((c) => c.case === id)!.baselinesPt;
        const actual = exported[id].baselines;
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
      const firsts = cases.map((id) => exported[id].baselines[0]);
      expect(Math.max(...firsts) - Math.min(...firsts), `first baselines ${firsts.join(', ')}`)
        .toBeLessThanOrEqual(LINE_TOLERANCE_PT);
    });

    test(`${font}: every first baseline lands within 0.1 pt of Word's (#942)`, async ({ page }) => {
      // Word puts a line's baseline its natural height less the font's descent below the line's top. Chromium
      // rounded the ascent and descent and floored half the leading, one pixel high for both fonts; the export
      // now takes that back. The extra height of a multiple goes below the text, so all three spacings agree.
      const exported = await exportedBaselines(page, file);
      for (const id of cases) {
        const want = word.cases.find((c) => c.case === id)!.baselinesPt[0];
        const got = exported[id].baselines[0];
        expect(Math.abs(got - want), `CASE${id} first baseline ${got} pt vs Word ${want} pt`)
          .toBeLessThanOrEqual(WORD_BASELINE_TOLERANCE_PT);
      }
    });

    test(`${font}: each line box keeps its multiplied height`, async ({ page }) => {
      // Only the glyphs move: every line is still the multiple of the paragraph's natural line height, so
      // pagination and PageMap fragments are what they were.
      const exported = await exportedBaselines(page, file);
      for (const id of cases) {
        const { baselines, heightPx, strutPx } = exported[id];
        const expected = baselines.length * MULTIPLE[id] * strutPx;
        expect(Math.abs(heightPx - expected), `CASE${id}: ${heightPx} px tall, ${baselines.length} lines of ` +
          `${MULTIPLE[id]} x ${strutPx} px`).toBeLessThanOrEqual(baselines.length / 64 + 0.01);
      }
    });
  }

  test('the printed PDF puts each first baseline on Word\'s pixel (#942)', async ({ page }) => {
    // Chromium's PDF output snaps text origins to whole CSS pixels (0.75 pt), so a sub-pixel correction only
    // shows in print by landing on the right pixel. Word's first baselines here (82.53 and 83.28 pt) sit within
    // 0.05 pt of a whole pixel; the export used to print them one pixel higher.
    await page.goto('/standalone-export-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
    const result = await page.evaluate(async ({ bytes, fontPlan }) => {
      const api = (window as any).DocxodusStandalone;
      return api.convertWithFontResolver(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, fontPlan);
    }, {
      bytes: Array.from(plainMarks),
      fontPlan: {
        ...plan('baseline-test-b.woff2'),
        byFamily: { Calibri: plan('baseline-test-a.woff2'), Arial: plan('baseline-test-b.woff2') },
      },
    });
    await page.setContent(result.html);
    await page.evaluate(() => document.fonts.ready);
    const pdf = await page.pdf({ preferCSSPageSize: true, printBackground: true });
    const printed = await getDocument({ data: new Uint8Array(pdf), verbosity: 0 }).promise;
    for (const id of [CALIBRI[0], ARIAL[0]]) {
      const pdfPage = await printed.getPage(id + 1);
      const height = pdfPage.view[3];
      const origins = (await pdfPage.getTextContent()).items
        .filter((item: any) => item.str?.trim())
        .map((item: any) => height - item.transform[5]);
      const got = Math.min(...origins);
      const want = word.cases.find((c) => c.case === id)!.baselinesPt[0];
      expect(Math.abs(got - want), `CASE${id} printed first baseline ${got} pt vs Word ${want} pt`)
        .toBeLessThanOrEqual(0.375);
    }
  });

  test('the font resolver gives a paragraph its configured face, not only its runs', async ({ page }) => {
    // A paragraph holds no text of its own, but its own line box (the strut every auto multiple is built on)
    // is sized from its font. The resolver used to restyle only elements holding text, so the paragraph kept
    // the requested family and fell back to whatever the machine had installed.
    await page.goto('/standalone-export-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
    const result = await page.evaluate(async ({ bytes, fontPlan }) => {
      const api = (window as any).DocxodusStandalone;
      return api.convertWithFontResolver(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, fontPlan);
    }, { bytes: Array.from(fixture), fontPlan: plan('baseline-test-a.woff2') });
    await page.setContent(result.html);
    const families = await page.evaluate(() => Array.from(document.querySelectorAll('.page-box p'))
      .map((p) => getComputedStyle(p).fontFamily));
    expect(families.length).toBeGreaterThan(0);
    for (const family of families) expect(family).toMatch(/^__DocxodusConfigured_/);
  });

  test('Arial runs under a plain Calibri mark keep their own line pitch (#940)', async ({ page }) => {
    // The paragraph's own line box used to pair the mark's Calibri with the runs' 12 pt: a font-and-size
    // combination found nowhere in the document, taller than any line (14.65 pt for Word's 13.84). Each family
    // gets its own face here, so the mark's family really differs from the runs'.
    const twoFaces = {
      ...plan('baseline-test-b.woff2'),
      byFamily: { Calibri: plan('baseline-test-a.woff2'), Arial: plan('baseline-test-b.woff2') },
    };
    const exported = await exportedBaselines(page, twoFaces, plainMarks);
    const wordOrigin = word.cases.find((c) => c.case === ARIAL[0])!.baselinesPt[0];
    const origin = exported[ARIAL[0]].baselines[0];
    for (const id of ARIAL) {
      const expected = word.cases.find((c) => c.case === id)!.baselinesPt;
      const actual = exported[id].baselines;
      expected.forEach((y, line) => {
        const got = actual[line] - origin;
        expect(Math.abs(got - (y - wordOrigin)), `CASE${id} line ${line + 1}: ${got} pt below the first ` +
          `baseline, Word ${y - wordOrigin} pt`).toBeLessThanOrEqual(LINE_TOLERANCE_PT);
      });
    }
  });

  test('a raised run grows its line the way Word does (#941)', async ({ page }) => {
    // CASE6 is CASE1 with "raised" at w:position 6 (3 pt up) on its first line. Word keeps the raised text at
    // the line's usual top and moves the line's baseline down by the raise, and every later line with it.
    const exported = await exportedBaselines(page, 'baseline-test-a.woff2');
    const origin = exported[1].baselines[0];
    const wordOrigin = word.cases.find((c) => c.case === 1)!.baselinesPt[0];
    const expected = word.cases.find((c) => c.case === 6)!.baselinesPt;
    const actual = exported[6].baselines;
    expected.forEach((y, line) => {
      const got = actual[line] - origin;
      expect(Math.abs(got - (y - wordOrigin)), `CASE6 baseline ${line + 1}: ${got} pt below CASE1's first, ` +
        `Word ${y - wordOrigin} pt`).toBeLessThanOrEqual(LINE_TOLERANCE_PT);
    });
  });

  test('a lowered run grows its line downward the way Word does (#948)', async ({ page }) => {
    // CASE7 is CASE1 with "lowered" at w:position -6 (3 pt down) on its first line. Word keeps the line's
    // baseline where it was and grows the line below by the drop, so every later line moves down 3 pt. A
    // relative offset moved the glyphs without growing the line, leaving the later lines where CASE1 has them.
    const exported = await exportedBaselines(page, 'baseline-test-a.woff2');
    const origin = exported[1].baselines[0];
    const wordOrigin = word.cases.find((c) => c.case === 1)!.baselinesPt[0];
    const expected = word.cases.find((c) => c.case === 7)!.baselinesPt;
    const actual = exported[7].baselines;
    expect(actual.length, 'CASE7 baselines').toBeGreaterThanOrEqual(expected.length);
    expected.forEach((y, line) => {
      const got = actual[line] - origin;
      expect(Math.abs(got - (y - wordOrigin)), `CASE7 baseline ${line + 1}: ${got} pt below CASE1's first, ` +
        `Word ${y - wordOrigin} pt`).toBeLessThanOrEqual(LINE_TOLERANCE_PT);
    });
  });

  test('a lowered run at 1.15 sits its w:position below the line', async ({ page }) => {
    // Word draws "lowered" 3 pt below its line's baseline.
    const [line, lowered] = (await exportedBaselines(page, 'baseline-test-a.woff2'))[7].baselines;
    const [wordLine, wordLowered] = word.cases.find((c) => c.case === 7)!.baselinesPt;
    expect(Math.abs((lowered - line) - (wordLowered - wordLine)), `line ${line}, lowered ${lowered}`)
      .toBeLessThanOrEqual(0.05);
  });

  test('a raised run at 1.15 still sits its w:position above the line', async ({ page }) => {
    // CASE6 is 11 pt Calibri at 1.15 whose first line holds "raised" at w:position 6 (3 pt up). Word draws it
    // 3 pt above the line's baseline; the relative offset that raises it composes with the new placement.
    const [raised, line] = (await exportedBaselines(page, 'baseline-test-a.woff2'))[6].baselines;
    const [wordRaised, wordLine] = word.cases.find((c) => c.case === 6)!.baselinesPt;
    expect(Math.abs((line - raised) - (wordLine - wordRaised)), `raised ${raised}, line ${line}`)
      .toBeLessThanOrEqual(0.05);
  });
});
