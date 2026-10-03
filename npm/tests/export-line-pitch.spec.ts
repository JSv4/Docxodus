import { expect, test, type Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { storedZip, W_NS, xml } from './docx-zip.js';

/**
 * Exported line pitch follows the font's natural line height, not a whole-pixel rounding of it
 * (issue #850). Chromium sizes a `line-height: normal` line from the font's ascent + descent + line
 * gap but rounds the result to a CSS pixel, so 11 pt Carlito advanced 18 px instead of 17.904 px and
 * the error accumulated down every page.
 *
 * The metric cases load the repository's own test font through the export's font resolver, so they
 * do not depend on what a machine has installed (installing fonts system-wide on CI changes other
 * screenshot tests). The expected height comes from that font's `hhea` table, the metrics Chromium
 * uses on Linux, not from the code under test: Docxodus Canvas Mono has ascender 1901 +
 * |descender| 483 + lineGap 0 = 2384 per 2048 units/em.
 */
const TEST_FONT = 'Docxodus Canvas Mono';
const TEST_FONT_NATURAL_EM = 2384 / 2048;
const testFontBytes = readFileSync(join(dirname(fileURLToPath(import.meta.url)),
  '..', '..', 'docs', 'demo', 'fonts', 'docxodus-canvas-mono.woff2'));
/** Serve the test font for every requested family, through the harness's configured resolver. */
const testFontPlan = {
  mode: 'exact',
  format: 'woff2',
  mediaType: 'font/woff2',
  byteLength: testFontBytes.byteLength,
  sha256: createHash('sha256').update(testFontBytes).digest('hex'),
  bytesBase64: testFontBytes.toString('base64'),
  licenseIdentity: 'a'.repeat(64),
};
/** Each line may differ from the model by at most this much (the issue's ~0.05 pt target). */
const TOLERANCE_PT_PER_LINE = 0.05;

const SENTENCE = 'The quick brown fox jumps over the lazy dog while the editor counts every line. ';

const CJK = '漢字テキストの行の高さを測ります。';

function docx(font: string, halfPoints: number, lines: number[],
  rule = 'auto', text = SENTENCE): Uint8Array {
  const paragraphs = lines.map((line, index) =>
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="${rule}"/></w:pPr>`
    + `<w:r><w:t xml:space="preserve">CASE${index} ${text.repeat(8)}</w:t></w:r></w:p>`).join('');
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
  <w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:eastAsia="${font}"/>
  <w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>
</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>${paragraphs}
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

/**
 * Mean distance between consecutive line tops of each CASE paragraph in the exported HTML, in pt.
 * With `normalReference`, each paragraph is also cloned with every line height forced back to
 * `normal`, and that clone's pitch is returned alongside: what Chromium's own model gives the text.
 */
async function exportedPitches(page: Page, source: Uint8Array, withTestFont?: boolean): Promise<number[]>;
async function exportedPitches(page: Page, source: Uint8Array, withTestFont: boolean,
  normalReference: true): Promise<Array<[number, number]>>;
async function exportedPitches(page: Page, source: Uint8Array, withTestFont = false,
  normalReference = false): Promise<unknown[]> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(async ({ bytes, plan }) => {
    const api = (window as any).DocxodusStandalone;
    const options = { reviewProfile: 'final', commentProfile: 'hidden' };
    return plan ? api.convertWithFontResolver(bytes, options, plan) : api.convert(bytes, options);
  }, { bytes: Array.from(source), plan: withTestFont ? testFontPlan : undefined });
  expect(result.renderReport.status).toBe('complete');
  await page.setContent(result.html);
  return page.evaluate((withReference) => {
    const pitchOf = (element: Element) => {
      const range = document.createRange();
      range.selectNodeContents(element);
      const tops = [...new Set(Array.from(range.getClientRects())
        .filter((rect) => rect.width > 0)
        .map((rect) => Math.round(rect.top * 64) / 64))].sort((a, b) => a - b);
      return ((tops[tops.length - 1] - tops[0]) / (tops.length - 1)) * 0.75;
    };
    const pitches: unknown[] = [];
    for (let index = 0; ; index++) {
      const paragraph = Array.from(document.querySelectorAll('p'))
        .find((p) => (p.textContent ?? '').trimStart().startsWith(`CASE${index} `));
      if (!paragraph) break;
      if (!withReference) {
        pitches.push(pitchOf(paragraph));
        continue;
      }
      const clone = paragraph.cloneNode(true) as HTMLElement;
      for (const element of [clone, ...Array.from(clone.querySelectorAll<HTMLElement>('*'))]) {
        element.style.setProperty('line-height', 'normal');
      }
      clone.style.width = `${paragraph.getBoundingClientRect().width}px`;
      paragraph.after(clone);
      pitches.push([pitchOf(paragraph), pitchOf(clone)]);
      clone.remove();
    }
    return pitches;
  }, normalReference);
}

test.describe('exported line pitch (#850)', () => {
  test('12pt single, 1.08, 1.15 and double follow the font metrics', async ({ page }) => {
    // 12 pt is 18.625 px natural against 19 px rounded: 0.28 pt per line before the fix.
    const lines = [240, 259, 276, 480];
    const pitches = await exportedPitches(page, docx(TEST_FONT, 24, lines), true);

    expect(pitches).toHaveLength(lines.length);
    lines.forEach((line, index) => {
      const expected = TEST_FONT_NATURAL_EM * 12 * (line / 240);
      expect(Math.abs(pitches[index] - expected), `w:line=${line}: ${pitches[index]} pt vs ${expected} pt`)
        .toBeLessThanOrEqual(TOLERANCE_PT_PER_LINE);
    });
  });

  test('14pt single and double follow the font metrics', async ({ page }) => {
    // 14 pt is 21.729 px natural against 22 px rounded: 0.20 pt per line before the fix.
    const lines = [240, 480];
    const pitches = await exportedPitches(page, docx(TEST_FONT, 28, lines), true);

    expect(pitches).toHaveLength(lines.length);
    lines.forEach((line, index) => {
      const expected = TEST_FONT_NATURAL_EM * 14 * (line / 240);
      expect(Math.abs(pitches[index] - expected), `w:line=${line}: ${pitches[index]} pt vs ${expected} pt`)
        .toBeLessThanOrEqual(TOLERANCE_PT_PER_LINE);
    });
  });

  test('exact spacing keeps its twentieths of a point', async ({ page }) => {
    // w:line="253" is 12.65 pt; a one-decimal CSS value made it 12.7 pt.
    const [pitch] = await exportedPitches(page, docx('Carlito', 22, [253], 'exact'));

    expect(Math.abs(pitch - 12.65), `${pitch} pt`).toBeLessThanOrEqual(0.02);
  });

  test('a line of CJK text keeps the height its fallback font gives it', async ({ page }) => {
    // Under normal, Chromium grows a line for a taller fallback font. An explicit height taken from
    // the primary font alone would squeeze it; the pitch must stay within a pixel of normal's.
    const [[exported, normal]] = await exportedPitches(page, docx('Carlito', 22, [240], 'auto', CJK), false, true);

    expect(Math.abs(exported - normal), `exported ${exported} pt vs normal ${normal} pt`)
      .toBeLessThan(0.75);
  });
});
