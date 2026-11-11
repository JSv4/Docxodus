import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { lineGridDocx } from './docx-line-grid-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const atLeast = (points: number) => `<w:spacing w:line="${points * 20}" w:lineRule="atLeast"/>`;

// `baseline` is Word 2024's first baseline on page 1 (pt from the top), where measured. It centres a line's text
// in its grid lines (#1551). HTML paint rounds a baseline to a whole pixel, so half a point is allowed, and so is
// a step to one decimal (`digits`) where a line is not a whole number of pixels.
for (const entry of [
  { pitch: 18, step: 18, lines: 36, pages: 4, baseline: 85.34 },
  { pitch: 12, step: 24, lines: 27, pages: 5, baseline: 88.34 },
  { pitch: 18, snap: false, step: 2355 / 2048 * 12, lines: 46, pages: 3, baseline: 83.18 },
  { pitch: 0, step: 2355 / 2048 * 12, lines: 46, pages: 3, baseline: 83.18 },
  // Word 2024 (#1551): an at-least line is max(whole grid lines, minimum).
  { pitch: 18, spacing: atLeast(15), step: 18, lines: 36, pages: 4, baseline: 85.34 },
  { pitch: 0, spacing: atLeast(15), step: 15, lines: 43, pages: 3, baseline: 84.38 },
  // Above the pitch too: 24pt text with 30pt at least takes two 18pt grid lines.
  { pitch: 18, size: 24, spacing: atLeast(30), step: 36, lines: 18, pages: 7, baseline: 98.66 },
  { pitch: 0, size: 24, spacing: atLeast(30), step: 30, lines: 21, pages: 6, baseline: 96.86 },
  // A minimum that wins puts its surplus over the grid lines above them, and the text is centred in those.
  { pitch: 18, spacing: atLeast(20), step: 20, lines: 32, pages: 4, baseline: 87.26, digits: 1 },
  { pitch: 0, spacing: atLeast(20), step: 20, lines: 32, pages: 4, baseline: 89.42, digits: 1 },
  // A revision-bearing document keeps its grid in markup, and one line before is one pitch (12pt without a grid).
  { pitch: 18, spacing: '<w:spacing w:beforeLines="100"/>', tracked: true, step: 36, lines: 18, pages: 7 },
  { pitch: 0, spacing: '<w:spacing w:beforeLines="100"/>', tracked: true, step: 2355 / 2048 * 12 + 12, lines: 25, pages: 5 },
  // So does a paragraph that shows a tracked deletion.
  { pitch: 18, tracked: 'del' as const, step: 18, lines: 36, pages: 4 },
  { pitch: 0, tracked: 'del' as const, step: 2355 / 2048 * 12, lines: 46, pages: 3 },
  // And one whose runs differ in weight or size: each line takes the grid lines its tallest text needs.
  { pitch: 18, word: '<w:b/>', step: 18, lines: 36, pages: 4, baseline: 85.34 },
  { pitch: 0, word: '<w:b/>', step: 2355 / 2048 * 12, lines: 46, pages: 3, baseline: 83.18, digits: 1 },
  { pitch: 18, word: '<w:sz w:val="48"/>', step: 36, lines: 18, pages: 7, baseline: 98.66 },
  { pitch: 0, word: '<w:sz w:val="48"/>', step: 2355 / 2048 * 24, lines: 23, pages: 6, baseline: 94.46, digits: 1 },
  // Off the grid, a bold word's line puts its at-least surplus above the text, as a uniform line does.
  { pitch: 0, word: '<w:b/>', spacing: atLeast(20), step: 20, lines: 32, pages: 4, baseline: 89.42, digits: 1 },
  // A page's last line needs only its centred text: 18 such lines fit a 645pt body (17 x 36pt + 31.8pt), 17 a 643pt one.
  { pitch: 18, word: '<w:sz w:val="48"/>', bottom: 1500, step: 36, lines: 18, pages: 7 },
  { pitch: 18, word: '<w:sz w:val="48"/>', bottom: 1540, step: 36, lines: 17, pages: 8 },
]) test(`single spacing PDF grid=${entry.pitch}, snap=${entry.snap !== false} ${entry.size ?? 12}pt ${entry.spacing ?? ''}${entry.tracked ? ` markup ${entry.tracked}` : ''}${entry.word ? ` word ${entry.word}` : ''}${entry.bottom ? ` bottom ${entry.bottom}` : ''}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(lineGridDocx(entry.pitch, entry.snap, entry.spacing, entry.tracked, entry.size, entry.word, entry.bottom),
    { browser, outputs: ['pdf'], reviewProfile: entry.tracked ? 'markup' : 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(entry.pages);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const labels: string[] = [];
    for (let number = 1; number <= pdf.numPages; number++) {
      const page = await pdf.getPage(number);
      const text = (await page.getTextContent()).items.filter(item => 'str' in item).filter(item => item.str.startsWith('Cedar'));
      // A deletion or a word shown after the label may share its text item.
      labels.push(...text.map(item => item.str.slice(0, 8)));
      if (number === 1) {
        expect(text).toHaveLength(entry.lines);
        expect((text[0].transform[5] - text.at(-1)!.transform[5]) / (text.length - 1)).toBeCloseTo(entry.step, entry.digits ?? 2);
        if (entry.baseline) expect(Math.abs(page.view[3] - text[0].transform[5] - entry.baseline)).toBeLessThan(.5);
      }
    }
    expect(labels).toEqual(Array.from({ length: 120 }, (_, i) => `Cedar${String(i + 1).padStart(3, '0')}`));
  } finally { await task.destroy(); }
});
