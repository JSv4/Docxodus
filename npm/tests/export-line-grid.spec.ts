import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { lineGridDocx } from './docx-line-grid-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const entry of [
  { pitch: 18, step: 18, lines: 36, pages: 4 },
  { pitch: 12, step: 24, lines: 27, pages: 5 },
  { pitch: 18, snap: false, step: 2355 / 2048 * 12, lines: 46, pages: 3 },
  { pitch: 0, step: 2355 / 2048 * 12, lines: 46, pages: 3 },
  // Word 2024 (#1551): an at-least line is max(whole grid lines, minimum).
  { pitch: 18, spacing: '<w:spacing w:line="300" w:lineRule="atLeast"/>', step: 18, lines: 36, pages: 4 },
  { pitch: 0, spacing: '<w:spacing w:line="300" w:lineRule="atLeast"/>', step: 15, lines: 43, pages: 3 },
  // Above the pitch too: 24pt text with 30pt at least takes two 18pt grid lines.
  { pitch: 18, size: 24, spacing: '<w:spacing w:line="600" w:lineRule="atLeast"/>', step: 36, lines: 18, pages: 7 },
  { pitch: 0, size: 24, spacing: '<w:spacing w:line="600" w:lineRule="atLeast"/>', step: 30, lines: 21, pages: 6 },
  // A revision-bearing document keeps its grid in markup, and one line before is one pitch (12pt without a grid).
  { pitch: 18, spacing: '<w:spacing w:beforeLines="100"/>', tracked: true, step: 36, lines: 18, pages: 7 },
  { pitch: 0, spacing: '<w:spacing w:beforeLines="100"/>', tracked: true, step: 2355 / 2048 * 12 + 12, lines: 25, pages: 5 },
  // So does a paragraph that shows a tracked deletion.
  { pitch: 18, tracked: 'del' as const, step: 18, lines: 36, pages: 4 },
  { pitch: 0, tracked: 'del' as const, step: 2355 / 2048 * 12, lines: 46, pages: 3 },
]) test(`single spacing PDF grid=${entry.pitch}, snap=${entry.snap !== false} ${entry.size ?? 12}pt ${entry.spacing ?? ''}${entry.tracked ? ` markup ${entry.tracked}` : ''}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(lineGridDocx(entry.pitch, entry.snap, entry.spacing, entry.tracked, entry.size), { browser,
    outputs: ['pdf'], reviewProfile: entry.tracked ? 'markup' : 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(entry.pages);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const labels: string[] = [];
    for (let number = 1; number <= pdf.numPages; number++) {
      const content = await (await pdf.getPage(number)).getTextContent();
      const text = content.items.filter(item => 'str' in item).filter(item => item.str.startsWith('Cedar'));
      // A deletion shown after the label may share its text item.
      labels.push(...text.map(item => item.str.slice(0, 8)));
      if (number === 1) {
        expect(text).toHaveLength(entry.lines);
        expect((text[0].transform[5] - text.at(-1)!.transform[5]) / (text.length - 1)).toBeCloseTo(entry.step, 2);
      }
    }
    expect(labels).toEqual(Array.from({ length: 120 }, (_, i) => `Cedar${String(i + 1).padStart(3, '0')}`));
  } finally { await task.destroy(); }
});
