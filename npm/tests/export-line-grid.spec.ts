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
]) test(`single spacing PDF grid=${entry.pitch}, snap=${entry.snap !== false}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(lineGridDocx(entry.pitch, entry.snap), { browser, outputs: ['pdf'],
    reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  expect(result.pageCount).toBe(entry.pages);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const labels: string[] = [];
    for (let number = 1; number <= pdf.numPages; number++) {
      const content = await (await pdf.getPage(number)).getTextContent();
      const text = content.items.filter(item => 'str' in item).filter(item => item.str.startsWith('Cedar'));
      labels.push(...text.map(item => item.str));
      if (number === 1) {
        expect(text).toHaveLength(entry.lines);
        expect((text[0].transform[5] - text.at(-1)!.transform[5]) / (text.length - 1)).toBeCloseTo(entry.step, 2);
      }
    }
    expect(labels).toEqual(Array.from({ length: 120 }, (_, i) => `Cedar${String(i + 1).padStart(3, '0')}`));
  } finally { await task.destroy(); }
});
