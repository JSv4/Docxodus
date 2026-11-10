import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { sectionMarkDocx, line, mark, SINGLE, TWO_COLUMNS } from './docx-section-mark-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

/** Baseline y of each text item on page 1. */
async function lineTops(bytes: Uint8Array): Promise<Record<string, number>> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const content = await (await (await task.promise).getPage(1)).getTextContent();
    return Object.fromEntries(content.items.filter(item => 'str' in item && item.str.trim())
      .map(item => ['str' in item ? item.str.trim() : '', 'transform' in item ? item.transform[5] : 0]));
  } finally { await task.destroy(); }
}

// A balanced two-column section holding one line, whether its break sits on that line or on an empty
// mark after it (#1572). Word steps one line to the next section. Vector text freezes the line's box,
// and the column rebalance that follows must not slice it across both columns.
test('a one-line balanced column section keeps its line whole', async ({ browser }) => {
  const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'final' as const,
    commentProfile: 'hidden' as const, timeoutMs: 120_000 };
  const [above, below] = [line('Line above.'), line('Line below.')];
  // The vector text layer splits each line into words.
  const step = async (body: string) => {
    const tops = await lineTops((await renderDocxArtifacts(sectionMarkDocx(body), options)).pdf!);
    return tops['above.'] - tops['below.'];
  };
  const single = await step(above + below);
  expect(single).toBeGreaterThan(10);
  expect(Math.abs(await step(line('Line above.', SINGLE + TWO_COLUMNS) + below) - single)).toBeLessThan(0.3);
  expect(Math.abs(await step(above + mark(SINGLE, '', TWO_COLUMNS) + below) - single)).toBeLessThan(0.3);
});
