import { test, expect } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { hiddenMarkFixture } from './docx-hidden-mark-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('delivered PDF keeps the hidden-mark heading, number and body on one line', async ({ browser }, info) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(hiddenMarkFixture(), { browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    const at = (needle: string) => {
      const item = items.find(item => item.str.includes(needle));
      expect(item, `PDF text ${needle}`).toBeDefined(); return item!;
    };
    const y = at('Cedar.').transform[5];
    expect(Math.abs(at('Return the cedar').transform[5] - y)).toBeLessThan(0.5);
    expect(Math.abs(at('1.').transform[5] - y)).toBeLessThan(0.5);
    expect(at('Maple.').transform[5] - at('Keep the maple').transform[5]).toBeGreaterThan(5);
    writeFileSync(info.outputPath('hidden-mark.pdf'), result.pdf!);
  } finally { await task.destroy(); }
});
