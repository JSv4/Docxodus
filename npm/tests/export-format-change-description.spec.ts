import { test, expect } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { formatRevisionFixture } from './docx-format-revision-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('delivered markup HTML retains accurate labels while its PDF retains all text', async ({ browser, page }, info) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(formatRevisionFixture(), { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120_000 });
  await page.setContent(result.html!);
  expect(await page.locator('.rev-format-change').evaluateAll(elements => elements.map(e => e.getAttribute('title'))))
    .toEqual(['Bold added', 'Font size changed', 'Font changed']);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const text = (await (await pdf.getPage(1)).getTextContent()).items
      .filter(item => 'str' in item).map(item => item.str).join(' ');
    for (const label of ['BoldExample', 'SizeExample', 'FontExample', 'PlainExample']) expect(text).toContain(label);
    writeFileSync(info.outputPath('format-revisions.pdf'), result.pdf!);
  } finally { await task.destroy(); }
});
