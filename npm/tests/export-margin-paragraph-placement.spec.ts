import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { marginParagraphPlacementDocx, placementPins } from './docx-margin-paragraph-placement-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const mode of ['outside', 'visible', 'straddling', 'internal-overflow', 'unknown-basis'] as const) {
  test(`margin/paragraph placement preserves the ${mode} boundary`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = marginParagraphPlacementDocx(mode);
    writeFileSync(testInfo.outputPath('mica-placement.docx'), bytes);
    const operation = renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    if (mode !== 'outside' && mode !== 'visible') {
      await expect(operation).rejects.toThrow(/has no measurable fragment/);
      return;
    }
    const result = await operation;
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.warnings.filter(w => w.code === 'source_text_clipped')).toHaveLength(mode === 'outside' ? 2 : 0);
    expect(result.html).toContain('data-docx-textbox-source-position="true"');
    expect(result.html).not.toContain('data-docx-textbox-source-clip="true"');
    for (const pin of placementPins) expect(result.html).toContain(pin);
    const pdf = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await pdf.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const text = items.map(item => item.str).join(' ');
      expect(text).toContain('Mica placement ledger');
      expect(text).toContain('Mica ledger closed');
      for (const pin of placementPins) expect(text.includes(pin)).toBe(mode === 'visible');
    } finally { await pdf.destroy(); }
  });
}

test('strict policy rejects wholly offpaper margin/paragraph source placement', async ({ browser }) => {
  test.setTimeout(120_000);
  await expect(renderDocxArtifacts(marginParagraphPlacementDocx('outside'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});
