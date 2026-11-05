import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { textboxSpacingDocx, textboxSpacingLines } from './docx-textbox-spacing-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const legacy of [false, true]) for (const [contextual, implicitDefault] of [[true, false], [true, true], [false, false]] as const) {
  test(`${legacy ? 'VML' : 'DrawingML'} text box preserves ${contextual ? 'contextual' : 'ordinary'} ${implicitDefault ? 'default-style' : 'named-style'} paragraph spacing in PDF`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const input = textboxSpacingDocx(legacy, contextual, implicitDefault);
    writeFileSync(testInfo.outputPath('swift.docx'), input);
    const result = await renderDocxArtifacts(input, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const words = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      for (const line of [...textboxSpacingLines, 'Swift body marker']) expect(words).toContain(line);
      const lines = content.items.filter(item => 'str' in item && textboxSpacingLines.includes(item.str));
      expect(lines).toHaveLength(4);
      const y = lines.map(item => 'transform' in item ? item.transform[5] : 0);
      for (let i = 1; i < y.length; i++) expect(y[i - 1] - y[i]).toBeCloseTo(contextual ? 12 : 21, 1);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const sizes = await page.locator('span').evaluateAll(elements => elements.map(e => getComputedStyle(e))
      .filter(s => s.overflow === 'hidden').map(s => ({ height: parseFloat(s.height), width: parseFloat(s.width) })));
    expect(sizes.some(s => Math.abs(s.height - (contextual ? 72 : 120) * 4 / 3) < .02
      && Math.abs(s.width - 190 * 4 / 3) < .02)).toBe(true);
  });
}
