import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { textboxClipDocx } from './docx-textbox-clip-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const visible of [false, true]) {
  test(`fixed textbox with ${visible ? 'visible' : 'blank'} clipped paragraph`, async ({ browser }) => {
    test.setTimeout(120_000);
    const pending = renderDocxArtifacts(textboxClipDocx(visible), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000,
    });
    const result = await pending;
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.warnings.filter(warning => warning.code === 'source_text_clipped')).toHaveLength(visible ? 1 : 0);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const content = await (await pdf.getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      for (const label of ['Sighting logged.', 'Orbit verified.']) expect(text.split(label)).toHaveLength(2);
      expect(text).not.toContain('Off-box signal.');
    } finally { await task.destroy(); }
  });
}
