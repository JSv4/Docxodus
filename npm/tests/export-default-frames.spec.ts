import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { defaultFrameDocx, defaultFrameLines } from './docx-default-frame-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const frame of ['empty', 'auto'] as const) {
  test(`inherited ${frame} frame defaults preserve the flowing story`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const render = async (kind: 'none' | 'empty' | 'auto') => {
      const input = defaultFrameDocx(kind);
      writeFileSync(testInfo.outputPath(`tern-${kind}.docx`), input);
      const result = await renderDocxArtifacts(input, { browser, outputs: ['html', 'pdf'],
        reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
      expect(result.pageCount).toBe(2);
      expect(result.html).not.toContain('data-docx-text-frame="true"');
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const pages = [];
        for (let page = 1; page <= pdf.numPages; page++) {
          const content = await (await pdf.getPage(page)).getTextContent();
          pages.push(content.items.filter(item => 'str' in item).map(item => ({
            text: item.str, transform: item.transform, width: item.width, height: item.height,
          })));
        }
        const text = pages.flat().map(item => item.text).join(' ');
        for (const line of ['Tern opening marker', ...defaultFrameLines, 'Tern final marker'])
          expect(text).toContain(line);
        return pages;
      } finally { await task.destroy(); }
    };
    const control = await render('none');
    expect(await render(frame)).toEqual(control);
  });
}
