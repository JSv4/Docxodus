import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storySpacingDocx } from './docx-story-spacing-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const story of ['header', 'footer'] as const) {
  test(`the delivered PDF retains all spaced ${story} paragraphs`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(storySpacingDocx(story), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(1);
    expect(result.pageMap.fragments.filter(fragment => fragment.story === story)).toHaveLength(4);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const document = await task.promise;
      expect(document.numPages).toBe(1);
      const content = await (await document.getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      expect(text.match(/Observatory label \d/g)).toEqual(Array.from({ length: 4 }, (_, i) => `Observatory label ${i}`));
      expect(text.match(/Night log caption\./g)).toEqual(['Night log caption.']);
    } finally { await task.destroy(); }
  });
}
