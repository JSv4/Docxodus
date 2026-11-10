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
      // Word 2024 collapses each 18pt after with the next 18pt before (#1553): 12pt lines step 30pt.
      const labels = content.items.filter(item => 'str' in item).filter(item => /^Observatory label \d$/.test(item.str));
      expect(labels).toHaveLength(4);
      for (let i = 1; i < labels.length; i++)
        expect(Math.abs(labels[i - 1].transform[5] - labels[i].transform[5] - 30)).toBeLessThanOrEqual(.5);
    } finally { await task.destroy(); }
  });
}
