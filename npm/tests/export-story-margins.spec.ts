import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storyMarginsDocx } from './docx-story-margins-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const story of ['header', 'footer'] as const) {
  for (const side of ['left', 'right'] as const) {
    test(`${story} PDF includes the ${side} margin cell`, async ({ browser }) => {
      test.setTimeout(120_000);
      const result = await renderDocxArtifacts(storyMarginsDocx(story, side), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
        commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000,
      });
      expect(result.pageCount).toBe(1);
      expect(result.pageMap.fragments.filter(fragment => fragment.story === story && fragment.anchorId.startsWith('tc:'))).toHaveLength(2);
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const content = await (await pdf.getPage(1)).getTextContent();
        const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
        for (const label of ['Orbital survey.', 'Survey table', 'Rim']) expect(text.split(label)).toHaveLength(2);
      } finally { await task.destroy(); }
    });
  }
}
