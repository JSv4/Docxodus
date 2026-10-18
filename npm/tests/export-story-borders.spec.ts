import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['footnote', 'endnote']) {
  for (const bordered of [true, false]) {
    const name = `paragraphs-story-borders-${kind}${bordered ? '' : '-control'}`;
    test(`${name} retains grouped story borders and note text in export`, async ({ browser, page }, info) => {
      test.setTimeout(120_000);
      const bytes = readFileSync(fileURLToPath(new URL(`../../TestFiles/StoryParagraphBorders/${name}.docx`, import.meta.url)));
      const result = await renderDocxArtifacts(bytes, {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
      });
      writeFileSync(info.outputPath(`${name}.pdf`), result.pdf!);
      await page.setContent(result.html!);
      const pages = page.locator('.page-box');
      expect(result.pageCount).toBe(1);
      await expect(pages).toHaveCount(1);
      await expect(pages.locator('[data-docx-border-paint]')).toHaveCount(bordered ? 4 : 0);
      // Four vector sides for each group, including every running/note story.
      await expect(pages.locator(':scope > svg > rect[fill="rgb(36, 107, 125)"]')).toHaveCount(bordered ? 16 : 0);
      for (const story of ['Body', 'Header', 'Footer', 'Note']) {
        const first = pages.locator('p').filter({ hasText: `Cedar${story}One` });
        await expect(first).toHaveCount(1);
        if (bordered) {
          await expect(first.locator('..')).toHaveAttribute('data-docx-border-paint');
          await expect(first.locator('..')).toContainText(`Cedar${story}Two`);
        }
      }
      for (const story of ['body', 'header', 'footer', kind])
        expect(result.pageMap.fragments.some(f => f.story === story && f.geometry.height > 0), story).toBe(true);
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const text = (await (await pdf.getPage(1)).getTextContent()).items
          .map(item => 'str' in item ? item.str : '').join(' ');
        for (const story of ['Body', 'Header', 'Footer', 'Note'])
          for (const suffix of ['One', 'Two', 'Plain'])
            expect(text.match(new RegExp(`Cedar${story}${suffix}`, 'g'))).toHaveLength(1);
      } finally { await task.destroy(); }
      await pages.screenshot({ path: info.outputPath(`${name}.png`) });
    });
  }
}
