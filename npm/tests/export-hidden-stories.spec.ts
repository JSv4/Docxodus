import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateHiddenStoryDocx } from './docx-hidden-story-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function pdfText(bytes: Uint8Array): Promise<string> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const document = await task.promise;
    const content = await (await document.getPage(1)).getTextContent();
    return content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
  } finally { await task.destroy(); }
}

for (const story of ['body', 'header', 'footer'] as const) {
  test(`hidden-only ${story} keeps the empty paragraph's verified line and omits hidden PDF text`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['html', 'pdf'] as const,
      reviewProfile: 'markup' as const, commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const control = await renderDocxArtifacts(generateHiddenStoryDocx(story, 'empty'), options);
    const actual = await renderDocxArtifacts(generateHiddenStoryDocx(story, 'hidden'), options);
    expect(actual.pageCount).toBe(1);
    expect(actual.pageMap.fragments.map(f => ({ story: f.story, geometry: f.geometry })))
      .toEqual(control.pageMap.fragments.map(f => ({ story: f.story, geometry: f.geometry })));
    const fragmentStory = story === 'body' ? 'body' : story;
    expect(actual.pageMap.fragments.some(f => f.story === fragmentStory && f.geometry.height > 0)).toBe(true);
    expect(await pdfText(actual.pdf!)).toEqual(await pdfText(control.pdf!));
    expect(await pdfText(actual.pdf!)).toContain('BODYMARK The archive tray is ready.');
    expect(await pdfText(actual.pdf!)).not.toContain('Concealed archive note.');
    writeFileSync(testInfo.outputPath(`hidden-${story}.pdf`), actual.pdf!);
  });
}

test('visible, mixed, and inherited header formatting retain their intended PDF text', async ({ browser }) => {
  test.setTimeout(120_000);
  for (const mode of ['visible', 'mixed', 'inherited'] as const) {
    const result = await renderDocxArtifacts(generateHiddenStoryDocx('header', mode), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const text = await pdfText(result.pdf!);
    expect(text).toContain('BODYMARK The archive tray is ready.');
    if (mode === 'visible') expect(text).toContain('Concealed archive note.');
    else expect(text).not.toContain('Concealed archive note.');
    if (mode === 'mixed') expect(text).toContain('Visible story label.');
  }
});
