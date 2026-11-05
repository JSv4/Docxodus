import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateThumbnailDocx } from './docx-thumbnail-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

const options = { outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
  commentProfile: 'margin' as const, timeoutMs: 120_000 };

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const doc = await task.promise, page = await doc.getPage(1), content = await page.getTextContent();
    expect(doc.numPages).toBe(1);
    return { view: page.view, items: content.items.filter(item => 'str' in item).map(item => ({
      text: item.str, transform: item.transform, width: item.width, height: item.height, font: content.styles[item.fontName],
    })) };
  } finally { await task.destroy(); }
}

for (const variant of [
  { label: 'package-relative', owner: 'package' as const, absolute: false },
  { label: 'package-rooted', owner: 'package' as const, absolute: true },
  { label: 'document-relative', owner: 'document' as const, absolute: false },
]) {
  if (variant.owner === 'package') {
    test(`${variant.label} missing thumbnail warns while preserving the printed body`, async ({ browser }) => {
      test.setTimeout(120_000);
      const actual = await renderDocxArtifacts(generateThumbnailDocx('missing', variant.owner, variant.absolute), { browser, ...options });
      const expected = await renderDocxArtifacts(generateThumbnailDocx('none'), { browser, ...options });
      expect(await facts(actual.pdf!)).toEqual(await facts(expected.pdf!));
      expect(actual.renderReport.warnings.some(warning => warning.code === 'missing_target')).toBe(true);
    });
  }

  test(`${variant.label} missing thumbnail rejects under strict policy`, async ({ browser }) => {
    await expect(renderDocxArtifacts(generateThumbnailDocx('missing', variant.owner, variant.absolute), {
      browser, ...options, unsupportedContent: 'strict',
    }))
      .rejects.toMatchObject({ code: 'invalid_document', phase: 'package_preflight',
        message: 'The input is not a valid DOCX OPC package (opc).' });
  });

  test(`${variant.label} present thumbnail prints like the clean control`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateThumbnailDocx('present', variant.owner, variant.absolute);
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    const expected = await renderDocxArtifacts(generateThumbnailDocx('none'), { browser, ...options });
    const physical = await facts(actual.pdf!);
    expect(physical).toEqual(await facts(expected.pdf!));
    expect(physical.items.map(item => item.text).join(' ')).toContain('MapleTop The original records remain complete.');
    writeFileSync(testInfo.outputPath('thumbnail.docx'), bytes);
    writeFileSync(testInfo.outputPath('thumbnail.pdf'), actual.pdf!);
  });
}

test('thumbnail entries count toward the package-entry ceiling', async ({ browser }) => {
  test.setTimeout(120_000);
  await renderDocxArtifacts(generateThumbnailDocx('none'), { browser, ...options, limits: { opcEntries: 5 } });
  await expect(renderDocxArtifacts(generateThumbnailDocx('present'), {
    browser, ...options, limits: { opcEntries: 5 },
  })).rejects.toMatchObject({ code: 'resource_limit', phase: 'package_preflight' });
});
