import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateCustomPropertiesDocx } from './docx-custom-properties-fixture.js';

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

for (const kind of ['mismatched', 'truncated', 'entity'] as const) {
  test(`${kind} custom properties reject before PDF rendering`, async ({ browser }) => {
    await expect(renderDocxArtifacts(generateCustomPropertiesDocx(kind), { browser, ...options }))
      .rejects.toMatchObject({ code: 'invalid_document', phase: 'package_preflight',
        message: 'The input is not a valid DOCX OPC package (opc).' });
  });
}

for (const kind of ['text', 'number', 'boolean'] as const) {
  test(`valid ${kind} custom properties print like the omitted-part control`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateCustomPropertiesDocx(kind);
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    const expected = await renderDocxArtifacts(generateCustomPropertiesDocx('none'), { browser, ...options });
    const physical = await facts(actual.pdf!);
    expect(physical).toEqual(await facts(expected.pdf!));
    const text = physical.items.map(item => item.text).join(' ');
    expect(text).toContain('CedarTop All body records stay visible.');
    expect(text).not.toContain('CedarMetadataOnly');
    writeFileSync(testInfo.outputPath('custom-properties.docx'), bytes);
    writeFileSync(testInfo.outputPath('custom-properties.pdf'), actual.pdf!);
  });
}

test('custom properties count toward the package-entry ceiling', async ({ browser }) => {
  test.setTimeout(120_000);
  await renderDocxArtifacts(generateCustomPropertiesDocx('none'), { browser, ...options, limits: { opcEntries: 5 } });
  await expect(renderDocxArtifacts(generateCustomPropertiesDocx('text'), {
    browser, ...options, limits: { opcEntries: 5 },
  })).rejects.toMatchObject({ code: 'resource_limit', phase: 'package_preflight' });
});
