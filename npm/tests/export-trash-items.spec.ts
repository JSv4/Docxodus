import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateTrashItemDocx } from './docx-trash-item-fixture.js';

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

for (const name of ['[trash]/0000.dat', '[trash]/ABCD.dat', '[trash]/00af.dat']) {
  for (const datDefault of [false, true]) test(`${name} with dat Default ${datDefault} prints like the clean control`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateTrashItemDocx(name, datDefault);
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    const expected = await renderDocxArtifacts(generateTrashItemDocx(null, datDefault), { browser, ...options });
    const physical = await facts(actual.pdf!);
    expect(physical).toEqual(await facts(expected.pdf!));
    const text = physical.items.map(item => item.text).join(' ');
    expect(text.match(/AspenTop/g)).toHaveLength(1);
    expect(text).not.toContain('DiscardedAspen');
    writeFileSync(testInfo.outputPath('trash.docx'), bytes);
    writeFileSync(testInfo.outputPath('trash.pdf'), actual.pdf!);
  });
}

for (const name of ['[trash]/unused.bin', '[trash]/000g.dat']) {
  test(`${name} remains invalid with a named package finding`, async ({ browser }) => {
    await expect(renderDocxArtifacts(generateTrashItemDocx(name), { browser, ...options }))
      .rejects.toMatchObject({ code: 'invalid_document', phase: 'package_preflight',
        partUri: '/' + name, detail: expect.stringContaining('missing_content_type') });
  });
}

test('trash entries count toward the package-entry ceiling', async ({ browser }) => {
  test.setTimeout(120_000);
  await renderDocxArtifacts(generateTrashItemDocx(null), { browser, ...options, limits: { opcEntries: 5 } });
  await expect(renderDocxArtifacts(generateTrashItemDocx('[trash]/0000.dat'), {
    browser, ...options, limits: { opcEntries: 5 },
  })).rejects.toMatchObject({ code: 'resource_limit', phase: 'package_preflight' });
});

test('a corrupted standard trash checksum still rejects the package', async ({ browser }) => {
  await expect(renderDocxArtifacts(generateTrashItemDocx('[trash]/0000.dat', false, true), { browser, ...options }))
    .rejects.toMatchObject({ code: 'invalid_document', phase: 'package_preflight',
      partUri: '/[trash]/0000.dat', detail: expect.stringContaining('crc_mismatch') });
});
