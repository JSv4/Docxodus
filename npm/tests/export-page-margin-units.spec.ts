import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generatePageMarginUnitDocx, type PageMarginName } from './docx-page-margin-unit-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const doc = await task.promise, pages = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i), content = await page.getTextContent();
      pages.push({ view: page.view, items: content.items.filter(item => 'str' in item).map(item => ({
        text: item.str, transform: item.transform, width: item.width, height: item.height, font: content.styles[item.fontName],
      })) });
    }
    return pages;
  } finally { await task.destroy(); }
}

for (const [edge, band, rows] of [['54pt', '18pt', 16], ['0.75in', '0.25in', 16], ['1.905cm', '0.635cm', 16],
  ['19.05mm', '6.35mm', 16], ['4.5pc', '1.5pc', 16], ['4.5pi', '1.5pi', 16], ['54pt', '18pt', 80]] as const) {
  test(`${edge} page and story margins print like twips (${rows} rows)`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const margins = Object.fromEntries((['top', 'right', 'bottom', 'left', 'header', 'footer'] as PageMarginName[])
      .map(name => [name, name === 'header' || name === 'footer' ? band : edge]));
    const bytes = generatePageMarginUnitDocx(margins, rows);
    const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
      commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const actual = await renderDocxArtifacts(bytes, options);
    const control = await renderDocxArtifacts(generatePageMarginUnitDocx({}, rows), options);
    const actualFacts = await facts(actual.pdf!);
    expect(actualFacts).toEqual(await facts(control.pdf!));
    expect(actualFacts).toHaveLength(rows === 80 ? 2 : 1);
    const text = actualFacts.flatMap(p => p.items).map(i => i.text).join(' ');
    for (let i = 1; i <= rows; i++) expect(text.match(new RegExp(`Juniper${String(i).padStart(3, '0')}`, 'g'))).toHaveLength(1);
    for (const marker of ['PageHead', 'PageFoot']) expect(text.match(new RegExp(marker, 'g'))).toHaveLength(actualFacts.length);
    writeFileSync(testInfo.outputPath(`margins-${edge}-${rows}.docx`), bytes);
    writeFileSync(testInfo.outputPath(`margins-${edge}-${rows}.pdf`), actual.pdf!);
  });
}

test('25mm left margin is honored in guarded PDF geometry', async ({ browser }, testInfo) => {
  const bytes = generatePageMarginUnitDocx({ left: '25mm' });
  const actual = await renderDocxArtifacts(bytes, { browser, outputs: ['pdf'],
    reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000 });
  const pages = await facts(actual.pdf!);
  expect(pages).toHaveLength(1);
  const marker = pages[0].items.find(item => item.text.startsWith('PageTop'))!;
  expect(marker).toBeDefined();
  // Existing HTML margins use one decimal place; Chromium snaps glyph origins to CSS pixels.
  expect(marker.transform[4]).toBeGreaterThan(70.8);
  expect(marker.transform[4]).toBeLessThan(71);
  writeFileSync(testInfo.outputPath('margins-25mm.docx'), bytes);
  writeFileSync(testInfo.outputPath('margins-25mm.pdf'), actual.pdf!);
});

test('malformed page margin is an explicit conversion failure', async ({ browser }) => {
  await expect(renderDocxArtifacts(generatePageMarginUnitDocx({ top: '12px' }), { browser, outputs: ['pdf'],
    reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
  })).rejects.toMatchObject({ code: 'conversion_failure', phase: 'docx_conversion' });
});
