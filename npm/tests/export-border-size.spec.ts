import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateBorderSizeDocx, type BorderSizeOptions } from './docx-border-size-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const page = await pdf.getPage(1);
    const text = await page.getTextContent();
    const operators = await page.getOperatorList();
    const drawing = new Set([OPS.constructPath, OPS.stroke, OPS.fill, OPS.eoFill, OPS.setFillRGBColor,
      OPS.setStrokeRGBColor, OPS.setLineWidth, OPS.transform, OPS.save, OPS.restore, OPS.clip, OPS.eoClip]);
    return { count: pdf.numPages, view: page.view, items: text.items.filter(item => 'str' in item).map(item => ({
      text: item.str, transform: item.transform, width: item.width, height: item.height, font: text.styles[item.fontName],
    })), drawing: operators.fnArray.flatMap((op, index) => drawing.has(op)
      ? [{ op, args: JSON.parse(JSON.stringify(operators.argsArray[index])) }] : []) };
  } finally { await task.destroy(); }
}

for (const [name, options] of Object.entries<BorderSizeOptions>({
  single: {}, double: { style: 'double' }, dashed: { style: 'dashed' },
  cell: { cell: true }, nil: { style: 'nil' }, none: { style: 'none' },
})) {
  test(`${name} omitted border size exports with the explicit-zero mapping`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const exportOptions = { browser, outputs: ['html', 'pdf'] as const,
      reviewProfile: 'markup' as const, commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const source = generateBorderSizeDocx(options);
    const actual = await renderDocxArtifacts(source, exportOptions);
    const control = await renderDocxArtifacts(generateBorderSizeDocx({ ...options, size: '0' }), exportOptions);
    const actualFacts = await facts(actual.pdf!);
    expect(actualFacts.count).toBe(1);
    expect(actualFacts.items.map(item => item.text).join('').trim()).toBe('BorderMarker Cypress archive checked.');
    expect(actualFacts).toEqual(await facts(control.pdf!));
    expect(actual.pageMap.fragments.some(f => f.story === 'body')).toBe(true);
    writeFileSync(testInfo.outputPath(`${name}.docx`), source);
    writeFileSync(testInfo.outputPath(`${name}.pdf`), actual.pdf!);
  });
}

test('malformed present border size remains a conversion failure', async ({ browser }) => {
  await expect(renderDocxArtifacts(generateBorderSizeDocx({ size: 'bad' }), {
    browser, outputs: ['pdf'], reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
  })).rejects.toMatchObject({ code: 'conversion_failure', phase: 'docx_conversion' });
});
