import { expect, test } from '@playwright/test';
import { getDocument, OPS, Util } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { writeFileSync } from 'node:fs';
import { changeBarFixture } from './docx-change-bar-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('PDF change bars share a column 30pt left of the text margin', async ({ browser }, info) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(changeBarFixture(), { browser, outputs: ['pdf'],
    reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'hidden', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const operators = await (await pdf.getPage(1)).getOperatorList();
    const bars: number[] = [];
    let transform = [1, 0, 0, 1, 0, 0];
    const stack: number[][] = [];
    operators.fnArray.forEach((op, i) => {
      const args = operators.argsArray[i];
      if (op === OPS.save) stack.push([...transform]);
      else if (op === OPS.restore) transform = stack.pop()!;
      else if (op === OPS.transform) transform = Util.transform(transform, args);
      else if (op === OPS.constructPath) {
        const bounds = args[2];
        const a = [bounds[0], bounds[1]];
        const b = [bounds[2], bounds[3]];
        Util.applyTransform(a, transform);
        Util.applyTransform(b, transform);
        // Paragraph bars are the only thin vertical drawing paths in this text-only fixture.
        if (Math.abs(Math.abs(b[0] - a[0]) - 0.75) < 0.01 && Math.abs(b[1] - a[1]) > 4)
          bars.push(Math.min(a[0], b[0]));
      }
    });
    expect(bars).toHaveLength(2);
    // The body starts at 72pt. Word's 30pt outside the column is 42pt for both paragraphs.
    for (const left of bars) expect(left).toBeCloseTo(42, 1);
    writeFileSync(info.outputPath('change-bars.pdf'), result.pdf!);
  } finally { await task.destroy(); }
});
