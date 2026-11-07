import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { customPathDocx } from './docx-custom-path-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['polygon', 'quadratic', 'cubic', 'open', 'hole', 'sameWinding', 'multiple', 'invisible'] as const) {
  for (const grouped of [false, true]) {
    test(`${kind} ${grouped ? 'grouped' : 'alone'} preserves custom path paint in public PDF`, async ({ browser, page }) => {
      const result = await renderDocxArtifacts(customPathDocx(kind, grouped), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
      });
      expect(result.pageCount).toBe(1);
      expect(result.renderReport.unsupportedContent).toEqual([]);
      await page.setContent(result.html!);
      const paths = page.locator('.page-content svg path');
      await expect(paths).toHaveCount(kind === 'multiple' ? 2 : 1);
      if (kind === 'multiple') {
        await expect(paths.nth(1)).toHaveAttribute('d', 'M 182880 457200 L 1645920 457200');
        await expect(paths.nth(1)).toHaveAttribute('fill', 'none');
      }
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdfPage = await (await task.promise).getPage(1);
        const text = (await pdfPage.getTextContent()).items.flatMap(item => 'str' in item ? [item.str] : []).join('');
        expect(text).toBe('Optical panelEnd of inspection');
        const ops = await pdfPage.getOperatorList();
        // Inspect the authored drawing colors, excluding the printed white page background.
        let fill = '', stroke = '';
        const stack: [string, string][] = [], fills: number[] = [], strokes: number[] = [];
        ops.fnArray.forEach((op, i) => {
          if (op === OPS.save) stack.push([fill, stroke]);
          else if (op === OPS.restore) [fill, stroke] = stack.pop() ?? ['', ''];
          else if (op === OPS.setFillRGBColor) fill = ops.argsArray[i][0];
          else if (op === OPS.setStrokeRGBColor) stroke = ops.argsArray[i][0];
          else if (op === OPS.constructPath) {
            const mode = ops.argsArray[i][0];
            if (fill === '#19647e' && [OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke].includes(mode)) fills.push(mode);
            if (stroke === '#883311' && [OPS.stroke, OPS.fillStroke, OPS.eoFillStroke].includes(mode)) strokes.push(mode);
          }
        });
        expect(fills.length).toBe(kind === 'open' || kind === 'invisible' ? 0 : 1);
        if (kind === 'hole' || kind === 'sameWinding') expect(fills).toEqual([OPS.eoFill]);
        expect(strokes.length).toBe(kind === 'open' || kind === 'multiple' ? 1 : 0);
      } finally { await task.destroy(); }
    });
  }
}

for (const kind of ['arc', 'guide', 'dimensions', 'fill', 'text', 'effect'] as const) {
  for (const grouped of [false, true]) {
    test(`${kind} ${grouped ? 'grouped' : 'alone'} remains diagnosed and rejected under strict policy`, async ({ browser }) => {
      const source = customPathDocx(kind === 'text' || kind === 'effect' ? 'polygon' : kind, grouped,
        { text: kind === 'text', effect: kind === 'effect' });
      const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const, commentProfile: 'hidden' as const };
      const result = await renderDocxArtifacts(source, options);
      expect(result.pageCount).toBe(1);
      expect(result.renderReport.unsupportedContent.length).toBeGreaterThan(0);
      if (kind === 'text') expect(result.html).toContain('Custom caption');
      if (kind === 'effect') expect(result.html).toContain('<path');
      await expect(renderDocxArtifacts(source, { ...options, unsupportedContent: 'strict' }))
        .rejects.toMatchObject({ code: 'resource_policy_failure' });
    });
  }
}

for (const opacity of ['alpha', 'alphaMod'] as const) for (const grouped of [false, true]) {
  test(`${opacity} ${grouped ? 'grouped' : 'alone'} preserves supported opacity or diagnoses loss`, async ({ browser }) => {
    const source = customPathDocx('polygon', grouped, { opacity });
    const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const, commentProfile: 'hidden' as const };
    const result = await renderDocxArtifacts(source, options);
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.unsupportedContent.length).toBe(opacity === 'alpha' ? 0 : 1);
    if (opacity === 'alpha') expect(result.html).toContain('fill-opacity="0.4"');
    else await expect(renderDocxArtifacts(source, { ...options, unsupportedContent: 'strict' }))
      .rejects.toMatchObject({ code: 'resource_policy_failure' });
  });
}
