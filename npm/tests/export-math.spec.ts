import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { mathDocx } from './docx-math-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['fraction', 'root', 'indexed'] as const) {
  for (const layout of ['inline', 'display', 'mixed', 'edge', 'long'] as const) {
    test(`${kind} ${layout} keeps selectable operands and surrounding prose in public PDF`, async ({ browser, page }) => {
      const result = await renderDocxArtifacts(mathDocx(kind, layout), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
      });
      expect(result.renderReport.unsupportedContent).toEqual([]);
      await page.setContent(result.html!);
      const math = page.locator('.page-content math');
      await expect(math).toHaveCount(1);
      await expect(math.locator(kind === 'fraction' ? 'mfrac' : kind === 'root' ? 'msqrt' : 'mroot')).toHaveCount(1);
      expect(await math.locator('mi').allTextContents()).toEqual(['x']);
      expect(await math.locator('mn').allTextContents()).toEqual(kind === 'fraction' ? ['31.4'] : kind === 'indexed' ? ['3'] : []);
      // The paragraph must not acquire a blank text line after an equation-only line.
      if (layout === 'display' || layout === 'inline')
        expect(await math.evaluate(element => element.parentElement!.textContent)).not.toContain('\u00a0');
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const items = (await Promise.all(Array.from({ length: pdf.numPages }, async (_, i) =>
          (await (await pdf.getPage(i + 1)).getTextContent()).items.filter(item => 'str' in item)))).flat();
        const text = items.map(item => item.str).join('').replace(/\s/g, '');
        expect(text.match(/MathBefore/g)).toHaveLength(1);
        expect(text.match(/MathAfter/g)).toHaveLength(1);
        expect(text.match(/x/g)).toHaveLength(1);
        if (kind === 'fraction') expect(text.match(/31\.4/g)).toHaveLength(1);
        if (layout === 'long') {
          expect(pdf.numPages).toBeGreaterThan(1);
          expect(text.match(/Birch\d{3}/g)).toEqual(Array.from({ length: 100 }, (_, i) => `Birch${String(i).padStart(3, '0')}`));
        }
        if (layout === 'mixed') {
          expect(text).toContain('MathAnchor');
          expect(text).toContain('MathTail');
        }
      } finally { await task.destroy(); }
    });
  }
}

for (const kind of ['accent', 'styled'] as const) {
  test(`${kind} keeps a whole-equation diagnostic and strict-policy refusal`, async ({ browser }) => {
    const source = mathDocx(kind, 'display');
    const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const, commentProfile: 'hidden' as const };
    const result = await renderDocxArtifacts(source, options);
    expect(result.renderReport.unsupportedContent).toEqual([expect.objectContaining({ contentType: 'MathEquation' })]);
    expect(result.html).not.toContain('<math');
    await expect(renderDocxArtifacts(source, { ...options, unsupportedContent: 'strict' }))
      .rejects.toMatchObject({ code: 'resource_policy_failure' });
  });
}

for (const paragraphSize of [undefined, 36]) {
  test(`math size follows ${paragraphSize ? 'paragraph style' : 'document defaults'} beside large prose and mark`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(mathDocx('fraction', 'mixed', '<m:mathFont m:val="Liberation Sans"/>',
      { largeProse: true, markSize: 60, paragraphSize }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
      fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
    });
    await page.setContent(result.html!);
    const math = page.locator('.page-content math');
    expect(await math.evaluate(element => getComputedStyle(element).fontSize)).toBe(paragraphSize ? '24px' : '16px');
    expect(await math.evaluate(element => getComputedStyle(element).fontFamily)).toContain('Liberation Sans');
    expect(result.renderReport.unsupportedContent).toEqual([]);
  });
}
