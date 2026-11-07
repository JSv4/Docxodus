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
        expect(await math.evaluate(element => element.closest('p')!.textContent)).not.toContain('\u00a0');
      if (layout === 'inline') {
        await page.evaluate(() => document.fonts.ready);
        expect(await math.evaluate(element => {
          const wrapper = element.parentElement!, paragraph = wrapper.closest('p')!;
          const box = (node: Element) => { const r = node.getBoundingClientRect(); return [r.x, r.y, r.width, r.height]; };
          const before = [box(element), box(paragraph)];
          wrapper.replaceWith(element);
          return JSON.stringify(before) === JSON.stringify([box(element), box(paragraph)]);
        })).toBe(true);
      }
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const pages = await Promise.all(Array.from({ length: pdf.numPages }, async (_, i) =>
          (await (await pdf.getPage(i + 1)).getTextContent()).items.filter(item => 'str' in item)));
        const items = pages.flat();
        const text = items.map(item => item.str).join('').replace(/\s/g, '');
        expect(text.match(/MathBefore/g)).toHaveLength(1);
        expect(text.match(/MathAfter/g)).toHaveLength(1);
        // MathML's normative math-auto transform paints the mathematical italic code point.
        expect(text.match(/𝑥/gu)).toHaveLength(1);
        if (kind === 'fraction') expect(text.match(/31\.4/g)).toHaveLength(1);
        if (layout === 'long') {
          expect(pdf.numPages).toBeGreaterThan(1);
          expect(text.match(/Birch\d{3}/g)).toEqual(Array.from({ length: 100 }, (_, i) => `Birch${String(i).padStart(3, '0')}`));
          const operandPage = pages.findIndex(items => items.some(item => item.str.includes('𝑥')));
          if (kind === 'fraction') expect(pages[operandPage].map(item => item.str).join('')).toContain('31.4');
          if (kind === 'indexed') expect(pages[operandPage].some(item => item.str.trim() === '3')).toBe(true);
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
      fontDirectories: ['/usr/share/fonts/truetype/liberation'],
    });
    await page.setContent(result.html!);
    const math = page.locator('.page-content math');
    expect(await math.evaluate(element => getComputedStyle(element).fontSize)).toBe(paragraphSize ? '24px' : '16px');
    expect(await math.evaluate(element => getComputedStyle(element).fontFamily)).toContain('Liberation Sans');
    expect(result.renderReport.unsupportedContent).toEqual([]);
  });
}

test('strict font coverage rejects a body font missing the actual mathematical italic glyph', async ({ browser }) => {
  await expect(renderDocxArtifacts(mathDocx('fraction', 'inline', '<m:mathFont m:val="Liberation Sans"/>'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
    fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'font_loading' });
});

for (const kind of ['fraction', 'root', 'indexed'] as const) {
  test(`${kind} uses the supplied math face for strict export and its layout containers`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(mathDocx(kind, 'inline', '<m:mathFont m:val="DejaVu Math TeX Gyre"/>'), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
      fontDirectories: ['/usr/share/fonts/truetype/liberation', '/usr/share/fonts/truetype/dejavu'], strictFonts: true,
    });
    expect(result.renderReport.fonts).toContainEqual(expect.objectContaining({ resolvedFamily: 'DejaVu Math TeX Gyre', verified: true }));
    const observation = result.renderReport.fontObservation;
    expect(observation?.status).toBe('observed');
    if (observation?.status === 'observed') expect(observation.uses.some(use => use.faces.some(face =>
      face.isCustomFont && face.family.includes('DejaVu Math')))).toBe(true);
    await page.setContent(result.html!);
    const families = await page.locator('.page-content math, .page-content msqrt, .page-content mroot').evaluateAll(elements =>
      elements.map(element => getComputedStyle(element).fontFamily));
    expect(new Set(families).size).toBe(1);
    expect(families[0]).not.toContain('DejaVu Math'); // The digest-bound face alias owns the outer math layout too.
  });
}
