import { expect, test } from '@playwright/test';
import { getDocument, OPS, Util } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { paragraphFragmentationDocx } from './docx-paragraph-fragmentation-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const cases: Array<[boolean, boolean, number]> = [
  [false, false, 42], [false, true, 42], [true, false, 42], [true, true, 42],
  [false, false, 84], [true, true, 84],
];
for (const [tracked, anchored, lineCount] of cases) {
  test(`${lineCount}-line paragraph: tracked=${tracked}, anchored=${anchored}`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(paragraphFragmentationDocx(tracked, anchored, lineCount), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'hidden', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(Math.ceil(lineCount / 22));
    expect(result.renderReport.unsupportedContent).toEqual([]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const text = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const sheet = await pdf.getPage(i);
        const content = await sheet.getTextContent();
        text.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
        for (const item of content.items) if ('str' in item && item.str.includes('Workshop row')) {
          expect(450 - item.transform[5]).toBeGreaterThanOrEqual(54);
          expect(450 - item.transform[5]).toBeLessThanOrEqual(396);
        }
        const ops = await sheet.getOperatorList();
        let transform = [1, 0, 0, 1, 0, 0], bars = 0;
        const stack: number[][] = [];
        ops.fnArray.forEach((op, index) => {
          const args = ops.argsArray[index];
          if (op === OPS.save) stack.push([...transform]);
          else if (op === OPS.restore) transform = stack.pop()!;
          else if (op === OPS.transform) transform = Util.transform(transform, args);
          else if (op === OPS.constructPath) {
            const bounds = args[2], a = [bounds[0], bounds[1]], b = [bounds[2], bounds[3]];
            Util.applyTransform(a, transform); Util.applyTransform(b, transform);
            if (Math.abs(Math.abs(b[0] - a[0]) - .75) < .01 && Math.abs(b[1] - a[1]) > 10) bars++;
          }
        });
        expect(bars).toBe(tracked ? 1 : 0);
      }
      expect(text.join(' ').match(/Workshop row \d{2} carries a copper token\./g)).toEqual(
        Array.from({ length: lineCount }, (_, i) => `Workshop row ${String(i + 1).padStart(2, '0')} carries a copper token.`));
      expect(text.join(' ').match(/ORANGE MARKER/g) ?? []).toHaveLength(anchored ? 1 : 0);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const paragraphs = page.locator('.page-content > p');
    await expect(paragraphs).toHaveCount(result.pageCount);
    const identities = await paragraphs.evaluateAll(elements => elements.map(element => ({
      source: element.getAttribute('data-source-anchor-id'), fragment: element.getAttribute('data-page-fragment-id'),
      author: element.querySelector('ins')?.getAttribute('data-author'),
      date: element.querySelector('ins')?.getAttribute('data-date'),
    })));
    expect(new Set(identities.map(i => i.source)).size).toBe(1);
    expect(new Set(identities.map(i => i.fragment)).size).toBe(result.pageCount);
    const bodyFragments = result.pageMap.fragments.filter(fragment => fragment.anchorId === identities[0].source);
    expect(bodyFragments.map(fragment => fragment.pageNumber)).toEqual(
      Array.from({ length: result.pageCount }, (_, i) => i + 1));
    if (tracked) for (const identity of identities)
      expect(identity).toMatchObject({ author: 'Fixture Author', date: '2026-02-03T10:00:00Z' });
    const anchors = page.locator('.page-box [data-docx-drawing-anchor]');
    await expect(anchors).toHaveCount(anchored ? 1 : 0);
    if (anchored) {
      const geometry = await anchors.evaluate(element => {
        const paper = element.closest('.page-box')!;
        const box = element.getBoundingClientRect(), rect = paper.getBoundingClientRect();
        return { page: Array.from(document.querySelectorAll('.page-box')).indexOf(paper) + 1,
          left: (box.left - rect.left) * .75, top: (box.top - rect.top) * .75 };
      });
      expect(geometry.page).toBe(1);
      expect(geometry.left).toBeCloseTo(54 + 4800000 / 12700, 1);
      expect(geometry.top).toBeCloseTo(54, 1);
      const source = await anchors.locator('[data-source-anchor-id]').getAttribute('data-source-anchor-id');
      expect(result.pageMap.fragments.filter(fragment => fragment.anchorId === source)
        .map(fragment => fragment.pageNumber)).toEqual([1]);
    }
  });
}
