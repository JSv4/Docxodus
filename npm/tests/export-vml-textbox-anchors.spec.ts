import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { vmlPanelLabels, vmlTextboxAnchorDocx } from './docx-vml-textbox-anchor-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const [basis, autoFit] of [['page', false], ['margin', false], ['inline', false], ['page', true]] as const) {
  test(`legacy ${basis}${autoFit ? ' auto-fit' : ''} text boxes retain their placement and visible text`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const input = vmlTextboxAnchorDocx(basis, autoFit);
    writeFileSync(testInfo.outputPath(`plover-${basis}.docx`), input);
    const result = await renderDocxArtifacts(input, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const geometry = await page.locator('.page-box span').evaluateAll(elements => {
      const boxes = elements.filter(e => getComputedStyle(e).overflow === 'hidden' && e.textContent?.trimStart().startsWith('Plover panel'));
      return boxes.map(box => {
        const paper = box.closest('.page-box')!;
        const rect = box.getBoundingClientRect(), pageRect = paper.getBoundingClientRect();
        return { x: (rect.x - pageRect.x) * .75, y: (rect.y - pageRect.y) * .75,
          width: rect.width * .75, height: rect.height * .75, anchored: box.hasAttribute('data-docx-drawing-anchor') };
      });
    });
    expect(geometry).toHaveLength(basis === 'inline' ? 1 : 8);
    geometry.forEach((box, index) => {
      expect(box.width).toBeCloseTo(242, 1);
      // Two 12pt lines plus two borders, each snapped to one CSS pixel.
      expect(box.height).toBeCloseTo(autoFit ? 25.5 : 96, 1);
      expect(box.anchored).toBe(basis !== 'inline');
      if (basis !== 'inline') {
        const margin = basis === 'margin' ? 72 : 0;
        expect(box.x).toBeCloseTo(margin + 29 + (index % 2) * 254, 1);
        expect(box.y).toBeCloseTo(margin + 58 + Math.floor(index / 2) * 137, 1);
      }
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const items = content.items.filter(item => 'str' in item);
      const text = items.map(item => item.str).join(' ');
      for (const label of [...(basis === 'inline' ? vmlPanelLabels.slice(0, 1) : vmlPanelLabels), 'Plover body control'])
        expect(text).toContain(label);
      geometry.forEach((box, index) => {
        const printed = items.find(item => item.str === vmlPanelLabels[index]);
        expect(printed).toBeDefined();
        // Printed origins can snap by one CSS pixel; the DOM coordinates above
        // still have to match the authored placement to a tenth of a point.
        expect(Math.abs(printed!.transform[4] - box.x - 1)).toBeLessThanOrEqual(.8);
        const baseline = 792 - printed!.transform[5] - box.y;
        expect(baseline).toBeGreaterThan(8);
        expect(baseline).toBeLessThan(14);
      });
      const body = items.find(item => item.str === 'Plover body control')!;
      if (basis !== 'inline') expect(792 - body.transform[5]).toBeLessThan(110);
      else expect(792 - body.transform[5]).toBeGreaterThan(160);
    } finally { await task.destroy(); }
  });
}
