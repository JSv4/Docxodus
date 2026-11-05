import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { textboxLineHeightDocx, textboxLineHeightLines } from './docx-textbox-line-height-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const legacy of [false, true]) for (const exactInner of [false, true]) {
  test(`${legacy ? 'VML' : 'DrawingML'} box keeps its ${exactInner ? 'exact' : 'natural'} lines independent of host spacing`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const render = async (hostMultiple: number) => {
      const input = textboxLineHeightDocx(legacy, hostMultiple, exactInner);
      writeFileSync(testInfo.outputPath(`heron-${hostMultiple}.docx`), input);
      const result = await renderDocxArtifacts(input, { browser, outputs: ['html', 'pdf'],
        reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
      expect(result.pageCount).toBe(1);
      await page.setContent(result.html!);
      await page.evaluate(() => document.fonts.ready);
      const geometry = await page.locator('span').evaluateAll(elements => {
        const box = elements.find(element => getComputedStyle(element).overflow === 'hidden')!;
        const rect = box.getBoundingClientRect();
        return { width: rect.width, height: rect.height, lines: Array.from(box.children).map(child => {
          const line = child.getBoundingClientRect();
          return { top: line.top - rect.top, bottom: line.bottom - rect.top };
        }) };
      });
      expect(geometry.width).toBeCloseTo(190 * 4 / 3, 1);
      expect(geometry.height).toBeCloseTo(68 * 4 / 3, 1);
      expect(geometry.lines).toHaveLength(4);
      for (const line of geometry.lines) {
        expect(line.top).toBeGreaterThanOrEqual(-.02);
        expect(line.bottom).toBeLessThanOrEqual(geometry.height + .02);
      }
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const content = await (await (await task.promise).getPage(1)).getTextContent();
        const items = content.items.filter(item => 'str' in item);
        const text = items.map(item => item.str).join(' ');
        for (const line of [...textboxLineHeightLines, 'Heron body marker']) expect(text).toContain(line);
        const boxed = items.filter(item => textboxLineHeightLines.includes(item.str));
        expect(boxed).toHaveLength(4);
        const first = boxed[0].transform;
        return { geometry, items: boxed.map(item => ({ text: item.str,
          transform: item.transform.map((value: number, index: number) => index >= 4 ? value - first[index] : value),
          width: item.width, height: item.height })) };
      } finally { await task.destroy(); }
    };
    const control = await render(1);
    const actual = await render(3);
    expect(actual.geometry).toEqual(control.geometry);
    actual.items.forEach((item, index) => {
      const expected = control.items[index];
      expect({ ...item, transform: item.transform.slice(0, 5) })
        .toEqual({ ...expected, transform: expected.transform.slice(0, 5) });
      // Chromium snaps PDF text origins to CSS pixels. Changing the containing
      // line's origin can move a relative printed baseline by one pixel (0.75pt),
      // while the DOM line geometry above must remain identical.
      expect(Math.abs(item.transform[5] - expected.transform[5])).toBeLessThanOrEqual(.75001);
    });
  });
}
