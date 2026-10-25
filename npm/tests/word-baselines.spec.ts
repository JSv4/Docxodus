import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';

const wordText = readFileSync(new URL('../dist/word-text.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const lineMetrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const observations = JSON.parse(readFileSync(new URL('../../TestFiles/BaselineFlow/word-observations.json', import.meta.url), 'utf8'));
const source = readFileSync(new URL(`../../TestFiles/BaselineFlow/${observations.sourceFile}`, import.meta.url));
const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

for (const paragraph of observations.paragraphs) {
  test(`${paragraph.fontFamily} ${paragraph.fontSizePt}pt ${paragraph.multiple} spacing retains native baseline pitch`, async ({ page }) => {
    expect(createHash('sha256').update(source).digest('hex')).toBe(observations.sourceSha256);
    // The open fonts have the captured hhea metrics. The separate public-export
    // comparison uses the exact declared Times New Roman and Arial fonts.
    await page.setContent(`<style>p{font:${paragraph.fontSizePt}pt '${paragraph.cssFont}',serif;
      line-height:normal;margin:0;width:${paragraph.widthPt}pt;--docx-auto-line-spacing:${paragraph.multiple}}
      .run{white-space:pre-wrap;position:static;line-height:calc(1lh * ${paragraph.multiple});--docx-font-kerning:none}
      </style><main data-docx-compatibility-mode="15"><p><span class="run">${escape(paragraph.text)}</span></p></main>`);
    await page.addScriptTag({ content: lineMetrics });
    await page.addScriptTag({ content: wordText });
    const actual = await page.evaluate(() => {
      const api = window as any, root = document.querySelector('main')!;
      api.applyUnroundedNormalLineHeights(root);
      const prepared = api.prepareWordText(root), painted = api.paintWordText(root);
      return { prepared, painted, text:root.textContent,
        lines:Array.from(root.querySelectorAll('svg > text')).map(t => ({
          text:t.textContent!.trimEnd(), baselinePt:Number(t.getAttribute('y')) * .75 / 4 })) };
    });
    expect(actual.prepared).toBe(1);
    expect(actual.painted).toBe(1);
    expect(actual.text).toBe(paragraph.text);
    expect(actual.lines.map(line => line.text)).toEqual(paragraph.nativeLines.map((line: any) => line.text));
    let segmentStart = 0;
    for (let i = 0; i < paragraph.nativeLines.length; i++) {
      if (i === 0 || paragraph.nativeLines[i].page !== paragraph.nativeLines[i - 1].page) segmentStart = i;
      const expected = paragraph.nativeLines[i].baselinePt - paragraph.nativeLines[segmentStart].baselinePt;
      const measured = actual.lines[i].baselinePt - actual.lines[segmentStart].baselinePt;
      // Captured PDF placement varies within this small band. Premature line
      // rounding accumulates to0.72pt in the long-paragraph controls.
      expect(Math.abs(measured - expected)).toBeLessThan(.275);
    }
  });
}
