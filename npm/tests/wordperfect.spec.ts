import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { wordTextScript } from './word-text-script.js';

const metrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');

for (const [name, text, indent, last] of [
  ['indented line', 'Copper lanterns swing above the quiet harbor wall while gull sleeps by the tide tables until evening.', 36, 'gull'],
  ['condensation beyond the modern gap floor', 'W'.repeat(22) + ' ' + 'M'.repeat(21) + ' ' + 'N'.repeat(21), 0, 'N'.repeat(21)],
  ['single word', 'W'.repeat(66), 0, null],
  ['exhausted gap', 'W'.repeat(33) + ' ' + 'M'.repeat(33), 0, null],
] as const) {
  test(`WordPerfect ${name} respects the actual available width`, async ({ page }) => {
    await page.setContent(`<main data-docx-compatibility-mode="14" data-docx-wordperfect-justification="true">
      <p style="width:468pt;font:12pt 'Liberation Mono';line-height:normal;--docx-auto-line-spacing:1;
        margin:0;text-align:justify;text-indent:${indent}pt"><span style="white-space:pre-wrap;line-height:normal;
        --docx-font-kerning:none">${text}</span></p></main>`);
    await page.addScriptTag({ content:metrics }); await page.addScriptTag({ content:wordTextScript });
    const result = await page.evaluate(() => {
      const root = document.querySelector('main')!, run = root.querySelector('span')!, api = window as any;
      api.applyUnroundedNormalLineHeights(root); api.prepareWordText(root);
      const fitted = run.hasAttribute('data-docx-word-fit');
      api.paintWordText(root);
      const line = root.querySelector('svg text');
      return { fitted, text:root.textContent!.trim(), first:line?.textContent?.trim().split(/\s+/),
        x:line?.querySelector('tspan')?.getBoundingClientRect().left,
        left:root.querySelector('p')!.getBoundingClientRect().left };
    });
    expect(result.text).toBe(text);
    expect(result.fitted).toBe(last !== null);
    if (last !== null) {
      expect(result.first?.at(-1)).toBe(last);
      expect((result.x! - result.left) * .75).toBeCloseTo(indent, 1);
    }
  });
}
