import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath:process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const [mode, flag, alignment, extra, tail, last] of [
  [14, '', 'both', '', true, 'gulls'],
  [14, '<w:wpJustification w:val="0"/>', 'both', '', true, 'gulls'],
  [14, '<w:wpJustification/>', 'both', '', true, 'naps'],
  [15, '<w:wpJustification/>', 'both', '', true, 'naps'],
  [14, '<w:wpJustification/>', 'left', '', true, 'gulls'],
  [14, '<w:wpJustification/>', 'both', 'X', true, 'napsX'],
  [14, '<w:wpJustification/>', 'both', 'XX', true, 'gulls'],
  [14, '<w:wpJustification/>', 'both', '', false, 'naps'],
] as const) {
  test(`WordPerfect mode ${mode} flag ${flag || 'absent'} ${alignment} extra ${extra.length} tail ${tail}`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const text = 'Copper lanterns swing above the quiet harbor wall while gulls naps' + extra
      + (tail ? ' beside the tide tables until the evening bell calls.' : '');
    const source = pageFittingDocx(`<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="${alignment}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/>
      <w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`,
      `${flag}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`);
    writeFileSync(testInfo.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, { browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden' });
    expect(result.renderReport.status).toBe('complete');
    writeFileSync(testInfo.outputPath('wordperfect.pdf'), result.pdf!);
    writeFileSync(testInfo.outputPath('wordperfect.html'), result.html!);
    await page.goto(pathToFileURL(testInfo.outputPath('wordperfect.html')).href);
    await page.evaluate(() => document.fonts.ready);
    expect(await page.locator('body').getAttribute('data-docx-wordperfect-justification'))
      .toBe(flag === '<w:wpJustification/>' ? 'true' : null);
    const advances = await page.locator('.page-box svg tspan').evaluateAll(spans => spans.map(span => {
      const text = span.textContent!.trimEnd(), range = document.createRange();
      range.setStart(span.firstChild!, 0); range.setEnd(span.firstChild!, text.length);
      return { text, width:range.getBoundingClientRect().width * .75 };
    }));
    expect(advances.length).toBeGreaterThan(0);
    for (const advance of advances) expect(advance.width).toBeCloseTo(advance.text.length * 1229 / 2048 * 12, 1);
    const task = getDocument({ data:new Uint8Array(result.pdf!), verbosity:0 });
    try {
      const pdf = await task.promise; expect(pdf.numPages).toBe(1);
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(i => 'str' in i && i.str.trim());
      const lines = new Map<number, string[]>();
      for (const item of items) if ('str' in item && 'transform' in item) {
        const y = Math.round(item.transform[5] * 100), row = lines.get(y) ?? [];
        row.push(item.str); lines.set(y, row);
        expect(item.transform[4]).toBeGreaterThanOrEqual(71.9);
        expect(item.transform[4] + item.width).toBeLessThanOrEqual(540.1);
      }
      const rows = [...lines.values()].map(row => row.join(' ').trim().split(/\s+/));
      expect(rows.flat()).toEqual(text.split(' '));
      expect(rows[0].at(-1)).toBe(last);
      if (last === 'naps') expect(rows[0].join(' ').length).toBe(66);
    } finally { await task.destroy(); }
  });
}
