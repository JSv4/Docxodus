import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath:process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const full = 'Copper lanterns swing above the quiet harbor wall while gulls nap';
const indented = 'Copper lanterns swing above the quiet harbor wall while gull';
// The issue author confirms the original 65-cell line in Word. Indented and
// near-boundary cases exercise the same inferred fixed-cell measurement policy.
for (const mode of [14, 15]) for (const alignment of ['left', 'both']) for (const indent of [0, 36]) {
  test(`exact-fit mode ${mode} ${alignment} indent ${indent} exports every word within the column`, async ({browser, page}, testInfo) => {
    test.setTimeout(120_000);
    const prefix = indent ? indented : full;
    const text = prefix + ' beside the tide tables until the evening bell calls the crews home.';
    const source = pageFittingDocx(`<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      <w:ind w:firstLine="${indent * 20}"/><w:jc w:val="${alignment}"/></w:pPr><w:r><w:rPr>
      <w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:sz w:val="24"/>
      </w:rPr><w:t>${text}</w:t></w:r></w:p>`,
      `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`);
    writeFileSync(testInfo.outputPath('source.docx'), source);
    const result = await renderDocxArtifacts(source, {browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden'});
    expect(result.renderReport.status).toBe('complete');
    writeFileSync(testInfo.outputPath('exact-fit.pdf'), result.pdf!);
    writeFileSync(testInfo.outputPath('exact-fit.html'), result.html!);
    const task = getDocument({data:new Uint8Array(result.pdf!), verbosity:0});
    try {
      const pdf = await task.promise; expect(pdf.numPages).toBe(1);
      const content = await (await pdf.getPage(1)).getTextContent();
      const lines = new Map<number, string[]>();
      for (const item of content.items) if ('str' in item && item.str.trim()) {
        const y = Math.round(item.transform[5] * 100), row = lines.get(y) ?? [];
        row.push(item.str); lines.set(y, row);
        expect(item.height).toBeCloseTo(12, 2);
        expect(item.transform[4]).toBeGreaterThanOrEqual(71.9);
        expect(item.transform[4] + item.width).toBeLessThanOrEqual(540.1);
        expect(item.width / item.str.length).toBeCloseTo(1229 / 2048 * 12, 2);
      }
      const rows = [...lines.values()].map(row => row.join(' ').trim().split(/\s+/));
      expect(rows.flat()).toEqual(text.split(' '));
      expect(rows[0].join(' ')).toBe(prefix);
      await page.goto(pathToFileURL(testInfo.outputPath('exact-fit.html')).href);
      await page.evaluate(() => document.fonts.ready);
      const htmlLines = await page.locator('.page-box p svg > text').allTextContents();
      expect(htmlLines.map(line => line.trim().split(/\s+/))).toEqual(rows);
    } finally { await task.destroy(); }
  });
}
