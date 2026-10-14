import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateFootnoteLineBreakDocx, type FootnoteLineBreakOptions } from './docx-footnote-line-break-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    return await Promise.all(Array.from({ length: pdf.numPages }, async (_, index) => {
      const page = await pdf.getPage(index + 1);
      const text = await page.getTextContent();
      return { view: page.view, items: text.items.filter(item => 'str' in item).map(item => ({
        text: item.str, transform: item.transform, width: item.width, height: item.height,
        font: text.styles[item.fontName],
      })), links: (await page.getAnnotations()).filter(a => a.subtype === 'Link').map(a => ({ url: a.url, rect: a.rect })) };
    }));
  } finally { await task.destroy(); }
}

for (const [name, options] of Object.entries<FootnoteLineBreakOptions>({
  plain: {}, formatted: { formatted: true }, blanks: { blankEvery: 19 }, hyperlinks: { hyperlink: true },
  short: { lineCount: 4 },
})) {
  test(`${name} footnote hard breaks preserve complete PDF text and note-band geometry`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const exportOptions = { browser, outputs: ['html', 'pdf'] as const,
      reviewProfile: 'markup' as const, commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const source = generateFootnoteLineBreakDocx(options);
    const actual = await renderDocxArtifacts(source, exportOptions);
    const control = await renderDocxArtifacts(generateFootnoteLineBreakDocx({ ...options, paragraphs: true }), exportOptions);
    expect(actual.pageCount).toBe(control.pageCount);
    expect(actual.pageCount).toBe(name === 'short' ? 1 : 5);
    const actualFacts = await facts(actual.pdf!);
    const controlFacts = await facts(control.pdf!);
    expect(actualFacts.map(p => p.view)).toEqual(controlFacts.map(p => p.view));
    // Later source paragraphs have the converter's existing 12pt note inset.
    // A hard break stays inside one source paragraph, so compare the complete
    // glyph inventory and typography without claiming identical paragraph layout.
    const glyphs = (pages: typeof actualFacts) => pages.flatMap(p => p.items)
      .filter(item => item.text.trim()).map(item => ({
        // PDF.js subtracts different absolute x positions to obtain widths;
        // the measured pair differs by <2e-13pt. Compare to nine decimals.
        text: item.text.trim(), width: Number(item.width.toFixed(9)), height: item.height,
        font: item.font, scale: item.transform.slice(0, 4),
      }));
    expect(glyphs(actualFacts)).toEqual(glyphs(controlFacts));
    expect(actualFacts.flatMap(p => p.links).map(a => a.url))
      .toEqual(controlFacts.flatMap(p => p.links).map(a => a.url));
    for (const page of actualFacts) {
      for (const item of page.items.filter(i => /^Note\d+/.test(i.text))) {
        // The PDF's scale matrix projects the 72pt origin as 71.999997pt.
        expect(item.transform[4]).toBeGreaterThanOrEqual(72 - 0.00001);
        expect(item.transform[4] + item.width).toBeLessThanOrEqual(540);
        expect(item.transform[5]).toBeGreaterThan(72);
        expect(item.transform[5] + item.height).toBeLessThan(720);
      }
    }
    expect(actualFacts.flatMap(p => p.items).map(i => i.text).join(' ').match(/Note\d+/g))
      .toEqual(Array.from({ length: options.lineCount ?? 100 }, (_, index) => `Note${String(index + 1).padStart(3, '0')}`));
    if (options.blankEvery) {
      const locations = actualFacts.flatMap((page, pageIndex) => page.items.map(item => ({
        line: Number(item.text.match(/^Note(\d+)/)?.[1]), pageIndex, baseline: item.transform[5],
      }))).filter(item => item.line > 0);
      let checkedBlankLines = 0;
      for (let line = options.blankEvery; line < (options.lineCount ?? 100); line += options.blankEvery) {
        const before = locations.find(item => item.line === line)!;
        const after = locations.find(item => item.line === line + 1)!;
        if (before.pageIndex !== after.pageIndex) continue;
        // One empty line leaves two line boxes between the text baselines.
        // Measured gaps are 31.5–32.25pt for the authored 16pt lines;
        // distinguish one empty line from zero or two.
        const gap = before.baseline - after.baseline;
        expect(gap).toBeGreaterThan(16 * 1.5);
        expect(gap).toBeLessThan(16 * 2.5);
        checkedBlankLines++;
      }
      expect(checkedBlankLines).toBeGreaterThan(0);
    }
    const fragments = actual.pageMap.fragments.filter(f => f.story === 'footnote' && f.anchorId.startsWith('p:fn:'));
    expect(new Set(fragments.map(f => f.pageNumber)).size).toBe(actual.pageCount);
    expect(new Set(fragments.map(f => f.anchorId)).size).toBe(1);
    expect(new Set(fragments.map(f => f.fragmentId)).size).toBe(fragments.length);
    writeFileSync(testInfo.outputPath(`${name}.docx`), source);
    writeFileSync(testInfo.outputPath(`${name}.pdf`), actual.pdf!);
  });
}
