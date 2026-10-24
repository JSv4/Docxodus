import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { adjacentTablesDocx } from './docx-adjacent-tables-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function pdfFacts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const doc = await task.promise;
    const page = await doc.getPage(1);
    const content = await page.getTextContent();
    const ops = await page.getOperatorList();
    const drawing = new Set([OPS.constructPath, OPS.stroke, OPS.fill, OPS.setFillRGBColor, OPS.setStrokeRGBColor]);
    return { count: doc.numPages, view: page.view,
      text: content.items.filter(item => 'str' in item).map(item => ({ text: item.str, transform: item.transform })),
      drawing: ops.fnArray.flatMap((op, i) => drawing.has(op)
        ? [{ op, args: JSON.parse(JSON.stringify(ops.argsArray[i])) }] : []) };
  } finally { await task.destroy(); }
}

for (const separated of [false, true]) {
  for (const reviewProfile of ['final', 'original'] as const) {
    test(reviewProfile + ' preserves properties of ' + (separated ? 'separated tables (control)' : 'adjacent tables'),
      async ({ browser, page }, info) => {
        test.setTimeout(120_000);
        const bytes = adjacentTablesDocx(separated);
        const options = { browser, outputs: ['html', 'pdf'] as const, commentProfile: 'hidden' as const, timeoutMs: 120_000 };
        const actual = await renderDocxArtifacts(bytes, { ...options, reviewProfile });
        const control = await renderDocxArtifacts(bytes, { ...options, reviewProfile: 'markup' });
        await page.setContent(actual.html!);
        const tables = page.locator('.page-content table');
        await expect(tables).toHaveCount(2);
        await expect(tables.nth(0)).toContainText('FirstDock');
        await expect(tables.nth(1)).toContainText('SecondDock');
        const geometry = await tables.evaluateAll(elements => elements.map(table => {
          const cell = table.querySelector('td')!;
          const style = getComputedStyle(cell);
          return { width: table.getBoundingClientRect().width, padding: parseFloat(style.paddingLeft),
            border: style.borderTopStyle, color: style.borderTopColor };
        }));
        expect(geometry[1].width).toBeGreaterThan(geometry[0].width + 50);
        expect(geometry[1].padding).toBeCloseTo(48, 1);
        expect(geometry[1]).toMatchObject({ border: 'solid', color: 'rgb(36, 107, 125)' });
        const facts = await pdfFacts(actual.pdf!);
        expect(facts.count).toBe(1);
        expect(facts.drawing.length).toBeGreaterThan(0);
        expect(facts.text.map(item => item.text).join(' ')).toContain('SecondDock');
        expect(facts).toEqual(await pdfFacts(control.pdf!));
        writeFileSync(info.outputPath('tables.docx'), bytes);
        writeFileSync(info.outputPath('tables.pdf'), actual.pdf!);
        await page.locator('.page-box').screenshot({ path: info.outputPath('tables.png') });
      });
  }
}
