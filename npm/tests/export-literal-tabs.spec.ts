import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { literalTabFixture, type LiteralTabCase } from './docx-literal-tab-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function pdfFacts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const doc = await task.promise, page = await doc.getPage(1), content = await page.getTextContent();
    const ops = await page.getOperatorList();
    const drawing = new Set([OPS.constructPath, OPS.stroke, OPS.fill, OPS.setFillRGBColor, OPS.setStrokeRGBColor]);
    return { count: doc.numPages, text: content.items.filter(item => 'str' in item).map(item => ({
      text: item.str, transform: item.transform, width: item.width,
    })), drawing: ops.fnArray.flatMap((op, i) => drawing.has(op)
      ? [{ op, args: JSON.parse(JSON.stringify(ops.argsArray[i])) }] : []) };
  } finally { await task.destroy(); }
}

for (const kind of ['body', 'header', 'footer', 'multiple', 'deleted', 'inserted'] as LiteralTabCase[]) {
  test('literal tabs export like explicit tabs in ' + kind, async ({ browser }, info) => {
    test.setTimeout(120_000);
    const bytes = literalTabFixture(kind);
    const options = { browser, outputs: ['pdf'] as const, commentProfile: 'hidden' as const, timeoutMs: 120_000,
      reviewProfile: kind === 'deleted' || kind === 'inserted' ? 'markup' as const : 'final' as const };
    const actual = await renderDocxArtifacts(bytes, options);
    const control = await renderDocxArtifacts(literalTabFixture(kind, false), options);
    const facts = await pdfFacts(actual.pdf!);
    expect(facts.count).toBe(1);
    const marker = facts.text.filter(item => item.text.includes('HarborMark'));
    expect(marker).toHaveLength(1);
    expect(marker[0].transform[4] + marker[0].width).toBeCloseTo(540, 0);
    expect(facts).toEqual(await pdfFacts(control.pdf!));
    if (kind === 'deleted' || kind === 'inserted') {
      const withoutLines = await renderDocxArtifacts(bytes, { ...options,
        additionalCss: 'del.rev-del, ins.rev-ins { text-decoration: none !important; }' });
      const plain = await pdfFacts(withoutLines.pdf!);
      expect(facts.text).toEqual(plain.text);
      expect(facts.drawing).not.toEqual(plain.drawing);
    }
    writeFileSync(info.outputPath('literal-tabs.docx'), bytes);
    writeFileSync(info.outputPath('literal-tabs.pdf'), actual.pdf!);
  });
}
