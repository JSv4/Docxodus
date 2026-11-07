import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { literalTabFixture, literalLineEndingCases, literalLineEndingFixture, type LiteralTabCase } from './docx-literal-tab-fixture.js';

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

function expectSameVisibleText(actual: Awaited<ReturnType<typeof pdfFacts>>, control: Awaited<ReturnType<typeof pdfFacts>>) {
  expect(actual.count).toBe(control.count);
  const words = actual.text.filter(t => t.text.trim()), expected = control.text.filter(t => t.text.trim());
  expect(words.map(t => t.text)).toEqual(expected.map(t => t.text));
  words.forEach((word, i) => {
    word.transform.forEach((value, j) => expect(value).toBeCloseTo(expected[i].transform[j], 5));
    expect(word.width).toBeCloseTo(expected[i].width, 5);
  });
}

for (const [kind, name, ending, spaces] of literalLineEndingCases) {
  test(`${kind} text ${name} exports with space-equivalent text, ink and positions`, async ({ browser }) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['pdf'] as const, commentProfile: 'hidden' as const,
      reviewProfile: 'markup' as const, timeoutMs: 120_000 };
    const actual = await renderDocxArtifacts(literalLineEndingFixture(kind, ending), options);
    const control = await renderDocxArtifacts(literalLineEndingFixture(kind, spaces), options);
    const facts = await pdfFacts(actual.pdf!);
    expect(facts.count).toBe(1);
    expect(facts).toEqual(await pdfFacts(control.pdf!));
    for (const marker of ['NorthMark', 'HarborMark', 'TailMark'])
      expect(facts.text.filter(item => item.text.includes(marker))).toHaveLength(1);
  });
}

test('explicit br and cr elements retain a second PDF line', async ({ browser }) => {
  test.setTimeout(120_000);
  const results = [];
  for (const ending of ['<w:br/>', '<w:cr/>']) {
    const output = await renderDocxArtifacts(literalLineEndingFixture('body', ending, false), {
      browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    results.push(await pdfFacts(output.pdf!));
  }
  expectSameVisibleText(results[0], results[1]);
  const north = results[0].text.find(item => item.text.includes('NorthMark'))!;
  const harbor = results[0].text.find(item => item.text.includes('HarborMark'))!;
  expect(north.transform[5] - harbor.transform[5]).toBeCloseTo(16, 0);
});

for (const lf of [false, true]) {
  test(`encoded CR retains its break${lf ? ' and following LF space' : ''}`, async ({ browser }) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'final' as const,
      commentProfile: 'hidden' as const, timeoutMs: 120_000 };
    const actual = await renderDocxArtifacts(literalLineEndingFixture('body', '&#13;' + (lf ? '&#10;' : ''), false), options);
    const control = await renderDocxArtifacts(literalLineEndingFixture('body',
      '<w:br/>' + (lf ? '<w:t xml:space="preserve"> </w:t>' : ''), false), options);
    expectSameVisibleText(await pdfFacts(actual.pdf!), await pdfFacts(control.pdf!));
  });
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
