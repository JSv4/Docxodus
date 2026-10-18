import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { decorationFixture } from './docx-text-decoration-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function drawingFacts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const page = await pdf.getPage(1);
    const text = await page.getTextContent();
    const operators = await page.getOperatorList();
    return { pages: pdf.numPages,
      text: text.items.filter(item => 'str' in item).map(item => item.str).join(''),
      paths: operators.fnArray.filter(op => op === OPS.constructPath).length,
      // Colors are retained even when the PDF renderer batches several dots into one path.
      colors: operators.fnArray.flatMap((op, i) => op === OPS.setFillRGBColor || op === OPS.setStrokeRGBColor
        ? [JSON.parse(JSON.stringify(operators.argsArray[i]))] : []),
    };
  } finally { await task.destroy(); }
}

for (const [name, properties, control] of [
  ['double-underline', '<w:u w:val="double"/>', '<w:u w:val="single"/>'],
  ['double-strike', '<w:dstrike/>', '<w:strike/>'],
  ['combined', '<w:strike/><w:u w:val="single"/>', '<w:u w:val="single"/>'],
  ['dotted-blue', '<w:u w:val="dotted" w:color="0000FF"/>', '<w:u w:val="single"/>'],
]) {
  test(`PDF retains ${name} decoration`, async ({ browser }, info) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'final' as const,
      commentProfile: 'hidden' as const, timeoutMs: 120_000 };
    const actual = await renderDocxArtifacts(decorationFixture([['CedarRules', properties]]), options);
    const simple = await renderDocxArtifacts(decorationFixture([['CedarRules', control]]), options);
    const facts = await drawingFacts(actual.pdf!);
    const baseline = await drawingFacts(simple.pdf!);
    expect(facts.pages).toBe(1);
    expect(facts.text).toBe('CedarRules');
    if (name === 'dotted-blue') expect(facts.colors).not.toEqual(baseline.colors);
    else expect(facts.paths).toBeGreaterThan(baseline.paths);
    writeFileSync(info.outputPath(`${name}.pdf`), actual.pdf!);
    writeFileSync(info.outputPath(`${name}-control.pdf`), simple.pdf!);
  });
}
