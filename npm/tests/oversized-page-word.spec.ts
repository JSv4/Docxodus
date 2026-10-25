import { expect, test } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';

// Native Word captures are the oracle. Preserve the observed fourth page after
// the oversized header and the first-page paragraph before the nested table.
const fixtures = fileURLToPath(new URL('../../TestFiles/OversizedPageBlocks/', import.meta.url));
const observations = JSON.parse(readFileSync(join(fixtures, 'word-observations.json'), 'utf8')) as {
  cases: Array<{ id: string; inputSha256: string; pageWords: Record<string, string[]> }>;
};

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function pageTexts(bytes: Uint8Array): Promise<string[]> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const doc = await task.promise;
    const texts: string[] = [];
    for (let page = 1; page <= doc.numPages; page++) {
      const content = await (await doc.getPage(page)).getTextContent();
      texts.push(content.items.map(item => 'str' in item ? item.str : '').join(' '));
    }
    return texts;
  } finally { await task.destroy(); }
}

for (const observed of observations.cases) {
  test(`${observed.id}: exports with the captured Word page inventory`, async ({ browser }) => {
    test.setTimeout(120_000);
    const bytes = new Uint8Array(readFileSync(join(fixtures, observed.id + '.docx')));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(observed.inputSha256);
    const actual = await renderDocxArtifacts(bytes, {
      browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 100_000,
    });
    const texts = await pageTexts(actual.pdf!);
    expect(actual.pageCount).toBe(Object.keys(observed.pageWords).length);
    for (const [page, words] of Object.entries(observed.pageWords)) {
      const counts = (items: string[]) => items.reduce<Record<string, number>>((result, word) => {
        result[word] = (result[word] ?? 0) + 1;
        return result;
      }, {});
      expect(counts(texts[Number(page) - 1].trim().split(/\s+/).filter(Boolean)), `Word page ${page}`)
        .toEqual(counts(words));
    }
  });
}
