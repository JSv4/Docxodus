import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';

const fixtures = new URL('../../TestFiles/ExactJustificationFlow/', import.meta.url);
const observations = JSON.parse(readFileSync(new URL('word-observations.json', fixtures), 'utf8'));

for (const source of observations.cases) {
  test(`exact justified lines preserve Word's mode ${source.compatibilityMode} choices`, async ({ browser }) => {
    const bytes = readFileSync(new URL(source.file, fixtures));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(source.inputSha256);
    const result = await renderDocxArtifacts(bytes, { browser, outputs: ['pdf'], timeoutMs: 120_000,
      reviewProfile: 'final', commentProfile: 'hidden' });
    expect(result.pageCount).toBe(source.nativePageCount);
    expect(result.renderReport.unsupportedContent).toEqual([]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const content = await (await pdf.getPage(1)).getTextContent();
      const rows = new Map<number, string[]>();
      for (const item of content.items) if ('str' in item && item.str.trim()) {
        const baseline = Math.round(item.transform[5] * 100) / 100;
        const words = rows.get(baseline) ?? []; words.push(item.str.trim()); rows.set(baseline, words);
      }
      expect(Array.from(rows.values(), words => words.join(' '))).toEqual(source.nativeLines);
      const baselines = Array.from(rows.keys());
      expect(792 - baselines[0]).toBeCloseTo(57, 1);
      expect(baselines[0] - baselines[1]).toBeCloseTo(15, 1);
    } finally { await task.destroy(); }
  });
}
