import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { columnFragmentationDocx, type ColumnContent } from './docx-column-fragmentation-fixture.js';

for (const kind of ['paragraph', 'wrapped', 'table'] as ColumnContent[]) {
  for (const [columns, sequential, count] of [[1, false, 77], [2, false, 8], [2, false, 77], [2, true, 77]] as const) {
    test(`${kind}: ${count} labels flow through ${columns} ${sequential ? 'sequential' : 'balanced'} columns`, async ({ browser }) => {
      const result = await renderDocxArtifacts(columnFragmentationDocx(kind, columns, sequential, count), {
        browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden',
        fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
      });
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise, labels: string[] = [], text: string[] = [], columnsUsed = new Set<number>();
        expect(pdf.numPages).toBe(result.pageCount);
        expect(pdf.numPages).toBe(count === 8 ? 1 : kind === 'wrapped' ? 3 : columns === 1 ? 8 : 4);
        for (let n = 1; n <= pdf.numPages; n++) {
          const items = (await (await pdf.getPage(n)).getTextContent()).items;
          for (const item of items) {
            if (!('str' in item) || !item.str.trim()) continue;
            expect(item.transform[4]).toBeGreaterThanOrEqual(19.9);
            expect(item.transform[4] + item.width).toBeLessThanOrEqual(320.1);
            expect(item.transform[5]).toBeGreaterThan(20);
            expect(item.transform[5]).toBeLessThan(180);
            columnsUsed.add(item.transform[4] >= 179.9 ? 2 : 1);
          }
          const pageText = items.map(item => 'str' in item ? item.str : '').join('');
          text.push(pageText);
          labels.push(...pageText.match(/Maple(?:\d{3}|End)/g) ?? []);
        }
        expect(labels).toEqual([...Array.from({ length: count }, (_, i) => `Maple${String(i + 1).padStart(3, '0')}`), 'MapleEnd']);
        expect(text.at(-1)).toContain(`Maple${String(count).padStart(3, '0')}`);
        if (kind === 'wrapped') expect(text.join('').match(/register\./g)).toHaveLength(count);
        if (columns === 2 && count === 77) expect([...columnsUsed].sort()).toEqual([1, 2]);
      } finally { await task.destroy(); }
    });
  }
}
