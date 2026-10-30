import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { unpaintedSourceDocx } from './docx-unpainted-source-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['rowless-table', 'empty-cell', 'joined-hidden', 'joined-visible', 'standalone-hidden', 'cell-hidden', 'joined-hidden-empty'] as const) {
  test(`${kind} exports with complete visible content and no invented fragments`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(unpaintedSourceDocx(kind), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(1);
    if (kind === 'rowless-table')
      expect(result.pageMap.fragments.some(fragment => fragment.anchorId.startsWith('tbl:'))).toBe(false);
    if (kind === 'empty-cell')
      expect(result.pageMap.fragments.some(fragment => fragment.anchorId.startsWith('tc:'))).toBe(true);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const content = await (await pdf.getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      if (kind.startsWith('joined-')) {
        expect(text).toContain('Observation:');
        expect(text).toContain('Orbit verified.');
        if (kind === 'joined-hidden' || kind === 'joined-hidden-empty') expect(text).not.toContain('Internal telescope note.');
        else expect(text).toContain('Internal telescope note.');
      } else expect(text).toContain('Comet observation.');
    } finally { await task.destroy(); }
  });
}
