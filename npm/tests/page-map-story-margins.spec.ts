import { expect, test } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const bundle = join(dirname(fileURLToPath(import.meta.url)), '../dist/pagination.bundle.js');

for (const story of ['header', 'footer'] as const) {
  for (const side of ['left', 'right'] as const) {
    test(`${story} table preserves its ${side} margin cell`, async ({ page }) => {
      await page.setContent('<div id="viewer"></div>');
      await page.addScriptTag({ path: bundle });
      const result = await page.evaluate(({ story, side }) => {
        const scope = story === 'header' ? 'hdr1' : 'ftr1';
        const cells = side === 'left' ? [24, 228] : [228, 24];
        const html = `<div id="pagination-staging">
          <div id="pagination-hf-registry" style="display:none">
            <div data-section="0" data-hf-type="${story}-default">
              <table data-source-anchor-id="tbl:${scope}:table" style="border-collapse:collapse;table-layout:fixed;width:252pt;margin-left:${side === 'left' ? -24 : 0}pt">
                <colgroup>${cells.map(width => `<col style="width:${width}pt">`).join('')}</colgroup>
                <tbody><tr>${cells.map((_, i) => `<td data-source-anchor-id="tc:${scope}:cell${i}" style="padding:0;font:8pt/12pt Arial">Rim</td>`).join('')}</tr></tbody>
              </table>
            </div>
          </div>
          <div data-section-index="0" data-page-width="300" data-page-height="360"
            data-content-width="228" data-content-height="288"
            data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36"
            data-header-distance="12" data-footer-distance="12">
            <p data-source-anchor-id="p:body:caption" style="font:10pt/12pt Arial;margin:0">Orbital survey.</p>
          </div>
        </div><div id="pagination-container"></div>`;
        const pagination = (window as any).DocxodusPagination.paginateHtml(html, 'viewer', {
          showPageNumbers: false,
          layoutToken: { documentVersion: 0, rendererFingerprint: 'story-margins' },
        });
        const band = document.querySelector<HTMLElement>(`.page-${story}`)!;
        const style = getComputedStyle(band);
        return { fragments: pagination.pageMap.fragments, overflowX: style.overflowX, overflowY: style.overflowY };
      }, { story, side });
      expect(result.fragments.filter((f: any) => f.anchorId.startsWith('tc:'))).toHaveLength(2);
      expect(result.overflowX).toBe('visible');
      expect(result.overflowY).toBe('clip');
    });
  }
}
