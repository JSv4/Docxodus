import { expect, test } from '@playwright/test';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const directory = path.dirname(fileURLToPath(import.meta.url));

for (const placement of ['body', 'header', 'unused-header'] as const) {
  for (const omitted of [false, true]) {
    test(`${placement} textbox source inventory detects pre-pagination omission=${omitted}`, async ({ page }) => {
      const scope = placement === 'body' ? 'body' : 'hdr1';
      const story = `<p data-source-anchor-id="p:${scope}:host"
          data-docx-required-source-anchors="p:${scope}:alder p:${scope}:birch" style="margin:0">
        <span style="display:inline-block;position:relative;width:200pt;height:30pt">
          <span data-docx-textbox="true" style="position:absolute;left:0;top:0;width:90pt;height:20pt">
            <span data-source-anchor-id="p:${scope}:alder" style="display:block">ALDER</span>
          </span>
          ${omitted ? '' : `<span data-docx-textbox="true" style="position:absolute;left:100pt;top:0;width:90pt;height:20pt">
            <span data-source-anchor-id="p:${scope}:birch" style="display:block">BIRCH</span></span>`}
        </span></p>`;
      await page.setContent(`<style>p{font:10pt/12pt Arial;margin:0}</style><div id="pagination-staging">
        <div id="pagination-hf-registry" style="display:none">
          <div data-section="0" data-hf-type="${placement === 'unused-header' ? 'header-even' : 'header-default'}">
            ${placement === 'body' ? '' : story}
          </div>
        </div>
        <div data-section-index="0" data-page-width="612" data-page-height="792"
          data-content-width="468" data-content-height="648" data-margin-top="72" data-margin-bottom="72"
          data-margin-left="72" data-margin-right="72" data-header-height="18" data-footer-height="18">
          <p data-source-anchor-id="p:body:control">Grove register control</p>${placement === 'body' ? story : ''}
        </div></div><div id="pagination-container"></div>`);
      await page.addScriptTag({ path: path.join(directory, '../dist/pagination.bundle.js') });
      const result = await page.evaluate(() => {
        try {
          const engine = new (window as any).DocxodusPagination.PaginationEngine(
            document.getElementById('pagination-staging'), document.getElementById('pagination-container'),
            { showPageNumbers: false, layoutToken: { documentVersion: 0, rendererFingerprint: 'source-inventory' } });
          const map = engine.paginate().pageMap;
          return { error: '', anchors: map.fragments.map((f: any) => f.anchorId) as string[] };
        } catch (error) { return { error: String(error), anchors: [] as string[] }; }
      });
      if (omitted && placement !== 'unused-header') {
        expect(result.error).toContain(`p:${scope}:birch`);
        expect(result.anchors).toEqual([]);
      } else {
        expect(result.error).toBe('');
        expect(result.anchors).toContain('p:body:control');
        if (placement !== 'unused-header') expect(result.anchors).toContain(`p:${scope}:birch`);
      }
    });
  }
}
