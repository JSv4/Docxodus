import { expect, test } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const bundle = join(dirname(fileURLToPath(import.meta.url)), '../dist/pagination.bundle.js');

for (const kind of ['blank', 'visible', 'unmarked', 'unclipped'] as const) {
  test(`textbox ${kind} paragraph keeps the source completeness contract`, async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(kind => {
      const html = `<div id="pagination-staging"><div data-section-index="0"
        data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
        data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
        <p data-source-anchor-id="p:body:host" style="font:8pt/12pt Arial;margin:0">
          <span ${kind === 'unmarked' ? '' : 'data-docx-textbox-clip="true"'}
            style="display:inline-block;width:144pt;height:12pt;overflow:${kind === 'unclipped' ? 'visible' : 'hidden'};vertical-align:top">
            <span data-source-anchor-id="p:body:label" style="display:block;font:8pt/12pt Arial">Sighting logged.</span>
            <span data-source-anchor-id="p:body:last" style="display:block;font:8pt/12pt Arial">${kind === 'visible' ? 'Off-box signal.' : '&nbsp;'}</span>
          </span>
        </p><p data-source-anchor-id="p:body:caption" style="font:8pt/12pt Arial;margin:0">Orbit verified.</p>
      </div></div><div id="pagination-container"></div>`;
      try {
        const result = (window as any).DocxodusPagination.paginateHtml(html, 'viewer', {
          showPageNumbers: false,
          layoutToken: { documentVersion: 0, rendererFingerprint: 'textbox-clip' },
        });
        return { error: null, anchors: result.pageMap.fragments.map((f: any) => f.anchorId) };
      } catch (error) { return { error: String(error), anchors: [] }; }
    }, kind);
    if (kind === 'visible' || kind === 'unmarked') {
      expect(result.error).toContain('source anchor p:body:last has no measurable fragment');
    } else {
      expect(result.error).toBeNull();
      expect(result.anchors).toContain('p:body:label');
      expect(result.anchors).toContain('p:body:caption');
      if (kind === 'blank') expect(result.anchors).not.toContain('p:body:last');
      else expect(result.anchors).toContain('p:body:last');
    }
  });
}
