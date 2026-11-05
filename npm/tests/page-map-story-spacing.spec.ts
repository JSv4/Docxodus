import { expect, test } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const bundle = join(dirname(fileURLToPath(import.meta.url)), '../dist/pagination.bundle.js');

for (const story of ['header', 'footer'] as const) {
  test(`${story} spacing reserves every paragraph's actual margin box`, async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(story => {
      const scope = story === 'header' ? 'hdr1' : 'ftr1';
      const html = `<div id="pagination-staging">
        <div id="pagination-hf-registry" style="display:none">
          <div data-section="0" data-hf-type="${story}-default">
            ${Array.from({ length: 4 }, (_, i) => `<p data-source-anchor-id="p:${scope}:label${i}"
              style="font:10pt/12pt Arial;margin:18pt 0">Observatory label ${i}</p>`).join('')}
          </div>
        </div>
        <div data-section-index="0" data-page-width="300" data-page-height="360"
          data-content-width="228" data-content-height="288"
          data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36"
          data-header-distance="12" data-footer-distance="12">
          <p data-source-anchor-id="p:body:caption" style="font:10pt/12pt Arial;margin:0">Night log caption.</p>
        </div>
      </div><div id="pagination-container"></div>`;
      const pagination = (window as any).DocxodusPagination.paginateHtml(html, 'viewer', {
        showPageNumbers: false,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'story-spacing' },
      });
      const band = document.querySelector<HTMLElement>(`.page-${story}`)!;
      const bounds = band.getBoundingClientRect();
      return {
        fragments: pagination.pageMap.fragments,
        height: bounds.height,
        contained: Array.from(band.children).every(child => {
          const rect = child.getBoundingClientRect();
          return rect.top >= bounds.top && rect.bottom <= bounds.bottom;
        }),
      };
    }, story);

    expect(result.contained).toBe(true);
    expect(result.height).toBeCloseTo(4 * (12 + 18 + 18) * 4 / 3, 1);
    expect(result.fragments.filter((fragment: any) => fragment.story === story)).toHaveLength(4);
    expect(result.fragments.some((fragment: any) => fragment.anchorId === 'p:body:caption')).toBe(true);
  });
}
