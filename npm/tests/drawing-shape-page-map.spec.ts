import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { SHAPE_VARIANTS, shapeDocx } from './docx-drawing-shape-fixture.js';

/**
 * A DrawingML rectangle (alone, or in a group) that is not a text box used to vanish from the HTML,
 * leaving its paragraph with no line box, so the PageMap found no measurable fragment for it and PDF
 * export failed with output_verification_failure (issue #1270). The rectangle now renders at its
 * 2 × 1 inch extent, inside the deletion when it is a tracked deletion, and every paragraph is mapped.
 * export-drawing-shape.spec.ts covers the PDF export itself.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const POINTS_PER_PIXEL = 72 / 96;

interface Result {
  error: string | null;
  sourceAnchors: string[];
  mapped: string[];
  registration: { success: boolean } | null;
  shape: { width: number; height: number; fill: string | null; inDeletion: boolean } | null;
}

async function paginate(page: Page, bytes: Uint8Array): Promise<Result> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const D = (window as any).Docxodus;
    const bin = new Uint8Array(input);
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      bin, 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, /* renderTrackedChanges */ true, /* showDeletedContent */ true, false, false, null, false,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    const sourceAnchors = Array.from(new Set(Array.from(
      staging.querySelectorAll<HTMLElement>('[data-source-anchor-id]'),
    ).map((element) => element.dataset.sourceAnchorId!)));
    try {
      const pageMap = new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
        showPageNumbers: false,
        fragmentParagraphs: true,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'drawing-shape' },
      }).paginate().pageMap;
      const bridge = D.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      let registration;
      try {
        registration = JSON.parse(bridge.RegisterPageMap(handle, JSON.stringify(pageMap), 'drawing-shape'));
      } finally {
        bridge.CloseSession(handle);
      }
      const svg = container.querySelector('svg');
      const box = svg?.getBoundingClientRect();
      return {
        error: null,
        sourceAnchors,
        mapped: pageMap.fragments.map((fragment: { anchorId: string }) => fragment.anchorId),
        registration,
        shape: svg && box ? {
          width: box.width,
          height: box.height,
          fill: svg.querySelector('rect')?.getAttribute('fill') ?? null,
          inDeletion: svg.closest('del') !== null,
        } : null,
      };
    } catch (error) {
      return { error: String(error), sourceAnchors, mapped: [], registration: null, shape: null };
    }
  }, Array.from(bytes));
}

test.describe('a DrawingML shape or group in its own paragraph', () => {
  for (const variant of SHAPE_VARIANTS) {
    test(`${variant.name}: paginates with the rectangle at its extent and every paragraph mapped`, async ({ page }) => {
      const result = await paginate(page, shapeDocx(variant));

      expect(result.error).toBeNull();
      expect(result.sourceAnchors).toHaveLength(3);
      for (const anchor of result.sourceAnchors) expect(result.mapped).toContain(anchor);
      expect(result.registration?.success, JSON.stringify(result.registration)).toBe(true);
      expect(result.shape).not.toBeNull();
      expect(result.shape!.width * POINTS_PER_PIXEL).toBeCloseTo(144, 0);
      expect(result.shape!.height * POINTS_PER_PIXEL).toBeCloseTo(72, 0);
      expect(result.shape!.fill).toBe('#19647E');
      expect(result.shape!.inDeletion).toBe(variant.deleted);
    });
  }
});
