import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { outOfFlowDocx, textBoxRun } from './docx-out-of-flow-fixture.js';

/**
 * A paragraph whose only content is an anchored text box or shape is left with an empty,
 * zero-height line in the paginated view: the paginator promotes the drawing out of the paragraph
 * into the page box. The PageMap measured every source anchor by its own border box, found nothing
 * measurable for such a host paragraph, and failed the whole layout ("source anchor … has no
 * measurable fragment"), which the PDF export reported as output_verification_failure (issue #849).
 * The host's rendered content is the drawing, so that is where its fragment is.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

interface Result {
  error: string | null;
  sourceAnchors: string[];
  fragments: Array<{ anchorId: string; geometry: { x: number; y: number; width: number; height: number } }>;
  registration: { success: boolean; error?: string; message?: string } | null;
}

async function paginateWithPageMap(page: Page, bytes: Uint8Array): Promise<Result> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const D = (window as any).Docxodus;
    const bin = new Uint8Array(input);
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      bin, 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, false, false, false, false, null, false,
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
      const engine = new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
        showPageNumbers: false,
        fragmentParagraphs: true,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'out-of-flow' },
      });
      const pageMap = engine.paginate().pageMap;
      const bridge = D.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      let registration;
      try {
        registration = JSON.parse(bridge.RegisterPageMap(handle, JSON.stringify(pageMap), 'out-of-flow'));
      } finally {
        bridge.CloseSession(handle);
      }
      return { error: null, sourceAnchors, fragments: pageMap.fragments, registration };
    } catch (error) {
      return { error: String(error), sourceAnchors, fragments: [], registration: null };
    }
  }, Array.from(bytes));
}

function contains(
  outer: { x: number; y: number; width: number; height: number },
  inner: { x: number; y: number; width: number; height: number },
): boolean {
  const slack = 0.5;
  return inner.x >= outer.x - slack && inner.y >= outer.y - slack
    && inner.x + inner.width <= outer.x + outer.width + slack
    && inner.y + inner.height <= outer.y + outer.height + slack;
}

test.describe('PageMap for paragraphs whose only content is out of flow', () => {
  for (const verticalFrom of ['paragraph', 'page', 'margin']) {
    test(`a text-box-only paragraph (anchored to ${verticalFrom}) has a fragment enclosing its box`, async ({ page }) => {
      const result = await paginateWithPageMap(page, outOfFlowDocx(
        `<w:p><w:r><w:t>Before the box.</w:t></w:r></w:p>
         <w:p>${textBoxRun(verticalFrom, ['Inside the text box'])}</w:p>
         <w:p><w:r><w:t>After the box.</w:t></w:r></w:p>`,
      ));

      expect(result.error).toBeNull();
      const mapped = new Set(result.fragments.map((fragment) => fragment.anchorId));
      for (const anchor of result.sourceAnchors) expect(mapped).toContain(anchor);
      expect(result.registration?.success, JSON.stringify(result.registration)).toBe(true);

      // Document order: the paragraph before, the host, the text-box paragraph nested in the host,
      // the paragraph after. The host's fragment is where its content renders, around the box's.
      expect(result.sourceAnchors).toHaveLength(4);
      const [, hostAnchor, boxAnchor] = result.sourceAnchors;
      const host = result.fragments.find((fragment) => fragment.anchorId === hostAnchor)!;
      const box = result.fragments.find((fragment) => fragment.anchorId === boxAnchor)!;
      expect(contains(host.geometry, box.geometry), JSON.stringify({ host, box })).toBe(true);
    });
  }

  test('a text box holding an empty paragraph still maps every source paragraph', async ({ page }) => {
    const result = await paginateWithPageMap(page, outOfFlowDocx(
      `<w:p>${textBoxRun('paragraph', ['First line in the box', '', 'Third line in the box'])}</w:p>
       <w:p><w:r><w:t>Body text.</w:t></w:r></w:p>`,
    ));

    expect(result.error).toBeNull();
    const mapped = new Set(result.fragments.map((fragment) => fragment.anchorId));
    for (const anchor of result.sourceAnchors) expect(mapped).toContain(anchor);
    expect(result.registration?.success, JSON.stringify(result.registration)).toBe(true);
  });
});
