import { expect, Page, test } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function addBundle(page: Page): Promise<void> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
}

function shell(content: string): string {
  return `<div id="pagination-staging">${content}</div><div id="pagination-container"></div>`;
}

const words = Array.from({ length: 90 }, (_, i) => `word${i}`).join(' ');

test.describe('PageMap materialization and citation navigation', () => {
  test('the public helper returns scale-independent paragraph fragments and navigates them', async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await addBundle(page);

    const run = async (scale: number) => page.evaluate(({ html, scale }) => {
      const api = (window as any).DocxodusPagination;
      const viewer = document.getElementById('viewer') as HTMLElement;
      const result = api.paginateHtml(html, viewer, {
        scale,
        showPageNumbers: false,
        layoutToken: { documentVersion: 7, rendererFingerprint: 'chromium-layout-v1' },
      });
      const map = result.pageMap;
      const citation = {
        availability: 'available',
        anchorId: 'p:body:shared-unid',
        fragments: map.fragments.filter((f: any) => f.anchorId === 'p:body:shared-unid'),
      };
      const navigation = api.navigateToPageCitation(viewer, citation, { behavior: 'auto' });
      return {
        totalPages: result.totalPages,
        map: {
          schemaVersion: map.schemaVersion,
          documentVersion: map.documentVersion,
          rendererFingerprint: map.rendererFingerprint,
          pages: map.pages,
          fragments: citation.fragments,
        },
        active: viewer.querySelectorAll(
          '[data-source-anchor-id="p:body:shared-unid"][data-anchor]',
        ).length,
        navigated: navigation.navigated,
        highlighted: navigation.target?.style.outline.includes('solid') ?? false,
      };
    }, {
      scale,
      html: shell(`
        <div data-section-index="0"
             data-page-width="122" data-page-height="58"
             data-content-width="120" data-content-height="56"
             data-margin-top="1" data-margin-right="1"
             data-margin-bottom="1" data-margin-left="1">
          <p data-anchor="shared-unid" data-source-anchor-id="p:body:shared-unid"
             style="font: 10pt/10pt Arial; margin:0">${words}</p>
        </div>`),
    });

    const normal = await run(1);
    const scaled = await run(0.55);
    expect(normal.totalPages).toBeGreaterThan(1);
    expect(normal.map.schemaVersion).toBe(1);
    expect(normal.map.documentVersion).toBe(7);
    expect(normal.map.rendererFingerprint).toBe('chromium-layout-v1');
    expect(normal.map.fragments.length).toBeGreaterThan(1);
    expect(normal.map.fragments.map((f: any) => f.fragmentIndex))
      .toEqual(normal.map.fragments.map((_: any, i: number) => i));
    expect(normal.active).toBe(1);
    expect(normal.navigated).toBe(true);
    expect(normal.highlighted).toBe(true);

    expect(scaled.map.pages.map((p: any) => [p.width, p.height]))
      .toEqual(normal.map.pages.map((p: any) => [p.width, p.height]));
    for (let i = 0; i < normal.map.fragments.length; i++) {
      const a = normal.map.fragments[i].geometry;
      const b = scaled.map.fragments[i].geometry;
      expect(b.x).toBeCloseTo(a.x, 1);
      expect(b.y).toBeCloseTo(a.y, 1);
      expect(b.width).toBeCloseTo(a.width, 1);
      expect(b.height).toBeCloseTo(a.height, 1);
    }
  });

  test('maps split tables, repeated stories, continued notes, comments, and collision-safe identities', async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await addBundle(page);

    const result = await page.evaluate((html) => {
      const api = (window as any).DocxodusPagination;
      const viewer = document.getElementById('viewer') as HTMLElement;
      const pagination = api.paginateHtml(html, viewer, {
        showPageNumbers: false,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'stories-v1' },
      });
      const map = pagination.pageMap;
      const activeByCanonical = Array.from(viewer.querySelectorAll<HTMLElement>('[data-source-anchor-id]'))
        .filter((node) => node.hasAttribute('data-anchor'))
        .reduce<Record<string, number>>((counts, node) => {
          const id = node.dataset.sourceAnchorId!;
          counts[id] = (counts[id] ?? 0) + 1;
          return counts;
        }, {});
      const storyPages = (story: string) => Array.from(new Set(
        map.fragments.filter((f: any) => f.story === story).map((f: any) => f.pageNumber),
      ));
      return {
        totalPages: pagination.totalPages,
        fragments: map.fragments,
        storyPages: {
          header: storyPages('header'),
          footer: storyPages('footer'),
          footnote: storyPages('footnote'),
          endnote: storyPages('endnote'),
          comment: storyPages('comment'),
        },
        activeByCanonical,
        activeBareIds: Array.from(viewer.querySelectorAll<HTMLElement>('[data-anchor]'))
          .map((node) => node.dataset.anchor),
        sharedBareSources: Array.from(viewer.querySelectorAll<HTMLElement>('[data-anchor="same"]'))
          .map((node) => node.dataset.sourceAnchorId),
      };
    }, shell(`
      <div id="pagination-hf-registry" style="display:none">
        <div data-section="0" data-hf-type="header-default">
          <p data-anchor="same" data-source-anchor-id="p:hdr1:same" style="margin:0">HEADER</p>
        </div>
        <div data-section="0" data-hf-type="footer-default">
          <p data-anchor="footer" data-source-anchor-id="p:ftr1:footer" style="margin:0">FOOTER</p>
        </div>
      </div>
      <div data-section-index="0" data-header-distance="1" data-footer-distance="1"
           data-page-width="160" data-page-height="135"
           data-content-width="158" data-content-height="108"
           data-margin-top="13" data-margin-right="1"
           data-margin-bottom="13" data-margin-left="1">
        <p data-anchor="same" data-source-anchor-id="p:body:same" style="height:18pt;margin:0">
          body
        </p>
        <div><table data-anchor="table" data-source-anchor-id="tbl:body:table"><tbody>
          ${Array.from({ length: 5 }, (_, i) => `
            <tr data-source-anchor-id="tr:body:r${i}" style="height:22pt">
              <td data-source-anchor-id="tc:body:c${i}">
                <p data-anchor="cp${i}" data-source-anchor-id="p:body:cp${i}" style="margin:0">cell ${i}</p>
              </td>
            </tr>`).join('')}
        </tbody></table></div>
        <p data-anchor="comment" data-source-anchor-id="p:cmt:comment"
           style="height:18pt;margin:0">visible comment body</p>
        <p data-anchor="note" data-source-anchor-id="p:fn:note"
           style="font:10pt/10pt Arial;margin:0">continued footnote ${words}</p>
        <p data-anchor="endnote" data-source-anchor-id="p:en:endnote"
           style="font:10pt/10pt Arial;margin:0">continued endnote ${words}</p>
      </div>`));

    expect(result.totalPages).toBeGreaterThan(2);
    expect(result.storyPages.header.length).toBe(result.totalPages);
    expect(result.storyPages.footer.length).toBe(result.totalPages);
    expect(result.storyPages.footnote.length).toBeGreaterThan(1);
    expect(result.storyPages.endnote.length).toBeGreaterThan(1);
    expect(result.storyPages.comment.length).toBe(1);

    const tableFragments = result.fragments.filter((f: any) => f.anchorId === 'tbl:body:table');
    expect(tableFragments.length).toBeGreaterThan(1);
    const cell = result.fragments.find((f: any) => f.anchorId === 'tc:body:c0');
    expect(cell.inTableCell).toBe(true);
    expect(result.fragments.some((f: any) => f.anchorId.startsWith('tr:body:'))).toBe(true);
    const indicesByAnchor = new Map<string, number[]>();
    for (const fragment of result.fragments) {
      const indices = indicesByAnchor.get(fragment.anchorId) ?? [];
      indices.push(fragment.fragmentIndex);
      indicesByAnchor.set(fragment.anchorId, indices);
    }
    for (const indices of indicesByAnchor.values()) {
      expect(indices).toEqual(indices.map((_: number, index: number) => index));
    }

    for (const count of Object.values(result.activeByCanonical)) expect(count).toBe(1);
    expect(result.activeByCanonical['p:body:same']).toBe(1);
    expect(result.activeByCanonical['p:hdr1:same']).toBeUndefined();
    expect(result.activeByCanonical['p:fn:note']).toBe(1);
    expect(new Set(result.activeBareIds).size).toBe(result.activeBareIds.length);
    expect(result.sharedBareSources).toEqual(['p:body:same']);
  });

  test('captures columns and mixed page names/sizes', async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await addBundle(page);

    const result = await page.evaluate((html) => {
      const api = (window as any).DocxodusPagination;
      const pagination = api.paginateHtml(html, 'viewer', {
        showPageNumbers: false,
        layoutToken: { documentVersion: 3, rendererFingerprint: 'mixed-v1' },
      });
      return {
        pages: pagination.pageMap.pages,
        columnFragments: pagination.pageMap.fragments
          .filter((f: any) => f.anchorId.startsWith('p:body:column-')),
      };
    }, shell(`
      <div data-section-index="0" data-cols="2" data-col-gap="8"
           data-page-width="122" data-page-height="80"
           data-content-width="120" data-content-height="78"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        ${Array.from({ length: 8 }, (_, i) =>
          `<p data-anchor="column-${i}" data-source-anchor-id="p:body:column-${i}"
              style="height:18pt;margin:0">column ${i}</p>`).join('')}
      </div>
      <div data-section-index="1" data-page-break-before="true"
           data-page-width="200" data-page-height="140"
           data-content-width="198" data-content-height="138"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p data-anchor="wide" data-source-anchor-id="p:body:wide" style="height:18pt;margin:0">wide</p>
      </div>`));

    expect(result.columnFragments.length).toBe(8);
    expect(new Set(result.pages.map((p: any) => p.pageName)))
      .toEqual(new Set(['docxodus-section-0', 'docxodus-section-1']));
    expect(result.pages.some((p: any) => p.width === 122 && p.height === 80)).toBe(true);
    expect(result.pages.some((p: any) => p.width === 200 && p.height === 140)).toBe(true);
    expect(result.pages.every((p: any) => p.pageInSection >= 1)).toBe(true);
  });

  test('refuses to publish an available map with an unmeasurable addressable block', async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await addBundle(page);
    const message = await page.evaluate((html) => {
      try {
        (window as any).DocxodusPagination.paginateHtml(html, 'viewer', {
          showPageNumbers: false,
          layoutToken: { documentVersion: 0, rendererFingerprint: 'strict-v1' },
        });
        return '';
      } catch (error) {
        return error instanceof Error ? error.message : String(error);
      }
    }, shell(`
      <div data-section-index="0"
           data-page-width="122" data-page-height="80"
           data-content-width="120" data-content-height="78"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p data-anchor="zero" data-source-anchor-id="p:body:zero"
           style="width:0;height:0;padding:0;border:0;margin:0;overflow:hidden"></p>
      </div>`));
    expect(message).toContain('has no measurable fragment');
  });

  test('refuses a map when an addressable staging block never reaches a page', async ({ page }) => {
    await page.setContent(shell(`
      <div data-section-index="0"
           data-page-width="122" data-page-height="80"
           data-content-width="120" data-content-height="78"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p data-anchor="same" data-source-anchor-id="p:body:kept"
           style="height:18pt;margin:0">kept</p>
        <p data-anchor="same" data-source-anchor-id="p:body:dropped"
           data-drop-from-flow="true" style="height:18pt;margin:0">dropped</p>
      </div>`));
    await addBundle(page);

    const message = await page.evaluate(() => {
      const api = (window as any).DocxodusPagination;
      const engine = new api.PaginationEngine('pagination-staging', 'pagination-container', {
        showPageNumbers: false,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'dropped-source-v1' },
      });
      const originalMeasureBlocks = engine.measureBlocks.bind(engine);
      engine.measureBlocks = (section: HTMLElement, dimensions: unknown) =>
        originalMeasureBlocks(section, dimensions)
          .filter((block: any) => block.element.dataset.dropFromFlow !== 'true');
      try {
        engine.paginate();
        return '';
      } catch (error) {
        return error instanceof Error ? error.message : String(error);
      }
    });

    expect(message).toContain('source anchor p:body:dropped has no measurable fragment');
  });
});
