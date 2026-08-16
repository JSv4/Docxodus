import { expect, Page, test } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { generateFootnoteDocx } from './docx-footnote-fixture.js';
import { generateLongEndnoteDocx, generateTableCommentDocx } from './docx-page-map-fixture.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function readyPage(page: Page): Promise<void> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
}

interface RealPaginationResult {
  htmlHasBareAnchors: boolean;
  htmlCanonicalCount: number;
  pages: number;
  fragments: Array<{
    fragmentId: string;
    anchorId: string;
    fragmentIndex: number;
    pageNumber: number;
    story: string;
    inTableCell: boolean;
    geometry: { x: number; y: number; width: number; height: number };
  }>;
  registration: { success: boolean; error?: string; message?: string };
  commentSectionInStaging: boolean;
  marginRegistryInStaging: boolean;
  marginColumns: number;
  duplicatePageCommentIds: number;
  activeMarginBackrefs: number;
  endnoteReferenceCount: number;
  endnoteTargetCount: number;
  endnoteReferenceResolves: boolean;
  endnoteLinkTargetsUnique: boolean;
  endnoteMarkerText: string | null;
  visibleFootnoteText: string;
  clippedFootnoteParagraphs: Array<{
    anchorId: string | null;
    pageNumber: number;
    paragraphBottom: number;
    bandBottom: number;
  }>;
}

async function convertPaginateAndRegister(
  page: Page,
  docx: Uint8Array,
  commentMode: number,
  renderNotes: boolean,
  fingerprint: string,
): Promise<RealPaginationResult> {
  return page.evaluate(({ bytes, commentMode, renderNotes, fingerprint }) => {
    const D = (window as any).Docxodus;
    const bin = new Uint8Array(bytes);
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      bin, 'Document', 'docx-', false, '', commentMode, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      renderNotes, false, false, false, false, false, null,
      /* stampAnchors */ false,
    );
    if (html.startsWith('{')) throw new Error(`conversion failed: ${html.slice(0, 300)}`);

    const host = document.createElement('div');
    host.className = 'real-page-map-host';
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    const engine = new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
      showPageNumbers: false,
      fragmentParagraphs: true,
      layoutToken: { documentVersion: 0, rendererFingerprint: fingerprint },
    });
    const pagination = engine.paginate();
    const footnoteBands = Array.from(
      container.querySelectorAll<HTMLElement>('.page-footnotes'),
    );
    const visibleFootnoteText = footnoteBands.map((band) => {
      const clone = band.cloneNode(true) as HTMLElement;
      clone.querySelectorAll('.footnote-number').forEach((number) => number.remove());
      clone.querySelectorAll('p').forEach((paragraph) => paragraph.after(' '));
      return clone.textContent ?? '';
    }).join(' ').replace(/\s+/g, ' ').trim();
    const clippedFootnoteParagraphs = footnoteBands.flatMap((band) => {
      const bandRect = band.getBoundingClientRect();
      const pageNumber = Number(
        band.closest<HTMLElement>('.page-box')?.dataset.pageNumber ?? '0',
      );
      return Array.from(band.querySelectorAll<HTMLElement>('p'))
        .filter((paragraph) => {
          const rect = paragraph.getBoundingClientRect();
          return rect.top < bandRect.top - 1 || rect.bottom > bandRect.bottom + 1;
        })
        .map((paragraph) => ({
          anchorId: paragraph.dataset.sourceAnchorId ?? null,
          pageNumber,
          paragraphBottom: paragraph.getBoundingClientRect().bottom,
          bandBottom: bandRect.bottom,
        }));
    });

    const bridge = D.DocxSessionBridge;
    const handle = bridge.OpenSession(bin, '');
    let registration: { success: boolean; error?: string; message?: string };
    try {
      registration = JSON.parse(bridge.RegisterPageMap(
        handle, JSON.stringify(pagination.pageMap), fingerprint,
      ));
    } finally {
      bridge.CloseSession(handle);
    }

    return {
      htmlHasBareAnchors: /\bdata-anchor=/.test(html),
      htmlCanonicalCount: (html.match(/\bdata-source-anchor-id=/g) || []).length,
      pages: pagination.totalPages,
      fragments: pagination.pageMap.fragments,
      registration,
      commentSectionInStaging: Array.from(staging.querySelectorAll<HTMLElement>(
        '[data-section-index] aside',
      )).some((node) => node.className.includes('comments-section')),
      marginRegistryInStaging: staging.querySelector(
        '#pagination-comment-margin-registry',
      ) !== null,
      marginColumns: container.querySelectorAll('.page-comment-margin').length,
      duplicatePageCommentIds: container.querySelectorAll('.page-comment-margin [id]').length,
      activeMarginBackrefs: container.querySelectorAll(
        '.page-comment-margin a[href^="#"]',
      ).length,
      endnoteReferenceCount: container.querySelectorAll(
        'a[href^="#en-"]:not([href^="#en-ref-"])',
      ).length,
      endnoteReferenceResolves: Array.from(container.querySelectorAll<HTMLAnchorElement>(
        'a[href^="#en-"]',
      )).every((link) => {
        const target = document.querySelector(link.getAttribute('href')!);
        return target !== null && container.contains(target);
      }),
      endnoteLinkTargetsUnique: Array.from(container.querySelectorAll<HTMLAnchorElement>(
        'a[href^="#en-"]',
      )).every((link) => document.querySelectorAll(
        `[id="${CSS.escape(link.getAttribute('href')!.slice(1))}"]`,
      ).length === 1),
      endnoteTargetCount: document.querySelectorAll('#en-1').length,
      endnoteMarkerText: container.querySelector<HTMLElement>('#en-1')?.textContent?.trim()
        ?? null,
      visibleFootnoteText,
      clippedFootnoteParagraphs,
    };
  }, {
    bytes: Array.from(docx),
    commentMode,
    renderNotes,
    fingerprint,
  });
}

test.describe('Real converter PageMap pipeline', () => {
  test.beforeEach(async ({ page }) => readyPage(page));

  for (const mode of [
    { name: 'endnote-style', value: 0 },
    { name: 'inline', value: 1 },
    { name: 'margin', value: 2 },
  ]) {
    test(`default no-stamp ${mode.name} comments paginate and register`, async ({ page }) => {
      const result = await convertPaginateAndRegister(
        page,
        generateTableCommentDocx(),
        mode.value,
        false,
        `real-comment-${mode.name}-v1`,
      );
      const comments = result.fragments.filter((fragment) => fragment.story === 'comment');
      expect(result.htmlHasBareAnchors).toBe(false);
      expect(result.htmlCanonicalCount).toBeGreaterThan(0);
      expect(result.registration.success, JSON.stringify(result.registration)).toBe(true);
      expect(comments.length).toBeGreaterThanOrEqual(3); // cmt + two p:cmt definitions

      if (mode.value === 0) {
        expect(result.commentSectionInStaging).toBe(true);
        expect(comments.every((fragment) => !fragment.inTableCell)).toBe(true);
      } else if (mode.value === 1) {
        expect(comments.every((fragment) => fragment.inTableCell)).toBe(true);
      } else {
        expect(result.marginRegistryInStaging).toBe(true);
        expect(result.marginColumns).toBeGreaterThan(0);
        expect(comments.every((fragment) => !fragment.inTableCell)).toBe(true);
        expect(result.duplicatePageCommentIds).toBe(0);
        expect(result.activeMarginBackrefs).toBe(0);
      }
    });
  }

  for (const mode of [
    { name: 'inline', value: 1 },
    { name: 'margin', value: 2 },
  ]) {
    test(`collapsed ${mode.name} comment maps its visible reference`, async ({ page }) => {
      const result = await convertPaginateAndRegister(
        page,
        generateTableCommentDocx(true),
        mode.value,
        false,
        `collapsed-comment-${mode.name}-v1`,
      );
      const comments = result.fragments.filter((fragment) => fragment.story === 'comment');
      expect(result.registration.success, JSON.stringify(result.registration)).toBe(true);
      expect(comments.length).toBeGreaterThanOrEqual(3);
      if (mode.value === 1) {
        expect(comments.every((fragment) => fragment.inTableCell)).toBe(true);
      } else {
        expect(result.marginColumns).toBeGreaterThan(0);
        expect(comments.every((fragment) => !fragment.inTableCell)).toBe(true);
      }
    });
  }

  test('split real footnote keeps its fn definition identity on every continuation page', async ({ page }) => {
    const result = await convertPaginateAndRegister(
      page,
      generateFootnoteDocx(1, 1, 90),
      -1,
      true,
      'real-footnote-continuation-v1',
    );
    const definitions = result.fragments.filter((fragment) =>
      fragment.story === 'footnote' && fragment.anchorId.startsWith('fn:fn:'));
    const paragraphDefinitions = result.fragments.filter((fragment) =>
      fragment.story === 'footnote' && fragment.anchorId.startsWith('p:fn:'));
    expect(result.registration.success, JSON.stringify(result.registration)).toBe(true);
    expect(new Set(definitions.map((fragment) => fragment.pageNumber)).size).toBeGreaterThan(1);
    expect(new Set(paragraphDefinitions.map((fragment) => fragment.pageNumber)).size).toBeGreaterThan(1);
    expect(result.fragments.every((fragment) =>
      fragment.geometry.x >= 0
      && fragment.geometry.y >= 0
      && fragment.geometry.width > 0
      && fragment.geometry.height > 0)).toBe(true);
  });

  test('one long real footnote paragraph has contiguous PageMap fragments', async ({
    page,
  }, testInfo) => {
    let result: RealPaginationResult | undefined;
    let failure: { name?: string; message: string; stack?: string } | undefined;
    try {
      result = await convertPaginateAndRegister(
        page,
        generateFootnoteDocx(1, 1, 1, 600),
        -1,
        true,
        'real-long-footnote-paragraph-v1',
      );
      const paragraphs = result.fragments.filter((fragment) =>
        fragment.story === 'footnote' && fragment.anchorId.startsWith('p:fn:'));
      const expectedText = Array.from(
        { length: 600 },
        (_, index) => `footnote-1-1-${index + 1}`,
      ).join(' ');

      expect(result.registration.success, JSON.stringify(result.registration)).toBe(true);
      expect(result.pages).toBeGreaterThan(1);
      expect(paragraphs.length).toBeGreaterThan(1);
      expect(new Set(paragraphs.map((fragment) => fragment.anchorId)).size).toBe(1);
      expect(new Set(paragraphs.map((fragment) => fragment.pageNumber)).size).toBeGreaterThan(1);
      expect(paragraphs.map((fragment) => fragment.fragmentIndex))
        .toEqual(paragraphs.map((_, index) => index));
      expect(new Set(paragraphs.map((fragment) => fragment.fragmentId)).size)
        .toBe(paragraphs.length);
      expect(paragraphs.every((fragment) =>
        fragment.geometry.x >= 0
        && fragment.geometry.y >= 0
        && fragment.geometry.width > 0
        && fragment.geometry.height > 0)).toBe(true);
      expect(result.visibleFootnoteText).toBe(expectedText);
      expect(result.clippedFootnoteParagraphs).toEqual([]);
    } catch (error) {
      failure = error instanceof Error
        ? { name: error.name, message: error.message, stack: error.stack }
        : { message: String(error) };
      throw error;
    } finally {
      const captureWarnings: string[] = [];
      try {
        await testInfo.attach('issue-489-real-long-footnote.html', {
          body: Buffer.from(await page.content()),
          contentType: 'text/html',
        });
      } catch (error) {
        captureWarnings.push(`HTML capture failed: ${String(error)}`);
      }
      try {
        const host = page.locator('.real-page-map-host');
        if (await host.count()) {
          await testInfo.attach('issue-489-real-long-footnote.png', {
            body: await host.screenshot(),
            contentType: 'image/png',
          });
        }
      } catch (error) {
        captureWarnings.push(`screenshot capture failed: ${String(error)}`);
      }
      await testInfo.attach('issue-489-real-long-footnote.json', {
        body: Buffer.from(JSON.stringify({ result, failure, captureWarnings }, null, 2)),
        contentType: 'application/json',
      });
    }
  });

  test('one long real endnote paragraph fragments across final flow pages', async ({ page }) => {
    const result = await convertPaginateAndRegister(
      page,
      generateLongEndnoteDocx(),
      -1,
      true,
      'real-long-endnote-v1',
    );
    const definitions = result.fragments.filter((fragment) =>
      fragment.story === 'endnote' && fragment.anchorId.startsWith('en:en:'));
    const paragraphs = result.fragments.filter((fragment) =>
      fragment.story === 'endnote' && fragment.anchorId.startsWith('p:en:'));
    expect(result.registration.success, JSON.stringify(result.registration)).toBe(true);
    expect(result.pages).toBeGreaterThan(1);
    expect(
      new Set(definitions.map((fragment) => fragment.pageNumber)).size,
      JSON.stringify({ pages: result.pages, definitions, paragraphs }),
    ).toBeGreaterThan(1);
    expect(new Set(paragraphs.map((fragment) => fragment.pageNumber)).size).toBeGreaterThan(1);
    expect(result.endnoteReferenceCount).toBeGreaterThan(0);
    expect(result.endnoteTargetCount).toBe(1);
    expect(result.endnoteReferenceResolves).toBe(true);
    expect(result.endnoteLinkTargetsUnique).toBe(true);
    expect(result.endnoteMarkerText).toMatch(/^i\.\s/);
  });
});
