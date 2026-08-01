import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { fileURLToPath } from 'url';
import type { EditErrorCode } from '../src/types.js';

/**
 * The annotation is the point: it states that the .NET enum member reached the TypeScript union,
 * which asserting the literal at runtime cannot show (the parsed JSON is untyped, so a consumer
 * narrowing on `error.code === "invalid_page_numbering"` would fail to compile while this spec
 * still passed).
 *
 * Note it is NOT checked by any command that runs today — `tsconfig.json` includes `src/**` only,
 * and Playwright strips types without checking them. Type-checking this file directly does enforce
 * it (`npx tsc --noEmit --strict --moduleResolution bundler tests/page-numbering.spec.ts`);
 * type-checking the whole `tests/` tree does not currently pass for unrelated reasons.
 */
const INVALID_PAGE_NUMBERING: EditErrorCode = 'invalid_page_numbering';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/** See editor-headerfooter.spec.ts — read one entry out of a DOCX without a zip dependency. */
function readZipEntry(zip: Buffer, name: string): string {
  const eocd = zip.lastIndexOf(Buffer.from('PK\x05\x06', 'latin1'));
  if (eocd < 0) throw new Error('not a zip');
  const count = zip.readUInt16LE(eocd + 10);
  let p = zip.readUInt32LE(eocd + 16);
  for (let i = 0; i < count; i++) {
    const nameLen = zip.readUInt16LE(p + 28);
    const extraLen = zip.readUInt16LE(p + 30);
    const commentLen = zip.readUInt16LE(p + 32);
    const entry = zip.toString('latin1', p + 46, p + 46 + nameLen);
    if (entry === name) {
      const method = zip.readUInt16LE(p + 10);
      const compSize = zip.readUInt32LE(p + 20);
      const localOff = zip.readUInt32LE(p + 42);
      const lNameLen = zip.readUInt16LE(localOff + 26);
      const lExtraLen = zip.readUInt16LE(localOff + 28);
      const start = localOff + 30 + lNameLen + lExtraLen;
      const data = zip.subarray(start, start + compSize);
      return (method === 0 ? data : zlib.inflateRawSync(data)).toString('utf8');
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error(`entry not found: ${name}`);
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** Independent roman-numeral rendering, so the expectations don't restate the implementation. */
function formatRoman(value: number): string {
  const table: Array<[number, string]> = [
    [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
  ];
  let rest = value;
  let out = '';
  for (const [n, glyph] of table) {
    while (rest >= n) {
      out += glyph;
      rest -= n;
    }
  }
  return out;
}

/**
 * Issue #277 — page-number formatting.
 *
 * Two independent layers: the SECTION's `w:pgNumType` (which number the section starts at and in
 * which format — Word's *Format Page Numbers…*), and a per-field `\*` general-formatting switch
 * that overrides it for one field. Plus the paginated preview, which is the only renderer here
 * that can show a *different* number on each page: a header/footer is authored once and cloned
 * onto every page, so a field's single cached result would otherwise read the same everywhere.
 */
test.describe('Section page numbering (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('SetPageNumbering writes w:pgNumType, GetSectionInfo reads it back, Clear removes it', async ({
    page,
  }) => {
    const res = await page.evaluate(() => {
      const B = (window as any).Docxodus.DocxSessionBridge;
      const h = B.OpenSession(B.CreateBlankDocx(), '{}');
      try {
        const proj = JSON.parse(B.Project(h));
        const body = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const before = JSON.parse(B.GetSectionInfo(h, body));
        const set = JSON.parse(
          B.SetPageNumbering(h, body, JSON.stringify({ start: 1, format: 'lowerRoman' })),
        );
        const after = JSON.parse(B.GetSectionInfo(h, body));

        // An omitted field leaves the other attribute alone.
        B.SetPageNumbering(h, body, JSON.stringify({ format: 'upperRoman' }));
        const merged = JSON.parse(B.GetSectionInfo(h, body));

        const savedWithNumbering: Uint8Array = B.Save(h);
        const cleared = JSON.parse(B.ClearPageNumbering(h, body));
        const afterClear = JSON.parse(B.GetSectionInfo(h, body));

        return {
          setSuccess: set.success,
          clearSuccess: cleared.success,
          beforeStart: before.pageNumberStart ?? null,
          beforeFormat: before.pageNumberFormat ?? null,
          afterStart: after.pageNumberStart ?? null,
          afterFormat: after.pageNumberFormat ?? null,
          mergedStart: merged.pageNumberStart ?? null,
          mergedFormat: merged.pageNumberFormat ?? null,
          clearedStart: afterClear.pageNumberStart ?? null,
          clearedFormat: afterClear.pageNumberFormat ?? null,
          savedWithNumbering: Array.from(savedWithNumbering),
        };
      } finally {
        B.CloseSession(h);
      }
    });

    expect(res.setSuccess).toBe(true);
    expect(res.clearSuccess).toBe(true);
    // Absent attributes read back as absent, not as a fabricated decimal/1 default.
    expect(res.beforeStart).toBeNull();
    expect(res.beforeFormat).toBeNull();
    expect(res.afterStart).toBe(1);
    expect(res.afterFormat).toBe('lowerRoman');
    expect(res.mergedStart).toBe(1);
    expect(res.mergedFormat).toBe('upperRoman');
    expect(res.clearedStart).toBeNull();
    expect(res.clearedFormat).toBeNull();

    const documentXml = readZipEntry(Buffer.from(res.savedWithNumbering), 'word/document.xml');
    expect(documentXml).toContain('w:pgNumType');
    expect(documentXml).toContain('w:fmt="upperRoman"');
  });

  test('a page-number field takes an optional \\* format switch; omitting it stays plain', async ({
    page,
  }) => {
    const res = await page.evaluate(() => {
      const B = (window as any).Docxodus.DocxSessionBridge;
      const h = B.OpenSession(B.CreateBlankDocx(), '{}');
      try {
        const proj = JSON.parse(B.Project(h));
        const body = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const footer = JSON.parse(B.SetFooterText(h, body, 'default', 'Page '));
        const anchor: string = footer.created[0].id;
        const switched = JSON.parse(
          B.InsertPageNumberField(h, anchor, 'currentPage', 'lowerRoman'),
        );

        const header = JSON.parse(B.SetHeaderText(h, body, 'default', ''));
        const plainAnchor: string = header.created[0].id;
        const plain = JSON.parse(B.InsertPageNumberField(h, plainAnchor, 'currentPage', ''));

        // Bullet is a list format with no page-number meaning — it must fail, not degrade.
        const bullet = JSON.parse(B.InsertPageNumberField(h, plainAnchor, 'currentPage', 'bullet'));

        return {
          switchedSuccess: switched.success,
          plainSuccess: plain.success,
          bulletSuccess: bullet.success,
          bulletCode: bullet.error?.code ?? null,
          saved: Array.from(B.Save(h) as Uint8Array),
        };
      } finally {
        B.CloseSession(h);
      }
    });

    expect(res.switchedSuccess).toBe(true);
    expect(res.plainSuccess).toBe(true);
    expect(res.bulletSuccess).toBe(false);
    expect(res.bulletCode).toBe(INVALID_PAGE_NUMBERING);

    const zip = Buffer.from(res.saved);
    const footerXml = readZipEntry(zip, 'word/footer1.xml');
    expect(footerXml).toContain('PAGE \\* roman');
    // The cached result agrees with the switch — it is what a non-recomputing renderer shows.
    expect(footerXml).toContain('<w:t>i</w:t>');

    const headerXml = readZipEntry(zip, 'word/header1.xml');
    expect(headerXml).toContain('PAGE');
    expect(headerXml).not.toContain('\\*');
  });
});

test.describe('DocxEditor — page-numbering band chrome', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('the band format/start controls write w:pgNumType and reflect the live section', async ({
    page,
  }) => {
    const res = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {
        headerFooter: true,
      });

      // Focus a body block so the bands know which section they describe.
      const block = container.querySelector('[data-anchor]') as HTMLElement;
      block.focus();
      block.dispatchEvent(new Event('focus', { bubbles: true }));

      const fmt = container.querySelector('[data-hf-pagefmt]') as HTMLSelectElement;
      const start = container.querySelector('[data-hf-pagestart]') as HTMLInputElement;

      fmt.value = 'lowerRoman';
      fmt.dispatchEvent(new Event('change'));
      start.value = '7';
      start.dispatchEvent(new Event('blur'));

      const readBack = editor.pageNumbering();
      // Both bands describe the SAME section, so both must show the same values.
      const fmtValues = Array.from(
        container.querySelectorAll('[data-hf-pagefmt]'),
      ).map((s) => (s as HTMLSelectElement).value);

      const saved = editor.save();
      editor.close();
      container.remove();
      return {
        readBackStart: readBack.start ?? null,
        readBackFormat: readBack.format ?? null,
        fmtValues,
        saved: Array.from(saved as Uint8Array),
      };
    });

    expect(res.readBackStart).toBe(7);
    expect(res.readBackFormat).toBe('lowerRoman');
    expect(res.fmtValues).toEqual(['lowerRoman', 'lowerRoman']);

    const documentXml = readZipEntry(Buffer.from(res.saved), 'word/document.xml');
    expect(documentXml).toContain('w:start="7"');
    expect(documentXml).toContain('w:fmt="lowerRoman"');
  });
});

test.describe('Paginated preview substitutes real page numbers', () => {
  /**
   * Drive the paginator over a synthetic staging DOM — the same shape the converter emits: a
   * section wrapper carrying the section's `w:pgNumType`, a header/footer registry entry whose
   * content holds `[data-field]` page-number markers, and enough body blocks to fill several pages.
   */
  async function paginate(
    page: Page,
    opts: { sectionAttrs: string; footerHtml: string; blocks: number },
  ) {
    await page.setContent(`
      <style>
        #staging { font: 16px/16px Arial; }
        .body { height: 20pt; margin: 0; }
      </style>
      <div id="staging">
        <div id="pagination-hf-registry" style="display:none">
          <div data-section="0" data-hf-type="footer-default">${opts.footerHtml}</div>
        </div>
        <div data-section-index="0" ${opts.sectionAttrs}
             data-page-width="122" data-page-height="122"
             data-content-width="100" data-content-height="60"
             data-margin-top="1" data-margin-right="1"
             data-margin-bottom="1" data-margin-left="1">
          ${Array.from({ length: opts.blocks }, (_, i) => `<p class="body">line ${i}</p>`).join('')}
        </div>
      </div>
      <div id="container"></div>`);
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

    return page.evaluate(() => {
      const staging = document.getElementById('staging') as HTMLElement;
      const container = document.getElementById('container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const result = new PaginationEngine(staging, container, {
        showPageNumbers: false,
      }).paginate();
      const rendered = Array.from(container.querySelectorAll('.page-box')).map((box) =>
        Array.from(box.querySelectorAll('[data-field]')).map((m) => m.textContent),
      );
      return {
        totalPages: result.totalPages,
        rendered,
        storyAnchorsInPages: container.querySelectorAll(
          '.page-header [data-anchor], .page-footer [data-anchor]',
        ).length,
        editableStoryClones: container.querySelectorAll(
          '.page-header [contenteditable="true"], .page-footer [contenteditable="true"]',
        ).length,
      };
    });
  }

  test('a cloned PAGE field counts up instead of repeating its cached result', async ({ page }) => {
    const res = await paginate(page, {
      sectionAttrs: '',
      footerHtml: '<span data-field="PAGE">1</span>',
      blocks: 9,
    });

    expect(res.totalPages).toBeGreaterThan(1);
    // The bug this closes: every page showed the one cached "1".
    expect(res.rendered.map((m) => m[0])).toEqual(
      Array.from({ length: res.totalPages }, (_, i) => String(i + 1)),
    );
  });

  test("the section's format and restart value drive the number, and NUMPAGES gets the total", async ({
    page,
  }) => {
    const res = await paginate(page, {
      sectionAttrs: 'data-page-num-start="4" data-page-num-fmt="lowerRoman"',
      footerHtml: '<span data-field="PAGE">1</span> of <span data-field="NUMPAGES">1</span>',
      blocks: 9,
    });

    expect(res.totalPages).toBeGreaterThan(1);
    // Page numbers count from the section's own start (4 → iv, v, vi …) in the section's format;
    // NUMPAGES is the document total, which is only knowable once the last page exists.
    const total = formatRoman(res.totalPages);
    expect(res.rendered).toEqual(
      Array.from({ length: res.totalPages }, (_, i) => [formatRoman(4 + i), total]),
    );
  });

  test('cloned header/footer content is inert, so a per-page number can never be committed', async ({
    page,
  }) => {
    const res = await paginate(page, {
      sectionAttrs: '',
      footerHtml:
        '<p data-anchor="ftr-unid-1" contenteditable="true" data-committed-text="1">' +
        '<span data-field="PAGE">1</span></p>',
      blocks: 9,
    });

    // A running story is authored once and cloned onto every page, so every clone carries the SAME
    // data-anchor. Substituting a different number into each one turns that latent duplicate into a
    // corruption vector: committing any one clone writes that page's number into the shared story
    // as literal text and destroys the PAGE field. Page-box story content is presentation — the
    // docked editing bands are the addressable affordance.
    expect(res.totalPages).toBeGreaterThan(1);
    expect(res.storyAnchorsInPages).toBe(0);
    expect(res.editableStoryClones).toBe(0);
    // …and the substitution itself still works.
    expect(res.rendered.map((m) => m[0])).toEqual(
      Array.from({ length: res.totalPages }, (_, i) => String(i + 1)),
    );
  });

  test("a field's own \\* switch overrides the section's format", async ({ page }) => {
    const res = await paginate(page, {
      sectionAttrs: 'data-page-num-fmt="lowerRoman"',
      footerHtml: '<span data-field="PAGE" data-field-format="ALPHABETIC">1</span>',
      blocks: 9,
    });

    // ALPHABETIC, not the section's lowerRoman — the field's own switch wins.
    expect(res.rendered.map((m) => m[0])).toEqual(
      Array.from({ length: res.totalPages }, (_, i) => 'ABCDEFGHIJ'[i]),
    );
  });
});
