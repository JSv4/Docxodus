import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): number[] {
  // Plain array: page.evaluate serializes structured-clone data, and a Uint8Array survives the
  // hop as a plain object, so hand over an array and rebuild it in the page.
  return Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

/**
 * Issue #275 — the DocxEditor's docked header/footer editing bands.
 *
 * Header/footer stories live in their own OOXML parts outside the body, so they cannot be
 * another block in the body flow. The bands render them per story paragraph via RenderBlockHtml
 * (which resolves hdr/ftr anchors natively) and wire them with the SAME wireBlock the body uses,
 * so a story paragraph is an ordinary editable block.
 */

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/**
 * Install `window.__hfSeed(header, footer)` — DOCX bytes carrying those stories, built with the
 * shipped session API — plus `window.__hfMount(bytes, opts)` for the common open-into-a-fresh-
 * container dance. Keeps each test's evaluate body about the behavior under test.
 */
async function installHelpers(page: Page) {
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    (window as any).__hfSeed = (header: string, footer: string): Uint8Array => {
      const h = D.DocxSessionBridge.OpenSession(D.DocxSessionBridge.CreateBlankDocx(), '{}');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      const body = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;
      if (header) D.DocxSessionBridge.SetHeaderText(h, body, 'default', header);
      if (footer) D.DocxSessionBridge.SetFooterText(h, body, 'default', footer);
      const bytes: Uint8Array = D.DocxSessionBridge.Save(h);
      D.DocxSessionBridge.CloseSession(h);
      return bytes;
    };
    (window as any).__hfMount = (bytes: Uint8Array, opts: object) => {
      const container = document.createElement('div');
      document.body.appendChild(container);
      return { container, editor: D.DocxEditor.open(container, bytes, D, opts) };
    };
    /** Select all of `el`'s content and type `text` over it, then commit on blur. */
    (window as any).__hfType = (el: HTMLElement, text: string) => {
      el.focus();
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.selectNodeContents(el);
      sel.removeAllRanges();
      sel.addRange(r);
      document.execCommand('insertText', false, text);
      el.dispatchEvent(new Event('blur'));
    };
    /** Reopen saved bytes and return { proj, xmlOf(anchorSubstring), sectionInfo }. */
    (window as any).__hfInspect = (bytes: Uint8Array) => {
      const h = D.DocxSessionBridge.OpenSession(bytes, '{}');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      const body = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;
      const sectionInfo = JSON.parse(D.DocxSessionBridge.GetSectionInfo(h, body));
      const xmlOf = (needle: string) => {
        const id = Object.keys(proj.anchorIndex).find((k) => k.includes(needle));
        return id ? D.DocxSessionBridge.RawGetXml(h, id) : '';
      };
      const textOf = (needle: string) =>
        Object.entries(proj.anchorIndex)
          .filter(([k]) => (k as string).includes(needle))
          .map(([, t]: any) => t.textPreview || '')
          .join(' ');
      const textInPart = (partUri: string) =>
        Object.entries(proj.anchorIndex)
          .filter(([, t]: any) => t.partUri === partUri)
          .map(([, t]: any) => t.textPreview || '')
          .join(' ');
      const close = () => D.DocxSessionBridge.CloseSession(h);
      return { proj, sectionInfo, xmlOf, textOf, textInPart, close };
    };
  });
}

test.describe('DocxEditor — header/footer region', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
    await installHelpers(page);
  });

  test("bands show the document's existing header and footer text", async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(
        w.__hfSeed('ACME CORP — CONFIDENTIAL', 'Last Updated October 2025'),
        { headerFooter: true },
      );

      const header = container.querySelector('[data-hf-band="header"]') as HTMLElement;
      const footer = container.querySelector('[data-hf-band="footer"]') as HTMLElement;
      const out = {
        headerText: (header?.querySelector('[data-hf-body]')?.textContent || '').trim(),
        footerText: (footer?.querySelector('[data-hf-body]')?.textContent || '').trim(),
        headerEditable: !!header?.querySelector('[data-anchor][contenteditable="true"]'),
        // Exactly one DOM node per story paragraph — no per-page duplicates.
        headerNodes: header?.querySelectorAll('[data-hf-anchor]').length ?? 0,
        bodyWrapped: !!container.querySelector(
          '.docx-body-flow [data-anchor][contenteditable="true"]',
        ),
        // The band must sit OUTSIDE the body edit root, or it would shift every block index.
        bandOutsideBody: !container.querySelector('.docx-body-flow [data-hf-band]'),
      };
      editor.close();
      container.remove();
      return out;
    });

    expect(res.headerText).toContain('ACME CORP — CONFIDENTIAL');
    expect(res.footerText).toContain('Last Updated October 2025');
    expect(res.headerEditable).toBe(true);
    expect(res.headerNodes).toBe(1);
    expect(res.bodyWrapped).toBe(true);
    expect(res.bandOutsideBody).toBe(true);
  });

  test('renders no bands and no body wrapper when the option is omitted (default off)', async ({
    page,
  }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(
        w.Docxodus.DocxSessionBridge.CreateBlankDocx(),
        {},
      );
      const out = {
        bands: container.querySelectorAll('[data-hf-band]').length,
        wrappers: container.querySelectorAll('.docx-body-flow').length,
      };
      editor.close();
      container.remove();
      return out;
    });
    expect(res.bands).toBe(0);
    expect(res.wrappers).toBe(0);
  });

  test('typing in the header band commits and survives save + reopen', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(w.__hfSeed('seed', ''), { headerFooter: true });
      const bodyBefore = container.querySelector('.docx-body-flow [data-anchor]');

      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'QUARTERLY REPORT',
      );

      // A header edit must NOT remount the body — node identity proves it.
      const bodyPreserved = bodyBefore === container.querySelector('.docx-body-flow [data-anchor]');
      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const headerText = doc.textOf(':hdr');
      doc.close();
      return { headerText, bodyPreserved };
    });

    expect(res.headerText).toContain('QUARTERLY REPORT');
    expect(res.bodyPreserved).toBe(true);
  });

  test('the ribbon formats band text: bold applies inside the header story', async ({ page }) => {
    const xml = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(w.__hfSeed('CONFIDENTIAL', ''), {
        headerFooter: true,
      });
      const blk = container.querySelector(
        '[data-hf-band="header"] [data-anchor][contenteditable="true"]',
      ) as HTMLElement;
      blk.focus();
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.selectNodeContents(blk);
      sel.removeAllRanges();
      sel.addRange(r);
      editor.format('bold', true);

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const out = doc.xmlOf(':hdr');
      doc.close();
      return out;
    });

    expect(xml).toContain('CONFIDENTIAL');
    expect(xml).toContain('<w:b');
  });

  test('footer: centered page-number field lands a real PAGE field', async ({ page }) => {
    const xml = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(
        w.Docxodus.DocxSessionBridge.CreateBlankDocx(),
        { headerFooter: true },
      );

      editor.setHeaderFooterKind('footer', 'default'); // seeds the story
      const blk = container.querySelector(
        '[data-hf-band="footer"] [data-anchor][contenteditable="true"]',
      ) as HTMLElement;
      blk.focus();
      editor.setAlignment('center');
      editor.insertPageNumber('currentPage');

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const out = doc.xmlOf(':ftr');
      doc.close();
      return out;
    });

    expect(xml).toContain('PAGE');
    expect(xml).toContain('fldChar');
    expect(xml).toContain('w:jc');
  });

  test('switching to the First kind seeds a separate, kind-labelled story', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(w.__hfSeed('Default header', ''), {
        headerFooter: true,
      });

      editor.setHeaderFooterKind('header', 'first');
      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'COVER PAGE ONLY',
      );

      const saved: Uint8Array = editor.save();
      const kindShown = editor.headerFooterKind('header');
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const refs = doc.sectionInfo.headerRefs || [];
      const firstRef = refs.find((x: any) => x.kind === 'first');
      const defaultRef = refs.find((x: any) => x.kind === 'default');
      const out = {
        kindShown,
        kinds: refs.map((x: any) => x.kind).sort(),
        firstText: firstRef ? doc.textInPart(firstRef.partUri) : '',
        defaultText: defaultRef ? doc.textInPart(defaultRef.partUri) : '',
      };
      doc.close();
      return out;
    });

    expect(res.kindShown).toBe('first');
    // Both stories exist and are distinguishable by kind — the whole point of headerRefs.
    expect(res.kinds).toEqual(['default', 'first']);
    expect(res.firstText).toContain('COVER PAGE ONLY');
    expect(res.defaultText).toContain('Default header');
  });

  test('selecting the Even header kind warns that even pages lose the footer', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(
        w.Docxodus.DocxSessionBridge.CreateBlankDocx(),
        { headerFooter: true },
      );

      editor.setHeaderFooterKind('header', 'even');
      const warn = container.querySelector(
        '[data-hf-band="header"] [data-hf-warning]',
      ) as HTMLElement;
      const before = { hidden: warn.hidden, text: (warn.textContent || '').trim() };

      // The fix button creates the matching even footer, which clears the warning.
      const fix = warn.querySelector('[data-hf-fix-even-footer]') as HTMLButtonElement;
      const hadFix = !!fix;
      fix?.click();
      const after = {
        hidden: (
          container.querySelector('[data-hf-band="header"] [data-hf-warning]') as HTMLElement
        ).hidden,
        footerKind: editor.headerFooterKind('footer'),
      };

      editor.close();
      container.remove();
      return { before, hadFix, after };
    });

    expect(res.before.hidden).toBe(false);
    expect(res.before.text).toMatch(/even pages/i);
    expect(res.hadFix).toBe(true);
    expect(res.after.footerKind).toBe('even');
    expect(res.after.hidden).toBe(true);
  });

  /**
   * Regression pin for a bug the live GUI smoke test found on a real Word document.
   *
   * Unids are CONTENT-ADDRESSED, so several *empty* story paragraphs in different parts share
   * one unid. HC031-Complicated-Document.docx carries all six stories (default/first/even for
   * both header and footer), all empty — its two remaining header paragraphs share one unid and
   * all three footer paragraphs share another. `data-anchor` carries only the bare unid, so a
   * unid-keyed lookup resolved such a block to whichever part was indexed LAST: selecting "Even
   * pages" and typing put the text in header3.xml, the FIRST-page header. The band stamps the
   * full `kind:scope:unid` anchor on each story block, and block→anchor resolution must prefer
   * it over the unid map.
   *
   * The same file also proves why `SectionInfo.HeaderRefs` is needed at all: its part ordinals
   * do NOT follow kind order (header1=even, header2=default, header3=first), so labelling bands
   * by part-collection order would mislabel every one.
   */
  test('editing a kind whose story shares a unid with another part lands in the right part', async ({
    page,
  }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate((raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__hfMount(new Uint8Array(raw), { headerFooter: true });

      editor.setHeaderFooterKind('header', 'even');
      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'EVEN PAGE RUNNING HEAD',
      );
      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const headers: Record<string, string> = {};
      for (const r of doc.sectionInfo.headerRefs) headers[r.kind] = doc.textInPart(r.partUri).trim();
      const footers: Record<string, string> = {};
      for (const r of doc.sectionInfo.footerRefs) footers[r.kind] = doc.textInPart(r.partUri).trim();
      doc.close();
      return { headers, footers };
    }, bytes);

    expect(res.headers.even).toContain('EVEN PAGE RUNNING HEAD');
    expect(res.headers.first).toBe('');
    expect(res.headers.default).toBe('');
    // Nothing leaked into the footers either.
    expect(Object.values(res.footers).join('')).toBe('');
  });

  test('bands survive a paginated toggle and keep their content', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(w.__hfSeed('RUNNING HEAD', 'FOOT'), {
        headerFooter: true,
      });

      const headerText = () =>
        (
          container.querySelector('[data-hf-band="header"] [data-hf-body]')?.textContent || ''
        ).trim();

      editor.setPaginated(true);
      const paginated = {
        bands: container.querySelectorAll('[data-hf-band]').length,
        headerText: headerText(),
        editable: !!container.querySelector(
          '[data-hf-band="header"] [data-anchor][contenteditable="true"]',
        ),
      };

      editor.setPaginated(false);
      const continuous = {
        bands: container.querySelectorAll('[data-hf-band]').length,
        headerText: headerText(),
      };

      editor.close();
      container.remove();
      return { paginated, continuous };
    });

    expect(res.paginated.bands).toBe(2);
    expect(res.paginated.headerText).toContain('RUNNING HEAD');
    expect(res.paginated.editable).toBe(true);
    expect(res.continuous.bands).toBe(2);
    expect(res.continuous.headerText).toContain('RUNNING HEAD');
  });
});
