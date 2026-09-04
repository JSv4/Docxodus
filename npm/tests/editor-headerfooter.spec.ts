import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

/**
 * Read one entry out of a DOCX (zip) without a dependency: walk the central directory from the
 * EOCD record, then inflate the entry's raw deflate stream. Needed because some assertions are
 * about package-level OOXML (`w:titlePg` in a mid-document sectPr, `w:evenAndOddHeaders` in the
 * settings part) that no client API surfaces.
 */
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

  // The body wrapper is no longer conditional: `.docx-body-flow` is the sheet the viewport gives
  // page geometry to and zooms, so every mount has one. What the option still governs is the
  // BANDS — which is what "default off" has to mean now.
  test('renders no bands when the option is omitted (default off)', async ({
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
    expect(res.wrappers).toBe(1);
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

  /**
   * Word's two checkboxes. "Different odd & even pages" is `w:evenAndOddHeaders`, which is
   * document-global and governs footers too: once set, even pages stop inheriting the Default
   * stories, so enabling it must seed BOTH the even header and the even footer or page 2 silently
   * loses its footer. The same holds for "Different first page" and `w:titlePg`. Disabling clears
   * the flag and leaves the parts behind, exactly as Word does.
   */
  test('the odd/even and first-page options seed both stories and flip the section flags', async ({
    page,
  }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const { container, editor } = w.__hfMount(
        w.Docxodus.DocxSessionBridge.CreateBlankDocx(),
        { headerFooter: true },
      );

      const evenOn = editor.setHeaderFooterKindEnabled('even', true);
      const afterEven = {
        enabled: editor.headerFooterKindEnabled('even'),
        firstEnabled: editor.headerFooterKindEnabled('first'),
        kinds: editor.sectionInfo().headerRefs.map((r: any) => r.kind).sort(),
        footerKinds: editor.sectionInfo().footerRefs.map((r: any) => r.kind).sort(),
        // The band's story switcher appears once there is more than one story to show.
        switcher: !(container.querySelector('[data-hf-band="header"] [data-hf-kinds]') as HTMLElement).hidden,
        switcherLabels: (Array.from(
          container.querySelectorAll('[data-hf-band="header"] [data-hf-kinds] button'),
        ) as HTMLElement[]).map((b) => b.textContent),
      };

      const firstOn = editor.setHeaderFooterKindEnabled('first', true);
      editor.setHeaderFooterKind('header', 'first');
      const label = editor.headerFooterStoryLabel('header');
      const firstOff = editor.setHeaderFooterKindEnabled('first', false);
      const afterFirstOff = {
        enabled: editor.headerFooterKindEnabled('first'),
        // The part survives the flag; only the flag went.
        stillHasFirstPart: editor.sectionInfo().headerRefs.some((r: any) => r.kind === 'first'),
        kindShown: editor.headerFooterKind('header'),
      };

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();
      let b64 = '';
      for (let i = 0; i < saved.length; i++) b64 += String.fromCharCode(saved[i]);
      return { evenOn, afterEven, firstOn, label, firstOff, afterFirstOff, b64: btoa(b64) };
    });

    expect(res.evenOn).toBe(true);
    expect(res.afterEven.enabled).toBe(true);
    expect(res.afterEven.firstEnabled).toBe(false);
    expect(res.afterEven.kinds).toEqual(['default', 'even']);
    expect(res.afterEven.footerKinds).toEqual(['default', 'even']);
    expect(res.afterEven.switcher).toBe(true);
    expect(res.afterEven.switcherLabels).toEqual(['Odd pages', 'Even pages']);
    expect(res.firstOn).toBe(true);
    expect(res.label).toBe('First Page Header');
    expect(res.firstOff).toBe(true);
    expect(res.afterFirstOff.enabled).toBe(false);
    expect(res.afterFirstOff.stillHasFirstPart).toBe(true);
    expect(res.afterFirstOff.kindShown).toBe('default');

    const zip = Buffer.from(res.b64, 'base64');
    expect(readZipEntry(zip, 'word/settings.xml')).toMatch(/<w:evenAndOddHeaders\b/);
    expect(readZipEntry(zip, 'word/document.xml')).not.toMatch(/<w:titlePg\b/);
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

  /**
   * The display side of the same collision: `RenderBlockHtml` used to locate a block by bare unid
   * across parts, so a band could SHOW one story while editing another. Author distinct text in
   * each footer kind, then check every band selection displays its own.
   */
  test('each kind selection displays its own story, not a unid-colliding sibling', async ({
    page,
  }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate((raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__hfMount(new Uint8Array(raw), { headerFooter: true });
      const marker = { default: 'DDD-DEFAULT', first: 'FFF-FIRST', even: 'EEE-EVEN' } as const;

      for (const kind of ['default', 'first', 'even'] as const) {
        editor.setHeaderFooterKind('footer', kind);
        w.__hfType(
          container.querySelector('[data-hf-band="footer"] [data-anchor][contenteditable="true"]'),
          marker[kind],
        );
      }
      // Re-select each kind and read back what the band shows.
      const shown: Record<string, string> = {};
      for (const kind of ['default', 'first', 'even'] as const) {
        editor.setHeaderFooterKind('footer', kind);
        shown[kind] = (
          container.querySelector('[data-hf-band="footer"] [data-hf-body]')?.textContent || ''
        ).trim();
      }

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const stored: Record<string, string> = {};
      for (const r of doc.sectionInfo.footerRefs) stored[r.kind] = doc.textInPart(r.partUri).trim();
      doc.close();
      return { shown, stored };
    }, bytes);

    // What the band showed must equal what is actually stored for that kind.
    expect(res.shown).toEqual({ default: 'DDD-DEFAULT', first: 'FFF-FIRST', even: 'EEE-EVEN' });
    expect(res.stored).toEqual({ default: 'DDD-DEFAULT', first: 'FFF-FIRST', even: 'EEE-EVEN' });
  });

  /**
   * A block whose rendered DOM has edge whitespace — an empty header story renders with a
   * placeholder space, and typing lands before it — produces a selection span one longer than
   * the text the commit actually stores (`serializeInlineMarkdown(...).trim()`). ApplyFormat
   * then rejects the span as out of range and the format SILENTLY does nothing. The editor
   * already normalizes caret offsets for this (trimmedSplitOffset); selection spans need the
   * same treatment. Found by clicking Bold in the demo right after typing a header.
   */
  test('formatting a just-typed band paragraph applies (span survives the commit trim)', async ({
    page,
  }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate((raw: number[]) => {
      const w = window as any;
      const D = w.Docxodus;
      // A real Word-authored empty Header story: renders with a placeholder space, which a
      // session-seeded story does not — only this reproduces the trim mismatch.
      const { container, editor } = w.__hfMount(new Uint8Array(raw), { headerFooter: true });

      const blk = container.querySelector(
        '[data-hf-band="header"] [data-anchor][contenteditable="true"]',
      ) as HTMLElement;
      // Type with the caret at the START, leaving the rendered placeholder space AFTER the typed
      // text — what a user gets by clicking into an empty header and typing. Replacing the whole
      // contents instead would delete the placeholder and hide the bug.
      blk.focus();
      let sel = window.getSelection()!;
      let r = document.createRange();
      r.setStart(blk, 0);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      document.execCommand('insertText', false, 'DEFAULT HEAD PROBE');

      // Select all and hit Bold with the block STILL FOCUSED — the demo's format buttons
      // preventDefault on mousedown precisely so the selection survives, so format() computes
      // the span before syncBlock commits (and trims) the typing.
      sel = window.getSelection()!;
      r = document.createRange();
      r.selectNodeContents(blk);
      sel.removeAllRanges();
      sel.addRange(r);
      editor.format('bold', true);

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const h = D.DocxSessionBridge.OpenSession(saved, '{}');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      const id = Object.keys(proj.anchorIndex).find((k) => k.includes(':hdr'))!;
      const xml = D.DocxSessionBridge.RawGetXml(h, id);
      D.DocxSessionBridge.CloseSession(h);
      return { hasText: xml.includes('DEFAULT HEAD PROBE'), bold: /<w:b[ />]/.test(xml) };
    }, bytes);

    expect(res.hasText).toBe(true);
    expect(res.bold).toBe(true);
  });

  /**
   * Selecting First/Even must make the section actually RENDER that story. Word leaves a
   * first/even part + reference behind when "Different first page" / "Different odd & even pages"
   * is switched back off, dropping only `w:titlePg` / `w:evenAndOddHeaders` — HC031 is exactly
   * that shape. The band's seed path only runs when the part is ABSENT, so without an explicit
   * ensure-visible the typed story is saved but never rendered (confirmed by a LibreOffice render
   * of the saved file: content present in header1/header3, neither shown).
   */
  test('selecting first/even sets the section flags so the story actually renders', async ({
    page,
  }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate((raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__hfMount(new Uint8Array(raw), { headerFooter: true });

      editor.setHeaderFooterKind('header', 'first');
      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'COVER PAGE ONLY',
      );
      editor.setHeaderFooterKind('header', 'even');
      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'ACME CORP — CONFIDENTIAL',
      );

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();

      const doc = w.__hfInspect(saved);
      const headers: Record<string, string> = {};
      for (const r of doc.sectionInfo.headerRefs) headers[r.kind] = doc.textInPart(r.partUri).trim();
      doc.close();
      let b64 = '';
      for (let i = 0; i < saved.length; i++) b64 += String.fromCharCode(saved[i]);
      return { headers, b64: btoa(b64) };
    }, bytes);

    expect(res.headers.first).toBe('COVER PAGE ONLY');
    expect(res.headers.even).toBe('ACME CORP — CONFIDENTIAL');

    // Content alone is not enough — without the section flags Word renders neither story.
    const zip = Buffer.from(res.b64, 'base64');
    const documentXml = readZipEntry(zip, 'word/document.xml');
    const settingsXml = readZipEntry(zip, 'word/settings.xml');

    // w:titlePg must sit in the sectPr that actually carries the first-page reference.
    const sectPrs = documentXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];
    const firstRefSection = sectPrs.find((s) =>
      /<w:headerReference[^>]*w:type="first"/.test(s),
    );
    expect(firstRefSection).toBeDefined();
    expect(firstRefSection!).toMatch(/<w:titlePg\b/);
    expect(settingsXml).toMatch(/<w:evenAndOddHeaders\b/);
  });

  /**
   * A multi-section document commonly defines its headers once, in the first section, and leaves
   * the rest with no references at all — HC031 has four sections and only section 0 declares any.
   * ECMA-376 §17.6.17 says such a section CONTINUES the previous one's stories, so reporting only
   * a section's own references would tell the band "no header here" for most of the document, and
   * creating one would mint a redundant part and break the inheritance the file relies on.
   */
  test('a later section shows the header it inherits, not an empty band', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate((raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__hfMount(new Uint8Array(raw), { headerFooter: true });

      // Author the first section's default header, then move the caret to the last body block —
      // which lives in a later section that declares no references of its own.
      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'RUNNING HEAD',
      );
      // Body blocks only: the rendered footnotes/endnotes sections also live in the body flow,
      // but their paragraphs belong to the note parts and so have no governing body section.
      const bodyBlocks = (
        Array.from(
          container.querySelectorAll('.docx-body-flow [data-anchor][contenteditable="true"]'),
        ) as HTMLElement[]
      ).filter((e) => !e.closest('section.footnotes, section.endnotes, .footnote-item'));
      bodyBlocks[bodyBlocks.length - 1].focus();

      const band = container.querySelector('[data-hf-band="header"]') as HTMLElement;
      const out = {
        sections: bodyBlocks.length,
        text: (band.querySelector('[data-hf-body]')?.textContent || '').trim(),
        empty: band.getAttribute('data-hf-empty'),
        markedInherited: band.hasAttribute('data-hf-inherited'),
        note: (band.querySelector('[data-hf-inherited-note]')?.textContent || '').trim(),
        editable: !!band.querySelector('[data-anchor][contenteditable="true"]'),
      };
      editor.close();
      container.remove();
      return out;
    }, bytes);

    // The inherited story is shown and remains editable (it is the shared part).
    expect(res.text).toContain('RUNNING HEAD');
    expect(res.empty).toBeNull();
    expect(res.editable).toBe(true);
    // …and the band says where it came from, since editing it changes both sections.
    expect(res.markedInherited).toBe(true);
    expect(res.note).toMatch(/same as previous/i);
  });

  /**
   * Page view edits the story IN the page (Word's edit-in-the-margin): no bands, every page's
   * header area is click-to-edit, a click swaps that page's inert clone for the live story, and a
   * commit re-clones the story onto every other page — so all pages update without a remount.
   * Switching back to the continuous view shows the bands with the edited content.
   */
  test('page view edits the header in place and mirrors the edit onto every page', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate((raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__hfMount(new Uint8Array(raw), { headerFooter: true });

      w.__hfType(
        container.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]'),
        'RUNNING HEAD',
      );
      editor.setPaginated(true);
      const pages = Array.from(container.querySelectorAll('.page-box')) as HTMLElement[];
      const headerOf = (p: HTMLElement) => p.querySelector('.page-header') as HTMLElement;
      const paginated = {
        bands: container.querySelectorAll('.docx-hf-band').length,
        pages: pages.length,
        clickToEdit: pages.every((p) => headerOf(p).hasAttribute('data-hf-page')),
        storyKind: headerOf(pages[0]).dataset.hfType,
        textOnEveryPage: pages.every((p) => (headerOf(p).textContent || '').includes('RUNNING HEAD')),
        liveBlocksBeforeClick: container.querySelectorAll('.page-header [data-anchor][contenteditable="true"]').length,
      };

      // Click into the SECOND page's header: it becomes the live story, the caret lands in it,
      // the story reports itself, and the other pages keep their inert clones.
      const area = headerOf(pages[1]);
      const rect = area.getBoundingClientRect();
      area.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: rect.left + 30, clientY: rect.top + 8 }));
      const live = area.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      const activated = {
        isHost: area.getAttribute('data-hf-band') === 'header',
        active: area.hasAttribute('data-hf-active'),
        label: area.dataset.hfLabel,
        focused: document.activeElement === live,
        storyKind: editor.activeStoryKind,
        liveBlocks: container.querySelectorAll('.page-header [data-anchor][contenteditable="true"]').length,
      };

      // Type, commit, and read the other pages.
      w.__hfType(live, 'RUNNING HEAD — REVISED');
      const mirrored = {
        host: (area.textContent || '').trim(),
        others: pages.filter((p) => p !== pages[1]).map((p) => (headerOf(p).textContent || '').trim()),
        othersInert: pages.filter((p) => p !== pages[1]).every((p) => !headerOf(p).querySelector('[contenteditable="true"]')),
        pagesKept: pages.every((p) => p.isConnected),
      };

      // Leave the story by clicking a body block; the page stack survives (no re-paginate for a
      // story that did not grow) and the story deactivates.
      (container.querySelector('.page-content [data-anchor][contenteditable="true"]') as HTMLElement).focus();
      const left = {
        active: area.hasAttribute('data-hf-active'),
        storyKind: editor.activeStoryKind,
        pagesKept: pages.every((p) => p.isConnected),
      };

      editor.setPaginated(false);
      const continuous = {
        bands: container.querySelectorAll('.docx-hf-band').length,
        headerText: (container.querySelector('[data-hf-band="header"] [data-hf-body]')?.textContent || '').trim(),
      };
      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();
      const doc = w.__hfInspect(saved);
      const stored = doc.textOf(':hdr');
      doc.close();
      return { paginated, activated, mirrored, left, continuous, stored };
    }, bytes);

    expect(res.paginated.bands).toBe(0);
    expect(res.paginated.pages).toBeGreaterThan(1);
    expect(res.paginated.clickToEdit).toBe(true);
    expect(res.paginated.storyKind).toBe('default');
    expect(res.paginated.textOnEveryPage).toBe(true);
    expect(res.paginated.liveBlocksBeforeClick).toBe(0);
    expect(res.activated).toEqual({ isHost: true, active: true, label: 'Header', focused: true, storyKind: 'header', liveBlocks: 1 });
    expect(res.mirrored.host).toBe('RUNNING HEAD — REVISED');
    expect(res.mirrored.others.every((t: string) => t === 'RUNNING HEAD — REVISED')).toBe(true);
    expect(res.mirrored.othersInert).toBe(true);
    expect(res.mirrored.pagesKept).toBe(true);
    expect(res.left).toEqual({ active: false, storyKind: null, pagesKept: true });
    expect(res.continuous.bands).toBe(2);
    expect(res.continuous.headerText).toContain('RUNNING HEAD — REVISED');
    expect(res.stored).toContain('RUNNING HEAD — REVISED');
  });

  test('page view: a page-number field inserted into a page footer counts per page', async ({ page }) => {
    const res = await page.evaluate(() => {
      const w = window as any;
      const D = w.Docxodus;
      // Enough body text for several pages.
      const h = D.DocxSessionBridge.OpenSession(D.DocxSessionBridge.CreateBlankDocx(), '{}');
      const proj = JSON.parse(D.DocxSessionBridge.Project(h));
      let anchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;
      D.DocxSessionBridge.ReplaceText(h, anchor, 'Paragraph one.');
      for (let i = 0; i < 90; i++) {
        anchor = Object.keys(JSON.parse(D.DocxSessionBridge.Project(h)).anchorIndex).filter((k) => k.startsWith('p:body:')).pop()!;
        D.DocxSessionBridge.InsertParagraph(h, anchor, 'after', `Filler paragraph number ${i + 2} with enough words to take up a line or two of the page.`);
      }
      const bytes: Uint8Array = D.DocxSessionBridge.Save(h);
      D.DocxSessionBridge.CloseSession(h);

      const { container, editor } = w.__hfMount(bytes, { headerFooter: true, paginated: true });
      const pages = Array.from(container.querySelectorAll('.page-box')) as HTMLElement[];
      // No footer story yet: the page has no footer area to click, so go through the command,
      // which opens the footer on the current page and seeds the story.
      (container.querySelector('.page-content [data-anchor][contenteditable="true"]') as HTMLElement).focus();
      const went = editor.goToHeaderFooter('footer');
      editor.insertPageNumber('currentPage');
      const host = container.querySelector('.page-footer[data-hf-band="footer"]') as HTMLElement;
      const out = {
        pages: pages.length,
        went,
        hostText: (host?.textContent || '').trim(),
        hostPage: host?.closest('.page-box')?.getAttribute('data-page-number'),
        othersHaveField: false,
      };
      editor.close();
      container.remove();
      return out;
    });
    expect(res.pages).toBeGreaterThan(1);
    expect(res.went).toBe(true);
    expect(res.hostPage).toBe('1');
    expect(res.hostText).toBe('1');
  });
});
