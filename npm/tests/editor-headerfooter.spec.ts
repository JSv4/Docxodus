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
      const fix = warn.querySelector('[data-hf-fix-counterpart]') as HTMLButtonElement;
      const hadFix = !!fix;
      fix?.click();
      const warnAfter = container.querySelector(
        '[data-hf-band="header"] [data-hf-warning]',
      ) as HTMLElement;
      const after = {
        // The note STAYS: turning even pages on really does stop them using the Default
        // stories, whether or not an even footer now exists. Only the offer goes away.
        hidden: warnAfter.hidden,
        stillHasFix: !!warnAfter.querySelector('[data-hf-fix-counterpart]'),
        footerKind: editor.headerFooterKind('footer'),
      };

      // Selecting Default carries no caveat, so the note clears entirely.
      editor.setHeaderFooterKind('header', 'default');
      const onDefault = (
        container.querySelector('[data-hf-band="header"] [data-hf-warning]') as HTMLElement
      ).hidden;

      editor.close();
      container.remove();
      return { before, hadFix, after, onDefault };
    });

    expect(res.before.hidden).toBe(false);
    expect(res.before.text).toMatch(/even pages/i);
    expect(res.hadFix).toBe(true);
    expect(res.after.footerKind).toBe('even');
    expect(res.after.hidden).toBe(false);
    expect(res.after.stillHasFix).toBe(false);
    expect(res.onDefault).toBe(true);
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
      const bodyBlocks = Array.from(
        container.querySelectorAll('.docx-body-flow [data-anchor][contenteditable="true"]'),
      ) as HTMLElement[];
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
    expect(res.note).toMatch(/inherited/i);
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
