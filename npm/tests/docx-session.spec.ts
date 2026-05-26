import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('DocxSession (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('open, project, replaceText, save, reopen — round-trip', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        // Pick first body heading/paragraph anchor by document order
        const anchorEntries = Object.entries(proj.anchorIndex) as [string, any][];
        const firstBody = anchorEntries
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map(t => ({ t, idx: proj.markdown.indexOf('{#' + t.id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a, b) => a.idx - b.idx)[0];

        const replaceResult = JSON.parse(bridge.ReplaceText(handle, firstBody.t.id, '**JSMARKER** replaced.'));

        const after = JSON.parse(bridge.Project(handle));
        const saved = bridge.Save(handle);

        return {
          replaceSuccess: replaceResult.success,
          replaceError: replaceResult.error,
          targetAnchor: firstBody.t.id,
          markdownContainsMarker: after.markdown.includes('JSMARKER'),
          markdownExcerpt: after.markdown.substring(0, 400),
          savedBytes: saved.length,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.replaceSuccess).toBe(true);
    expect(result.markdownContainsMarker, `target=${result.targetAnchor}\nexcerpt:\n${result.markdownExcerpt}`).toBe(true);
    expect(result.savedBytes).toBeGreaterThan(0);
  });

  test('error envelope: malformed markdown gives typed error code', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const anchorId = Object.keys(proj.anchorIndex)[0];
        // Pipe table → TableInsertNotSupported
        const r = JSON.parse(bridge.ReplaceText(handle, anchorId, '| a | b |\n|---|---|\n| 1 | 2 |'));
        return { success: r.success, errorCode: r.error?.code };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.success).toBe(false);
    expect(result.errorCode).toBe('table_insert_not_supported');
  });

  test('Undo restores prior state', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const before = JSON.parse(bridge.Project(handle)).markdown;
        const proj = JSON.parse(bridge.Project(handle));
        const anchorId = Object.keys(proj.anchorIndex).find(k => k.startsWith('h:body:'))!;
        bridge.ReplaceText(handle, anchorId, '**TEMPORARY**');
        const undidOk = bridge.Undo(handle);
        const after = JSON.parse(bridge.Project(handle)).markdown;
        return { undidOk, restored: before === after };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.undidOk).toBe(true);
    expect(result.restored).toBe(true);
  });

  test('ReplaceTextRange + replaceMatch round-trip through the WASM bridge', async ({ page }) => {
    // Grep finds the '[' placeholder markers; replaceMatch addresses each by its
    // (anchor, span) so the wrong one never gets picked when several share text.
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const matchesBefore = JSON.parse(bridge.Grep(handle, '\\[', JSON.stringify({ scope: 1 })));
        const totalBefore = matchesBefore.length;
        if (totalBefore < 1) return { totalBefore, error: 'fixture has no matches' };

        // Replace the FIRST match via span-addressed bridge call.
        const target = matchesBefore[0];
        const r = JSON.parse(bridge.ReplaceTextAtSpan(
          handle,
          target.enclosingAnchor.id,
          target.span.start,
          target.span.length,
          '⟪BRACKET⟫'
        ));

        const matchesAfter = JSON.parse(bridge.Grep(handle, '\\[', JSON.stringify({ scope: 1 })));
        const proj = JSON.parse(bridge.Project(handle));
        return {
          totalBefore,
          totalAfter: matchesAfter.length,
          rSuccess: r.success,
          rError: r.error,
          containsMarker: proj.markdown.includes('⟪BRACKET⟫'),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.error).toBeUndefined();
    expect(result.rSuccess).toBe(true);
    // Exactly one '[' replaced → match count drops by one. (We don't assert on
    // the marker string surviving in the projection — the projector escapes
    // markdown punctuation, but the run-text edit did land if the count dropped.)
    expect(result.totalAfter).toBe(result.totalBefore - 1);
  });

  test('applyFormatBySubstring resolves a visible substring to a span and formats it', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // Pick a paragraph with bracketed text we know exists.
        const placeholders = JSON.parse(bridge.FindPlaceholders(handle, 7, 1, 80, 0));
        if (placeholders.length === 0) return { error: 'no placeholders' };
        const target = placeholders[0];

        // Format the placeholder text via the substring overload.
        const r = JSON.parse(bridge.ApplyFormatBySubstring(
          handle, target.match.enclosingAnchor.id, target.match.text, JSON.stringify({ Bold: true })
        ));
        return { rSuccess: r.success, rError: r.error };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.error).toBeUndefined();
    expect(result.rSuccess).toBe(true);
  });

  test('findPlaceholders enumerates and classifies template slots', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // PlaceholderKinds.All = 7, body scope = 1.
        const placeholders = JSON.parse(bridge.FindPlaceholders(handle, 7, 1, 80, 0));
        const kinds = new Set(placeholders.map((p: any) => p.kind));
        return {
          total: placeholders.length,
          kinds: Array.from(kinds),
          firstKind: placeholders[0]?.kind,
          firstMatchHasFragments: Array.isArray(placeholders[0]?.match?.fragments),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    // HC001 has bracketed placeholders in every day-section paragraph.
    expect(result.total).toBeGreaterThan(0);
    expect(['blank_fill', 'alternative_clause', 'instruction']).toContain(result.firstKind);
    expect(result.firstMatchHasFragments).toBe(true);
  });

  test('Grep returns matches with run-fragment breakdown', async ({ page }) => {
    // HC001 is a multilingual tour-plan template full of `[placeholder]` slots;
    // search for an opening bracket which is reliably present regardless of language.
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // Body scope (1) + small context window. The pattern is a literal '['
        // (escaped for regex) — it appears in every template-placeholder slot.
        const matches = JSON.parse(bridge.Grep(handle, '\\[', JSON.stringify({ scope: 1, contextChars: 20 })));
        const first = matches[0];
        return {
          count: matches.length,
          firstText: first?.text,
          firstHasFragments: Array.isArray(first?.fragments) && first.fragments.length >= 1,
          firstFragmentHasUnid: typeof first?.fragments?.[0]?.unid === 'string',
          firstFragmentHasFormatting: typeof first?.fragments?.[0]?.formatting === 'object',
          firstContextBeforeIsString: typeof first?.contextBefore === 'string',
          firstEnclosingAnchorKind: first?.enclosingAnchor?.kind,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.count).toBeGreaterThan(0);
    expect(result.firstText).toBe('[');
    expect(result.firstHasFragments).toBe(true);
    expect(result.firstFragmentHasUnid).toBe(true);
    expect(result.firstFragmentHasFormatting).toBe(true);
    expect(result.firstContextBeforeIsString).toBe(true);
    expect(['p', 'h', 'li']).toContain(result.firstEnclosingAnchorKind);
  });

  test('FindByAnnotation / FindByLabel / ListAnnotations bridge end-to-end (#132)', async ({ page }) => {
    // Add two annotations sharing a labelId by discovering needles dynamically: the
    // CreateBookmarkFromSearch path requires the exact text to be present, so a
    // hard-coded needle would couple the test to fixture content drift. Probe the
    // projection for two distinct paragraph text fragments that both repeat at
    // least once in the document.
    const bytes = readTestFile('HC006-Test-01.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const docConv = (window as any).Docxodus.DocumentConverter;
      const bridge = (window as any).Docxodus.DocxSessionBridge;

      // Step 0: find a needle that appears at least twice in actual document text.
      // We use Grep (which searches block run text, not the projection's anchor
      // wrappers) so the chosen needle is guaranteed to round-trip through
      // AnnotationManager.CreateBookmarkFromSearch.
      const probeHandle = bridge.OpenSession(bin, '');
      let needle = '';
      let diag: any = {};
      try {
        const matches = JSON.parse(bridge.Grep(probeHandle, '[A-Za-z]{3,}', JSON.stringify({ scope: 1, contextChars: 0 })));
        const counts = new Map<string, number>();
        for (const m of matches) counts.set(m.text, (counts.get(m.text) ?? 0) + 1);
        diag = { totalMatches: matches.length, distinctWords: counts.size, sample: matches.slice(0, 5).map((m: any) => m.text) };
        for (const [word, n] of counts) {
          if (n >= 2) { needle = word; break; }
        }
      } finally {
        bridge.CloseSession(probeHandle);
      }
      if (!needle) throw new Error('Could not find a repeated word in fixture for annotation needle: ' + JSON.stringify(diag));

      // Step 1: add two annotations via the public AddAnnotation API. Each call
      // produces a fresh document with one more annotation persisted.
      const annotate = (input: Uint8Array, id: string, labelId: string, label: string, search: string, occurrence: number) => {
        const req = JSON.stringify({ Id: id, LabelId: labelId, Label: label, Color: '#FFEB3B', SearchText: search, Occurrence: occurrence });
        const respStr = docConv.AddAnnotation(input, req);
        const resp = JSON.parse(respStr);
        const b64 = resp.DocumentBytes || resp.documentBytes;
        if (!b64) throw new Error('AddAnnotation returned no DocumentBytes; raw=' + respStr.substring(0, 200));
        const raw = atob(b64);
        const buf = new Uint8Array(raw.length);
        for (let i = 0; i < raw.length; i++) buf[i] = raw.charCodeAt(i);
        return buf;
      };

      let annotated = annotate(bin, 'ann1', 'TOUR_REGION', 'First mention', needle, 1);
      annotated = annotate(annotated, 'ann2', 'TOUR_REGION', 'Second mention', needle, 2);

      // Step 2: open session over the annotated bytes and exercise the new bridges.
      const handle = bridge.OpenSession(annotated, '');
      try {
        const all = JSON.parse(bridge.ListAnnotations(handle));
        const byAnn1 = JSON.parse(bridge.FindByAnnotation(handle, 'ann1'));
        const byLabel = JSON.parse(bridge.FindByLabel(handle, 'TOUR_REGION'));
        const byBookmark = JSON.parse(bridge.FindByBookmark(handle, '_Docxodus_Ann_ann1'));
        const missing = JSON.parse(bridge.FindByAnnotation(handle, 'nope'));
        return {
          allCount: all.length,
          allHaveLabelId: all.every((a: any) => a.labelId === 'TOUR_REGION'),
          ann1AnchorCount: byAnn1.length,
          ann1FirstKind: byAnn1[0]?.kind,
          ann1FirstHasPartUri: typeof byAnn1[0]?.partUri === 'string' && byAnn1[0].partUri.length > 0,
          labelKeys: Object.keys(byLabel).sort(),
          bookmarkAnchorCount: byBookmark.length,
          missingIsEmpty: Array.isArray(missing) && missing.length === 0,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.allCount).toBe(2);
    expect(result.allHaveLabelId).toBe(true);
    expect(result.ann1AnchorCount).toBeGreaterThan(0);
    expect(['p', 'h', 'li', 'tc', 'tbl', 'tr']).toContain(result.ann1FirstKind);
    expect(result.ann1FirstHasPartUri).toBe(true);
    expect(result.labelKeys).toEqual(['ann1', 'ann2']);
    expect(result.bookmarkAnchorCount).toBe(result.ann1AnchorCount);
    expect(result.missingIsEmpty).toBe(true);
  });

  test('GrepCrossBlock returns matches spanning adjacent paragraphs (#146)', async ({ page }) => {
    // The Word file's projection joins paragraphs with `\n`, so a pattern that
    // includes a literal `\n` between two known fragments must match cross-block.
    // HC001 contains adjacent intro paragraphs; we search for any pattern that
    // straddles a paragraph boundary by anchoring on a real run of text from
    // the template — but to keep this stable across template edits, just use a
    // permissive `.+\n.+` regex that must produce at least one cross-block hit
    // in any non-empty doc with two paragraphs.
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // RegexOptions.None = 0. Pattern: any non-whitespace character, the
        // block-boundary separator `\n`, then any non-whitespace — must match
        // the gap between any two adjacent non-empty paragraphs in the doc.
        const opts = JSON.stringify({ regexOptions: 0, scope: 1, contextChars: 10 });
        const matches = JSON.parse(bridge.GrepCrossBlock(handle, '\\S\\n\\S', opts));
        const cross = (matches as any[]).filter(m => m.slices.length > 1);
        const first = cross[0];
        return {
          total: matches.length,
          crossCount: cross.length,
          firstSliceCount: first?.slices.length,
          firstAnchorsCount: first?.enclosingAnchors.length,
          firstTextHasNewline: typeof first?.text === 'string' && first.text.includes('\n'),
          firstSliceHasFragments: Array.isArray(first?.slices?.[0]?.fragments),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.crossCount).toBeGreaterThan(0);
    expect(result.firstSliceCount).toBeGreaterThanOrEqual(2);
    expect(result.firstAnchorsCount).toBe(result.firstSliceCount);
    expect(result.firstTextHasNewline).toBe(true);
    expect(result.firstSliceHasFragments).toBe(true);
  });
});
