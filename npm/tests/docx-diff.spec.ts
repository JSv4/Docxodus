import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

// Minimal browser/WASM smoke for the DocxDiff (IR diff engine) bridge — the
// NEW comparison surface exposed alongside the default WmlComparer-backed
// compareDocuments/getRevisions. Mirrors the WC comparison specs in
// docxodus.spec.ts: load the harness, run each of the three entry points
// against two real fixtures, assert the shape the npm wrappers depend on.

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout: 30000,
  });
}

test.describe('DocxDiff (IR diff engine) bridge', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Compare returns redlined DOCX bytes', async ({ page }) => {
    const left = readTestFile('WC/WC001-Digits.docx');
    const right = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([l, r]) => {
        const res = (window as any).DocxodusTests.docxDiffCompare(
          new Uint8Array(l),
          new Uint8Array(r)
        );
        return res.docxBytes ? { length: res.docxBytes.length } : res;
      },
      [Array.from(left), Array.from(right)]
    );

    expect(result.error).toBeUndefined();
    // A redlined DOCX is a non-trivial zip package.
    expect(result.length).toBeGreaterThan(1000);
  });

  test('GetRevisions returns anchor-addressed revisions', async ({ page }) => {
    const left = readTestFile('WC/WC001-Digits.docx');
    const right = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([l, r]) => {
        return (window as any).DocxodusTests.docxDiffGetRevisions(
          new Uint8Array(l),
          new Uint8Array(r)
        );
      },
      [Array.from(left), Array.from(right)]
    );

    expect(result.error).toBeUndefined();
    expect(Array.isArray(result.revisions)).toBe(true);
    expect(result.revisions.length).toBeGreaterThan(0);

    // The IR engine's differentiator: every revision carries the wire shape the
    // npm wrapper maps, and at least one anchor side is present per revision.
    for (const rev of result.revisions) {
      expect(typeof rev.revisionType).toBe('string');
      expect(rev.leftAnchor != null || rev.rightAnchor != null).toBe(true);
    }
  });

  test('GetEditScriptJson returns parseable diff-as-data', async ({ page }) => {
    const left = readTestFile('WC/WC001-Digits.docx');
    const right = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([l, r]) => {
        return (window as any).DocxodusTests.docxDiffGetEditScript(
          new Uint8Array(l),
          new Uint8Array(r)
        );
      },
      [Array.from(left), Array.from(right)]
    );

    expect(result.error).toBeUndefined();
    expect(typeof result.editScript).toBe('string');
    // The script is machine-readable JSON.
    const parsed = JSON.parse(result.editScript);
    expect(parsed).toBeTruthy();
  });

  test('settings JSON is honored (detectMoves=false still diffs)', async ({ page }) => {
    const left = readTestFile('WC/WC001-Digits.docx');
    const right = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([l, r]) => {
        return (window as any).DocxodusTests.docxDiffGetRevisions(
          new Uint8Array(l),
          new Uint8Array(r),
          JSON.stringify({ detectMoves: false, caseInsensitive: true })
        );
      },
      [Array.from(left), Array.from(right)]
    );

    expect(result.error).toBeUndefined();
    expect(Array.isArray(result.revisions)).toBe(true);
  });

  test('trackBlockFormatChanges=false suppresses block-format markup', async ({ page }) => {
    // A paragraph-property-only change (jc) produces NO w:pPrChange when the opt-out is off.
    const left = readTestFile('WC/WC001-Digits.docx');
    const right = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([l, r]) => {
        return (window as any).DocxodusTests.docxDiffGetRevisions(
          new Uint8Array(l),
          new Uint8Array(r),
          JSON.stringify({ trackBlockFormatChanges: false })
        );
      },
      [Array.from(left), Array.from(right)]
    );

    expect(result.error).toBeUndefined();
    // No block-format (non-Run) FormatChanged revisions when the flag is off.
    const blockFmt = (result.revisions as any[]).filter(
      (rev) => rev.formatChange && rev.formatChange.scope && rev.formatChange.scope !== 'run'
    );
    expect(blockFmt.length).toBe(0);
  });

  // The round-trip contract — NOT a shape/length check. compare(left,right) then
  // accept ≡ right and reject ≡ left at the per-block text level. This rides the
  // full client wire: Compare (bytes out), AcceptRevisions/RejectRevisions (bytes
  // in→out, the new surface), all marshalled JS↔WASM. A wire/type-mapping break in
  // any of those diff paths corrupts the bytes and breaks the text equality below.
  test('accept/reject round-trip: accept ≡ right, reject ≡ left (text level)', async ({ page }) => {
    const left = readTestFile('WC/WC001-Digits.docx');
    const right = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([l, r]) => {
        const T = (window as any).DocxodusTests;
        const text = (bytes: Uint8Array): string | { __err: unknown } => {
          const h = T.convertToHtml(bytes);
          if (h.error) return { __err: h.error };
          return (h.html as string)
            .replace(/<[^>]+>/g, ' ')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/\s+/g, ' ')
            .trim();
        };
        const cmp = T.docxDiffCompare(new Uint8Array(l), new Uint8Array(r));
        if (cmp.error) return { stage: 'compare', err: cmp.error };
        const acc = T.docxDiffAcceptRevisions(cmp.docxBytes);
        if (acc.error) return { stage: 'accept', err: acc.error };
        const rej = T.docxDiffRejectRevisions(cmp.docxBytes);
        if (rej.error) return { stage: 'reject', err: rej.error };
        return {
          leftText: text(new Uint8Array(l)),
          rightText: text(new Uint8Array(r)),
          acceptText: text(acc.docxBytes),
          rejectText: text(rej.docxBytes),
        };
      },
      [Array.from(left), Array.from(right)]
    );

    expect((result as any).err, `failed at stage: ${(result as any).stage}`).toBeUndefined();
    const r = result as {
      leftText: string;
      rightText: string;
      acceptText: string;
      rejectText: string;
    };
    expect(r.acceptText).toBe(r.rightText); // accept materializes the revised side
    expect(r.rejectText).toBe(r.leftText); // reject restores the original side
    // and the two sides genuinely differ — so the round-trip can't pass by echoing one input.
    expect(r.rightText).not.toBe(r.leftText);
  });
});

// Composite N-way consolidate: merge two reviewers' edits against a shared base.
// Uses the same WC fixtures — base is the original, both reviewers edit the same
// modified span (under distinct authors) so the second reviewer overlaps the
// first, exercising the conflict path.
test.describe('DocxDiff consolidate (composite N-way) bridge', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Consolidate two reviewers returns redlined DOCX bytes', async ({ page }) => {
    const base = readTestFile('WC/WC001-Digits.docx');
    const rev = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([b, r]) => {
        const res = (window as any).DocxodusTests.docxDiffConsolidate(
          new Uint8Array(b),
          [
            { author: 'Alice', document: new Uint8Array(r) },
            { author: 'Bob', document: new Uint8Array(r) },
          ]
        );
        return res.docxBytes ? { length: res.docxBytes.length } : res;
      },
      [Array.from(base), Array.from(rev)]
    );

    expect(result.error).toBeUndefined();
    expect(result.length).toBeGreaterThan(1000);
  });

  test('GetConflicts reports overlapping reviewer edits', async ({ page }) => {
    // Two reviewers edit the SAME word differently ("quick" -> "SLOW" vs "FAST"),
    // which is a genuine cross-reviewer token-overlap conflict. (Two reviewers making
    // the IDENTICAL edit would be a consensus, not a conflict, and yield zero.)
    const base = readTestFile('RC/RC100-Conflict-Base.docx');
    const alice = readTestFile('RC/RC100-Conflict-Alice.docx');
    const bob = readTestFile('RC/RC100-Conflict-Bob.docx');

    const result = await page.evaluate(
      ([b, a, f]) => {
        return (window as any).DocxodusTests.docxDiffGetConflicts(
          new Uint8Array(b),
          [
            { author: 'Alice', document: new Uint8Array(a) },
            { author: 'Bob', document: new Uint8Array(f) },
          ]
        );
      },
      [Array.from(base), Array.from(alice), Array.from(bob)]
    );

    expect(result.error).toBeUndefined();
    expect(Array.isArray(result.conflicts)).toBe(true);
    // The overlapping edit on the same word produces at least one conflict, each
    // with its competing per-author variants.
    expect(result.conflicts.length).toBeGreaterThan(0);
    for (const c of result.conflicts) {
      expect(typeof c.baseAnchor).toBe('string');
      expect(Array.isArray(c.competitors)).toBe(true);
      expect(c.competitors.length).toBeGreaterThan(0);
    }
  });

  test('GetConsolidatedRevisions returns multi-author revisions', async ({ page }) => {
    const base = readTestFile('WC/WC001-Digits.docx');
    const rev = readTestFile('WC/WC001-Digits-Mod.docx');

    const result = await page.evaluate(
      ([b, r]) => {
        return (window as any).DocxodusTests.docxDiffGetConsolidatedRevisions(
          new Uint8Array(b),
          [
            { author: 'Alice', document: new Uint8Array(r) },
            { author: 'Bob', document: new Uint8Array(r) },
          ]
        );
      },
      [Array.from(base), Array.from(rev)]
    );

    expect(result.error).toBeUndefined();
    expect(Array.isArray(result.revisions)).toBe(true);
    expect(result.revisions.length).toBeGreaterThan(0);
    for (const rv of result.revisions) {
      expect(typeof rv.revisionType).toBe('string');
      expect(typeof rv.author).toBe('string');
    }
  });
});
