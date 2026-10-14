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

// Exercises the *typed* DocxSession wrapper (Issue #171): findByText /
// findAllByText / findByRegex / findByKind / exists / replaceMatch. The harness
// exposes `window.Docxodus.openTypedSession(bytes)` which returns a real
// `DocxSession` instance (the IIFE-bundled TS wrapper), so these tests run the
// wrapper code, not just the raw [JSExport] bridge.
//
// Fixture: HC006-Test-01.docx — English body text (so text search has Latin
// words to match), multiple headings, and paragraphs. (HC001's body is CJK,
// with no Latin words for the regex needles to hit.)
test.describe('DocxSession find-by surface (typed wrapper — Issue #171)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('exists() true for a real anchor, false for a bogus id', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytesArray));
      try {
        const proj = session.project();
        const realId = Object.keys(proj.anchorIndex)[0];
        return {
          realId,
          realExists: session.exists(realId),
          bogusExists: session.exists('p:body:deadbeefdeadbeef'),
        };
      } finally {
        session.close();
      }
    }, Array.from(bytes));

    expect(result.realExists).toBe(true);
    expect(result.bogusExists).toBe(false);
  });

  test('findByText / findAllByText locate a real needle and honor ignoreCase', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytesArray));
      try {
        // Pull a real ≥5-letter word out of the body via grep so the needle is
        // not hard-coded to fixture text.
        const matches = session.grep('[A-Za-z]{5,}');
        const needle = matches[0].text as string;

        const single = session.findByText(needle);
        const all = session.findAllByText(needle);

        const upper = needle.toUpperCase();
        const upperStrict = session.findAllByText(upper);            // case-sensitive
        const upperLoose = session.findAllByText(upper, { ignoreCase: true });

        return {
          needle,
          single,                                  // AnchorTargetRef | null
          allCount: all.length,
          missing: session.findByText('zzqx_no_such_needle_42'),
          strictCount: upperStrict.length,
          looseCount: upperLoose.length,
        };
      } finally {
        session.close();
      }
    }, Array.from(bytes));

    expect(result.single).not.toBeNull();
    expect(result.single.id).toBeTruthy();
    expect(result.allCount).toBeGreaterThanOrEqual(1);
    expect(result.missing).toBeNull();
    // ignoreCase never finds fewer than the strict search, and finds the
    // original-cased occurrence(s).
    expect(result.looseCount).toBeGreaterThanOrEqual(result.strictCount);
    expect(result.looseCount).toBeGreaterThanOrEqual(1);
  });

  test('findByRegex returns multiple anchors for a broad pattern', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytesArray));
      try {
        const hits = session.findByRegex('\\S+');   // any non-whitespace text
        const ids = hits.map((h: any) => h.id);
        return { count: hits.length, distinct: new Set(ids).size };
      } finally {
        session.close();
      }
    }, Array.from(bytes));

    expect(result.count).toBeGreaterThanOrEqual(2);
    expect(result.distinct).toBe(result.count);   // anchors are unique
  });

  test('findByKind("p") returns body paragraphs; ("h","body") filters by scope', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytesArray));
      try {
        const paras = session.findByKind('p');
        const headings = session.findByKind('h', 'body');
        return {
          paraCount: paras.length,
          allParaKindP: paras.every((p: any) => p.kind === 'p'),
          headingCount: headings.length,
          allHeadingsScoped: headings.every((h: any) => h.kind === 'h' && h.scope === 'body'),
        };
      } finally {
        session.close();
      }
    }, Array.from(bytes));

    expect(result.paraCount).toBeGreaterThanOrEqual(1);
    expect(result.allParaKindP).toBe(true);
    expect(result.headingCount).toBeGreaterThanOrEqual(1);
    expect(result.allHeadingsScoped).toBe(true);
  });

  test('replaceMatch workflow: grep → replaceMatch → text updated in place', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytesArray));
      try {
        const matches = session.grep('[A-Za-z]{5,}');
        const match = matches[0];
        const original = match.text as string;

        const edit = session.replaceMatch(match, 'ZZMARKERZZ');

        const afterMarker = session.findAllByText('ZZMARKERZZ').length;
        // The replaced span no longer matches the original text at that anchor.
        const stillHasOriginalAtAnchor = session
          .grep('[A-Za-z]{5,}')
          .some((m: any) => m.enclosingAnchor.id === match.enclosingAnchor.id &&
                            m.span.start === match.span.start &&
                            m.text === original);

        return {
          editSuccess: edit.success,
          afterMarker,
          stillHasOriginalAtAnchor,
        };
      } finally {
        session.close();
      }
    }, Array.from(bytes));

    expect(result.editSuccess).toBe(true);
    expect(result.afterMarker).toBeGreaterThanOrEqual(1);
    expect(result.stillHasOriginalAtAnchor).toBe(false);
  });
});
