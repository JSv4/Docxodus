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
});
