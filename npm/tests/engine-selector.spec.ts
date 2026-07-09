import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  const fullPath = path.join(TEST_FILES_DIR, relativePath);
  return new Uint8Array(fs.readFileSync(fullPath));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout: 30000,
  });
}

// Compare with an explicit engine selector (0 = WmlComparer, 1 = DocxDiff) via the
// low-level harness, which forwards `engine` to DocumentComparer.CompareDocuments.
async function compareWithEngine(
  page: Page,
  originalBytes: Uint8Array,
  modifiedBytes: Uint8Array,
  engine: number | undefined
): Promise<{ docxBytes?: number[]; error?: any }> {
  return await page.evaluate(
    ([original, modified, eng]) => {
      const result = (window as any).DocxodusTests.compareDocuments(
        new Uint8Array(original as number[]),
        new Uint8Array(modified as number[]),
        'Test',
        eng as number | undefined
      );
      if (result.docxBytes) {
        return { docxBytes: Array.from(result.docxBytes as Uint8Array) };
      }
      return result;
    },
    [Array.from(originalBytes), Array.from(modifiedBytes), engine] as const
  );
}

async function getRevisions(page: Page, docxBytes: number[]): Promise<{ revisions?: any[]; error?: any }> {
  return await page.evaluate((bytesArray) => {
    return (window as any).DocxodusTests.getRevisions(new Uint8Array(bytesArray));
  }, docxBytes);
}

function expectValidDocx(docxBytes: number[] | undefined) {
  expect(docxBytes).toBeDefined();
  expect(docxBytes!.length).toBeGreaterThan(1000);
  // PK zip signature.
  expect(docxBytes![0]).toBe(0x50);
  expect(docxBytes![1]).toBe(0x4b);
}

test.describe('Shared comparison-engine selector (M-B)', () => {
  const ORIGINAL = 'WC/WC001-Digits.docx';
  const MODIFIED = 'WC/WC001-Digits-Mod.docx';

  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('default engine (omitted) produces a valid redline with revisions', async ({ page }) => {
    const original = readTestFile(ORIGINAL);
    const modified = readTestFile(MODIFIED);

    const result = await compareWithEngine(page, original, modified, undefined);
    expect(result.error).toBeUndefined();
    expectValidDocx(result.docxBytes);

    const revisions = await getRevisions(page, result.docxBytes!);
    expect(revisions.error).toBeUndefined();
    expect(revisions.revisions!.length).toBeGreaterThan(0);
  });

  test('explicit WmlComparer engine (0) matches the default path', async ({ page }) => {
    const original = readTestFile(ORIGINAL);
    const modified = readTestFile(MODIFIED);

    const result = await compareWithEngine(page, original, modified, 0);
    expect(result.error).toBeUndefined();
    expectValidDocx(result.docxBytes);

    const revisions = await getRevisions(page, result.docxBytes!);
    expect(revisions.revisions!.length).toBeGreaterThan(0);
  });

  test('DocxDiff engine (1) produces a valid redline with native-markup revisions', async ({ page }) => {
    const original = readTestFile(ORIGINAL);
    const modified = readTestFile(MODIFIED);

    const result = await compareWithEngine(page, original, modified, 1);
    expect(result.error).toBeUndefined();
    expectValidDocx(result.docxBytes);

    // Both engines emit native tracked-changes markup, so the WmlComparer-based
    // revision reader counts revisions on the DocxDiff output too.
    const revisions = await getRevisions(page, result.docxBytes!);
    expect(revisions.error).toBeUndefined();
    expect(revisions.revisions!.length).toBeGreaterThan(0);
  });
});
