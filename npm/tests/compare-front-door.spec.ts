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

// Compare through the low-level harness, which forwards to
// DocumentComparer.CompareDocuments. The engine selector this file used to sweep was
// removed in v11.0.0 along with the second engine; what survives is the behaviour the
// sweep was really pinning.
async function compare(
  page: Page,
  originalBytes: Uint8Array,
  modifiedBytes: Uint8Array
): Promise<{ docxBytes?: number[]; error?: any }> {
  return await page.evaluate(
    ([original, modified]) => {
      const result = (window as any).DocxodusTests.compareDocuments(
        new Uint8Array(original as number[]),
        new Uint8Array(modified as number[]),
        'Test'
      );
      if (result.docxBytes) {
        return { docxBytes: Array.from(result.docxBytes as Uint8Array) };
      }
      return result;
    },
    [Array.from(originalBytes), Array.from(modifiedBytes)] as const
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

test.describe('Shared comparison front door', () => {
  const ORIGINAL = 'WC/WC001-Digits.docx';
  const MODIFIED = 'WC/WC001-Digits-Mod.docx';

  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('produces a valid redline whose revisions are readable back', async ({ page }) => {
    const result = await compare(page, readTestFile(ORIGINAL), readTestFile(MODIFIED));

    expect(result.error).toBeUndefined();
    expectValidDocx(result.docxBytes);

    const revisions = await getRevisions(page, result.docxBytes!);
    expect(revisions.error).toBeUndefined();
    expect(revisions.revisions!.length).toBeGreaterThan(0);
  });

  test('byte-identical inputs preserve the exact package', async ({ page }) => {
    const original = readTestFile(ORIGINAL);

    const result = await compare(page, original, original);

    expect(result.error).toBeUndefined();
    expectValidDocx(result.docxBytes);
    expect(result.docxBytes).toEqual(Array.from(original));
  });
});
