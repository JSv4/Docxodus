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

async function convertToMarkdown(
  page: Page,
  bytes: Uint8Array,
  settings: Record<string, unknown> = {}
): Promise<{ Markdown?: string; AnchorIndex?: Record<string, unknown>; error?: { message: string } }> {
  return await page.evaluate(
    ([bytesArray, settingsArg]) =>
      (window as any).DocxodusTests.convertToMarkdown(new Uint8Array(bytesArray), settingsArg),
    [Array.from(bytes), settings] as [number[], Record<string, unknown>]
  );
}

test.describe('Markdown projection (WASM + npm bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('convertWmlToMarkdown returns markdown and anchor index for HC001', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');
    const result = await convertToMarkdown(page, bytes);

    expect(result.error).toBeUndefined();
    expect(result.Markdown).toBeDefined();
    expect(result.Markdown!.startsWith('# Document')).toBe(true);
    expect(result.AnchorIndex).toBeDefined();
    expect(Object.keys(result.AnchorIndex!).length).toBeGreaterThan(0);
  });

  test('AnchorMode=None suppresses inline anchor tokens', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');
    const result = await convertToMarkdown(page, bytes, { anchorMode: 2 /* None */ });
    expect(result.error).toBeUndefined();
    expect(result.Markdown).toBeDefined();
    expect(result.Markdown!.includes('{#p:')).toBe(false);
  });

  test('Body-only scope drops Headers/Footers sections', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const result = await convertToMarkdown(page, bytes, { scopes: 1 /* Body */ });
    expect(result.error).toBeUndefined();
    expect(result.Markdown).toBeDefined();
    expect(result.Markdown!.includes('# Headers')).toBe(false);
    expect(result.Markdown!.includes('# Footers')).toBe(false);
  });
});
