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

// Issue #236 — header/footer + page-number authoring through the WASM bridge.
test.describe('DocxSession header/footer authoring (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('SetFooterText + InsertPageNumberField compose, project, and round-trip', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const anchorEntries = Object.entries(proj.anchorIndex) as [string, any][];
        const firstBody = anchorEntries
          .map(([id, t]) => ({ id, ...t }))
          .filter((t) => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map((t) => ({ t, idx: proj.markdown.indexOf('{#' + t.id + '}') }))
          .filter((x) => x.idx >= 0)
          .sort((a, b) => a.idx - b.idx)[0];

        const setFooter = JSON.parse(
          bridge.SetFooterText(handle, firstBody.t.id, 'default', 'Last Updated October 2025'),
        );
        const footerAnchor: string | undefined = setFooter.created?.[0]?.id;
        const pageNum = footerAnchor
          ? JSON.parse(bridge.InsertPageNumberField(handle, footerAnchor, 'currentPage'))
          : { success: false };

        const after = JSON.parse(bridge.Project(handle));
        const saved = bridge.Save(handle);

        // Reopen the saved bytes in a fresh session to prove the footer persisted.
        const handle2 = bridge.OpenSession(saved, '');
        let reopenedHasFooter = false;
        try {
          reopenedHasFooter = JSON.parse(bridge.Project(handle2)).markdown.includes('Last Updated October 2025');
        } finally {
          bridge.CloseSession(handle2);
        }

        return {
          footerSuccess: setFooter.success,
          footerAnchor,
          footerScopeIsFtr: !!footerAnchor && footerAnchor.startsWith('p:ftr'),
          pageNumSuccess: pageNum.success,
          markdownHasFooter: after.markdown.includes('Last Updated October 2025'),
          savedBytes: saved.length,
          reopenedHasFooter,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.footerSuccess).toBe(true);
    expect(result.footerScopeIsFtr, `footerAnchor=${result.footerAnchor}`).toBe(true);
    expect(result.pageNumSuccess).toBe(true);
    expect(result.markdownHasFooter).toBe(true);
    expect(result.savedBytes).toBeGreaterThan(0);
    expect(result.reopenedHasFooter).toBe(true);
  });

  test('SetHeaderText error envelope: non-body anchor gives typed error code', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // First create a header so there is a hdr-scope anchor to (mis)use as the section anchor.
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:') || k.startsWith('h:body:'))!;
        const made = JSON.parse(bridge.SetHeaderText(handle, bodyAnchor, 'default', 'CONFIDENTIAL'));
        const headerAnchor: string = made.created[0].id;
        // A header anchor is NOT a valid section anchor → AnchorWrongKind.
        const r = JSON.parse(bridge.SetHeaderText(handle, headerAnchor, 'default', 'nope'));
        return { madeSuccess: made.success, headerScopeIsHdr: headerAnchor.startsWith('p:hdr'), success: r.success, errorCode: r.error?.code };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.madeSuccess).toBe(true);
    expect(result.headerScopeIsHdr).toBe(true);
    expect(result.success).toBe(false);
    expect(result.errorCode).toBe('anchor_wrong_kind');
  });
});
