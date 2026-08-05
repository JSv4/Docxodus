// Trim-safety canaries. The WASM build IL-trims Docxodus and DocumentFormat.OpenXml
// (wasm/DocxodusWasm/DocxodusWasm.csproj), and exactly two bridge-reachable paths
// depend on things the trimmer cannot see statically:
//   1. PtOpenXmlUtil.GetPackage() — reflection into the SDK's features chain, hit
//      when a WmlComparer-engine compare copies image/media parts. Pinned by
//      wasm/DocxodusWasm/ILLink.Descriptors.xml; this spec proves the pin works.
//   2. OpenXmlValidator — reached only through Raw ops with validateRawOps on; the
//      invalid-payload assertion proves the validator actually ran rather than
//      being trimmed to a no-op.
// If either test fails after a DocumentFormat.OpenXml or SDK bump, suspect the
// descriptor before suspecting this spec. See docs/architecture/wasm-packaging.md.
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

test.describe('Trim validation (paths uncovered by main suite)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('WmlComparer engine compare with images exercises reflective GetPackage()', async ({ page }) => {
    // WmlComparer.CoalesceRecurse's w:drawing handling calls the reflection-based
    // OpenXmlPackage.GetPackage() when copying media parts — the only bridge-reachable
    // route through PtOpenXmlUtil's reflection. engine: 0 = ComparisonEngine.WmlComparer.
    const left = readTestFile('HC042-Image-Png.docx');
    const right = readTestFile('HC006-Test-01.docx');

    const result = await page.evaluate(
      async (args: { left: number[]; right: number[] }) => {
        const comparer = (window as any).Docxodus.DocumentComparer;
        const redline = comparer.CompareDocumentsWithOptions(
          new Uint8Array(args.left), new Uint8Array(args.right),
          'TrimCheck', 0.15, false, /* engine: WmlComparer */ 0);
        const reverse = comparer.CompareDocumentsWithOptions(
          new Uint8Array(args.right), new Uint8Array(args.left),
          'TrimCheck', 0.15, false, 0);
        return { redlineLen: redline.length, reverseLen: reverse.length };
      },
      { left: Array.from(left), right: Array.from(right) },
    );

    // CompareDocumentsWithOptions returns Array.Empty on any exception —
    // a trimmed-away reflection target surfaces here as length 0.
    expect(result.redlineLen).toBeGreaterThan(0);
    expect(result.reverseLen).toBeGreaterThan(0);
  });

  test('rawInsertXml with validateRawOps exercises OpenXmlValidator', async ({ page }) => {
    const bytes = readTestFile('HC006-Test-01.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(new Uint8Array(bytesArray), JSON.stringify({ validateRawOps: true }));
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const anchorEntries = Object.entries(proj.anchorIndex) as [string, any][];
        const firstBody = anchorEntries
          .map(([id, t]) => ({ id, ...t }))
          .filter((t: any) => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map((t: any) => ({ t, idx: proj.markdown.indexOf('{#' + t.id + '}') }))
          .filter((x: any) => x.idx >= 0)
          .sort((a: any, b: any) => a.idx - b.idx)[0];

        const xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:r><w:t>trim-validation paragraph</w:t></w:r></w:p>';
        const insertResult = JSON.parse(bridge.RawInsertXml(handle, firstBody.t.id, 'after', xml));

        // An invalid payload must be REJECTED by the validator — proves OpenXmlValidator
        // actually ran rather than being trimmed to a no-op.
        const badXml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:bogusElement/></w:p>';
        const badResult = JSON.parse(bridge.RawInsertXml(handle, firstBody.t.id, 'after', badXml));

        return {
          insertSuccess: insertResult.success,
          insertError: insertResult.error,
          badRejected: badResult.success === false,
          badError: badResult.error,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.insertSuccess).toBe(true);
    expect(result.badRejected).toBe(true);
  });
});
