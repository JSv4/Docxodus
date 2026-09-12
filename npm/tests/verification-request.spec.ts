import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixture = new Uint8Array(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC001-5DayTourPlanTemplate.docx'),
));

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// The harness exposes the typed session and the raw bridges, not the async core wrappers, so
// the request shape is exercised through `session.verifyDeliverable(request)` — the same
// serializer the core `verifyDeliverable(document, baseline, request)` wrapper uses.
test.describe('full deliverable-verification requests (#747)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('a stale companion and an unexpected package change fail the deliverable', async ({ page }) => {
    const outcome = await page.evaluate(async (bytes: number[]) => {
      const baseline = new Uint8Array(bytes);
      const session = (window as any).Docxodus.openTypedSession(baseline);
      try {
        const projection = session.project();
        const anchor = (Object.entries(projection.anchorIndex) as [string, any][])
          .find(([, value]) => value.scope === 'body' && value.kind === 'p')![0];
        session.replaceText(anchor, 'Changed for delivery.');
        // The companion claims to have been rendered from the BASELINE package, which is not
        // the package being delivered.
        const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', baseline)))
          .map(b => b.toString(16).padStart(2, '0')).join('');
        const report = session.verifyDeliverable({
          options: { failOnUnexpectedChanges: true },
          expectedPackageChanges: [],
          companionArtifacts: [{
            artifactId: 'pdf-1', role: 'pdf', mediaType: 'application/pdf',
            bytes: new TextEncoder().encode('%PDF-1.4\n%stale\n'),
            pageCount: 1, rendererFingerprint: 'renderer/1.0',
            sourcePackageDigest: { algorithm: 'SHA-256', value: digest },
            renderDiagnostics: [{ kind: 'missingFont', message: 'Aptos substituted' }],
          }],
        });
        const plain = session.verifyDeliverable();
        return {
          codes: report.findings.map((f: any) => f.code),
          decision: report.decision,
          artifacts: report.companionArtifacts.map((a: any) => a.artifactId),
          plainDecision: plain.decision,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.codes).toContain('delta.package_change_unexpected');
    expect(outcome.codes).toContain('artifact.source_digest_mismatch');
    expect(outcome.decision).toBe('failed');
    expect(outcome.artifacts).toEqual(['pdf-1']);
    expect(outcome.plainDecision).not.toBe('failed');
  });

  test('a session verifies its own clean save under the request, and unknown options are rejected', async ({ page }) => {
    const outcome = await page.evaluate(async (bytes: number[]) => {
      const D = (window as any).Docxodus;
      const session = D.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const anchor = (Object.entries(projection.anchorIndex) as [string, any][])
          .find(([, value]) => value.scope === 'body' && value.kind === 'p')![0];
        session.replaceText(anchor, 'Changed for delivery.');
        // Expectations come from the same byte-level comparison the verifier runs.
        const expected = JSON.parse(D.DocxDiffBridge.GetSemanticChangesJson(
          new Uint8Array(bytes), session.save(), ''));
        const approved = session.verifyDeliverable({
          options: { failOnUnexpectedChanges: true },
          expectedSemanticChanges: expected,
        });
        let rejected = 'none';
        try { session.verifyDeliverable({ options: { failOnUnexpectedChange: true } as any }); }
        catch (error) { rejected = error instanceof Error ? error.message : String(error); }
        return {
          codes: approved.findings.map((f: any) => f.code),
          rejected,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.codes).not.toContain('delta.semantic_change_unexpected');
    expect(outcome.codes).not.toContain('delta.semantic_change_missing');
    expect(outcome.rejected).toContain('failOnUnexpectedChange');
  });
});
