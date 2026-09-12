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

test.describe('Host-captured delivery evidence (#748)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('described batches and direct calls become a receipt the bridge verifies', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const D = (window as any).Docxodus;
      const session = D.openTypedSession(new Uint8Array(bytes), JSON.stringify({ captureDeliveryEvidence: true }));
      try {
        const anchor = /\{#((?:p|h):body:[0-9a-f]+)\}/.exec(session.project().markdown as string)![1];
        // A described transactional batch, retried once.
        const batch = () => session.executeBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            args: { anchorId: anchor, position: 'after', markdown: 'Batched.' },
            mutation: () => session.insertParagraph(anchor, 'after', 'Batched.') },
        ], 'atomic', { transactionId: 'tx-deliver', request: { insert: anchor } });
        const first = batch();
        const retry = batch();
        // A direct call: exact packages, unlabeled request.
        session.replaceText(anchor, 'Directly edited.');
        session.undo();
        session.redo();
        const status = session.getDeliveryEvidenceStatus();
        const bundle = session.buildDeliveryReceipt({ privacyProfile: 'hashAndSummary' });
        const receipt = bundle.artifacts.find((a: any) => a.artifactId === 'change-receipt');
        const artifacts: Record<string, string> = {};
        for (const artifact of bundle.artifacts) {
          if (artifact.artifactId === 'change-receipt' || !artifact.bytes) continue;
          let binary = '';
          for (const b of artifact.bytes) binary += String.fromCharCode(b);
          artifacts[artifact.artifactId] = btoa(binary);
        }
        const receiptJson = new TextDecoder().decode(receipt.bytes);
        const verification = JSON.parse(D.DocumentConverter.VerifyDeliveryReceipt(receiptJson, JSON.stringify(artifacts)));
        const payload = JSON.parse(receiptJson).payload;
        return {
          replayIdentical: JSON.stringify(first) === JSON.stringify(retry),
          status,
          bundleStatus: bundle.status,
          verified: bundle.verified,
          receiptAvailable: receipt?.availability,
          verificationValid: verification.isValid,
          findings: verification.findings,
          transactionIds: payload.transactions.map((t: any) => t.transactionId),
          operations: payload.transactions.map((t: any) => t.operations[0].tool + '/' + t.operations[0].action),
          lineage: payload.lineage.map((l: any) => l.action),
          deliveredMarkdown: (() => {
            const final = bundle.artifacts.find((a: any) => a.artifactId === 'final-docx').bytes as Uint8Array;
            const delivered = D.openTypedSession(final);
            try { return delivered.project().markdown as string; } finally { delivered.close(); }
          })(),
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.replayIdentical).toBe(true);
    expect(outcome.status.enabled).toBe(true);
    expect(outcome.status.unavailableReason).toBeNull();
    expect(outcome.status.transactionCount).toBe(2);
    expect(outcome.status.lineageEventCount).toBe(2);
    expect(outcome.status.unlabeledTransactionCount).toBe(1);
    expect(outcome.bundleStatus).toBe('complete');
    expect(outcome.verified).toBe(true);
    expect(outcome.receiptAvailable).toBe('available');
    expect(outcome.verificationValid, outcome.findings.join('; ')).toBe(true);
    expect(outcome.transactionIds).toEqual(['tx-deliver', null]);
    expect(outcome.operations).toEqual(['docx_create/insert_paragraph', 'docx_session/unlabeled_mutation']);
    expect(outcome.lineage).toEqual(['undo', 'redo']);
    expect(outcome.deliveredMarkdown).toContain('Batched.');
    expect(outcome.deliveredMarkdown).toContain('Directly edited.');
  });

  test('a session that does not capture evidence says so instead of claiming a history', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const anchor = /\{#((?:p|h):body:[0-9a-f]+)\}/.exec(session.project().markdown as string)![1];
        session.replaceText(anchor, 'Unrecorded.');
        const status = session.getDeliveryEvidenceStatus();
        const bundle = session.buildDeliveryReceipt();
        const receipt = bundle.artifacts.find((a: any) => a.artifactId === 'change-receipt');
        return {
          enabled: status.enabled,
          reason: status.unavailableReason,
          bundleStatus: bundle.status,
          receiptAvailability: receipt?.availability,
          receiptHasBytes: receipt?.bytes !== undefined,
          evidenceReason: bundle.evidence?.unavailableReason,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.enabled).toBe(false);
    expect(outcome.reason).toContain('CaptureDeliveryEvidence');
    expect(outcome.bundleStatus).toBe('incomplete');
    expect(outcome.receiptAvailability).toBe('unavailable');
    expect(outcome.receiptHasBytes).toBe(false);
    expect(outcome.evidenceReason).toContain('CaptureDeliveryEvidence');
  });
});
