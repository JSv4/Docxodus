import { test, expect, Page } from '@playwright/test';

// End-to-end Playwright spec for Issue #166 (GetEditSummary + GetDiff).
// Mirrors the .NET tests in `Docxodus.Tests/DocxSessionTests.cs`:
//   - DS280  GetEditSummary counts placeholders on an unedited doc
//   - DS282  RemainingPlaceholders alias matches FindPlaceholders
//   - DS284  GetDiff on an unedited doc returns the empty JSON array
//   - DS285  GetDiff after DeleteBlock shows a "delete" entry
//   - DS286  GetDiff after ReplaceText shows a "modify" entry
//
// Harness pattern matches `fill-placeholders.spec.ts` (#163),
// `context-boundary.spec.ts` (#164), and `delete-range-section.spec.ts` (#165)
// — inline base64 fixture, raw `Docxodus.DocxSessionBridge`, per-test
// session opened in a try/finally that calls `CloseSession`.

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

function bytesFromBase64(b64: string): number[] {
  const bin = Buffer.from(b64, 'base64');
  const out = new Array<number>(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin[i]!;
  return out;
}

// Five-paragraph fixture covering every placeholder shape `findPlaceholders`
// recognizes: a plain `[_____]`, a `$[___]` (dollar-prefix), a third blank,
// a single bracketed alternative clause, and a nested `[outer [inner] clause]`.
// Same const used by `fill-placeholders.spec.ts` (#163) — duplicated here so
// this spec stays self-contained.
const FIXTURE_BRACKET_B64 =
  'UEsDBBQAAAAIAAuPuVwARQTL/gAAAO4BAAARAAAAd29yZC9kb2N1bWVudC54bWyNUcFOwzAM/RWr4gjN4IBQtXU/MHHiVk3IJN4aqYkjx6X070nQEJdJzIcXPcXvOX7Z7r/CBJ8k2XPcNY/tpgGKlp2P510z6+nhpdn326VzbOdAUaH0x9wtu2ZUTZ0x2Y4UMLecKJa7E0tALVTOZmFxSdhSzsUuTOZps3k2AX1squUHu7WeqYJU0P5tJIgYCPgEOvoMliWxoJbnQaHDe61juzW1u6L8YLpmlMTbgiSQRxSq+rvhNvmBLSo5QL2MvNQN0uGVdfE6ZsVYYyx7EJRY6MyF3QM65+s6OIGShAwBV8CUprU9/m/OcxHB4GMkOYKdcM50TWZ+8zV/f9d/A1BLAwQUAAAAAAALj7lct/en5yoBAAAqAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3dvcmQvZG9jdW1lbnQueG1sIiBJZD0iUmNlMDAyOTVlNTM1NzQ2YzciIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAAAAALj7lc3xsCV1oBAABaAQAAEwAAAFtDb250ZW50X1R5cGVzXS54bWzvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxUeXBlcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9jb250ZW50LXR5cGVzIj48RGVmYXVsdCBFeHRlbnNpb249InhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC53b3JkcHJvY2Vzc2luZ21sLmRvY3VtZW50Lm1haW4reG1sIiAvPjxEZWZhdWx0IEV4dGVuc2lvbj0icmVscyIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1wYWNrYWdlLnJlbGF0aW9uc2hpcHMreG1sIiAvPjwvVHlwZXM+UEsBAhQDFAAAAAgAC4+5XABFBMv+AAAA7gEAABEAAAAAAAAAAAAAAKSBAAAAAHdvcmQvZG9jdW1lbnQueG1sUEsBAhQDFAAAAAAAC4+5XLf3p+cqAQAAKgEAAAsAAAAAAAAAAAAAAKSBLQEAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAAAC4+5XN8bAldaAQAAWgEAABMAAAAAAAAAAAAAAKSBgAIAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAMAAwC5AAAACwQAAAAA';

test.describe('edit-summary-and-diff (WASM bridge — Issue #166)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // OpenSession is called with `''` for the settings JSON so the .NET side
  // uses default settings — including `CaptureInitialProjection = true`,
  // which is the precondition GetDiff checks before snapshotting.

  test('getEditSummary reports placeholder counts on unedited doc', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const summary = JSON.parse(bridge.GetEditSummary(handle));
        return {
          totalAnchors: summary.totalAnchors,
          remainingCount: (summary.remainingPlaceholders ?? []).length,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.totalAnchors).toBeGreaterThan(0);
    expect(result.remainingCount).toBeGreaterThanOrEqual(4);
  });

  test('remainingPlaceholders matches findPlaceholders', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // PlaceholderKinds.All = 7, ProjectionScopes.All = 0xFFFF (use full mask).
        const kinds = 7;
        const scope = 0xFFFF;
        const remaining = JSON.parse(bridge.RemainingPlaceholders(handle, kinds)) as any[];
        const found = JSON.parse(bridge.FindPlaceholders(handle, kinds, scope, 80, 0)) as any[];
        return {
          remainingCount: remaining.length,
          foundCount: found.length,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.remainingCount).toBe(result.foundCount);
    expect(result.remainingCount).toBeGreaterThan(0);
  });

  test('getDiff returns "[]" on an unedited doc', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // DiffFormat.Json = 0.
        const diffJson = bridge.GetDiff(handle, 0);
        return { diffJson };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.diffJson).toBe('[]');
  });

  test('getDiff after DeleteBlock shows a "delete" entry for the removed anchor', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // Find the first body paragraph in document order (same pattern as
        // delete-range-section.spec.ts).
        const proj = JSON.parse(bridge.Project(handle));
        const bodyPs = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t, idx: proj.markdown.indexOf('{#' + id + '}') }))
          .filter(t => t.scope === 'body' && t.kind === 'p' && t.idx >= 0)
          .sort((a, b) => a.idx - b.idx);
        if (bodyPs.length === 0) return { error: 'no body paragraphs found' };
        const firstId = bodyPs[0].id;

        const del = JSON.parse(bridge.DeleteBlock(handle, firstId));
        const diffJson = bridge.GetDiff(handle, 0);
        return {
          deleteSuccess: del.success,
          firstId,
          diffJson,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.error).toBeUndefined();
    expect(result.deleteSuccess).toBe(true);
    expect(result.diffJson).toContain('"delete"');
    expect(result.diffJson).toContain(result.firstId);
  });

  test('getDiff throws when captureInitialProjection is disabled (opt-out path)', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const errorMessage = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const settings = JSON.stringify({ captureInitialProjection: false });
      const handle = bridge.OpenSession(bin, settings);
      try {
        bridge.GetDiff(handle, 0);
        return null; // expected throw didn't happen
      } catch (e: any) {
        return e?.message ?? String(e);
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(errorMessage).not.toBeNull();
    expect(errorMessage).toContain('CaptureInitialProjection');
  });

  test('getDiff after ReplaceText shows a "modify" entry with new content', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyPs = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t, idx: proj.markdown.indexOf('{#' + id + '}') }))
          .filter(t => t.scope === 'body' && t.kind === 'p' && t.idx >= 0)
          .sort((a, b) => a.idx - b.idx);
        if (bodyPs.length === 0) return { error: 'no body paragraphs found' };
        const firstId = bodyPs[0].id;

        const rep = JSON.parse(bridge.ReplaceText(handle, firstId, 'NEW CONTENT'));
        const diffJson = bridge.GetDiff(handle, 0);
        return {
          replaceSuccess: rep.success,
          firstId,
          diffJson,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.error).toBeUndefined();
    expect(result.replaceSuccess).toBe(true);
    expect(result.diffJson).toContain('"modify"');
    expect(result.diffJson).toContain('NEW CONTENT');
  });
});
