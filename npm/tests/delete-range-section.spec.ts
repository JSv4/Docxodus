import { test, expect, Page } from '@playwright/test';

// End-to-end Playwright spec for Issue #165 (DeleteRange + DeleteSection).
// Mirrors the .NET tests in `Docxodus.Tests/DocxSessionTests.cs`:
//   - DS260 DeleteRange removes contiguous body blocks in one call
//   - DS261 DeleteRange rejects reversed (from > to) anchors
//   - DS264 DeleteRange refuses anchors that live in different parts (body + fn)
//   - DS265 DeleteRange Undo restores the entire deleted range
//   - DS267 DeleteSection removes a heading-bounded section
//   - DS268 DeleteSection on the last heading extends to end of parent
//
// Harness pattern matches `fill-placeholders.spec.ts` (#163) and
// `context-boundary.spec.ts` (#164) — inline base64 fixtures, raw
// `Docxodus.DocxSessionBridge`, per-test session opened in a try/finally.

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

function bytesFromBase64(b64: string): number[] {
  const bin = Buffer.from(b64, 'base64');
  const out = new Array<number>(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin[i]!;
  return out;
}

// ─── Fixtures ────────────────────────────────────────────────────────────
//
// All three fixtures were produced via a throwaway .NET program that mirrors
// the BuildDocFiveBodyParagraphs / BuildDocWithFootnotes / BuildDocWithHeadingSections
// helpers in `Docxodus.Tests/DocxSessionTests.cs` (using DocumentFormat.OpenXml
// 3.4.1, matching the rest of the suite). Bytes were base64-encoded inline so
// this spec stays self-contained.

// Five-paragraph body (A..E) — covers assertions 1, 2, and 4 (DS260/261/265).
const FIXTURE_FIVE_B64 =
  'UEsDBBQAAAAIAE+juVzV59WfpAAAAHMBAAARAAAAd29yZC9kb2N1bWVudC54bWyVkEkOwjAMRa8SZU9dWCBUdVAZ9lwhJOkgNXHkpBRuT4KE2MCim2fZfvqSXTYPM7G7Jj+irfg2yznTVqIabV/xOXSbA2/qcikUytloG1j0rS+Wig8huALAy0Eb4TN02sZdh2REiC31sCApRyi19zHOTLDL8z0YMVqeIm+onqm6BEoI9VWQ6Em4gbUlpEEiven+u8cV7mmFe17hXn648DkSvg+sX1BLAwQUAAAAAABPo7lczgKV/CoBAAAqAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3dvcmQvZG9jdW1lbnQueG1sIiBJZD0iUjg2ZTZlM2NiNjU2OTRkYjMiIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAAAABPo7lc3xsCV1oBAABaAQAAEwAAAFtDb250ZW50X1R5cGVzXS54bWzvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxUeXBlcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9jb250ZW50LXR5cGVzIj48RGVmYXVsdCBFeHRlbnNpb249InhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC53b3JkcHJvY2Vzc2luZ21sLmRvY3VtZW50Lm1haW4reG1sIiAvPjxEZWZhdWx0IEV4dGVuc2lvbj0icmVscyIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1wYWNrYWdlLnJlbGF0aW9uc2hpcHMreG1sIiAvPjwvVHlwZXM+UEsBAhQDFAAAAAgAT6O5XNXn1Z+kAAAAcwEAABEAAAAAAAAAAAAAAKSBAAAAAHdvcmQvZG9jdW1lbnQueG1sUEsBAhQDFAAAAAAAT6O5XM4ClfwqAQAAKgEAAAsAAAAAAAAAAAAAAKSB0wAAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAAAT6O5XN8bAldaAQAAWgEAABMAAAAAAAAAAAAAAKSBJgIAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAMAAwC5AAAAsQMAAAAA';

// Body paragraph + FootnotesPart with a user-authored footnote — covers
// assertion 3 (DS264: anchors_not_adjacent when from/to span body and fn).
const FIXTURE_FOOTNOTES_B64 =
  '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';

// Heading1/p/Heading2/p/Heading1/p — covers assertions 5 and 6 (DS267/268).
// Includes a real StyleDefinitionsPart so the projection recognizes Heading1/2.
const FIXTURE_HEADINGS_B64 =
  'UEsDBBQAAAAIAE+juVxfRnBz0wAAAB4CAAARAAAAd29yZC9kb2N1bWVudC54bWydkc9uwjAMh18lyp267WGaqraI226bBC8QUgOVmjhKDIG3X1o2uPBPXGxZ/vTTJ7ueH80gDuhDT7aRRZZLgVZT19ttI/e8mX3KeVvHqiO9N2hZJN6GKjZyx+wqgKB3aFTIyKFNuw15oziNfguRfOc8aQwhxZkByjz/AKN6K8fINXWnsbup/PipLfk0oIjVQQ2N/EI1ihRSQFvDBZoKt0vUnKzFt8VxyRPiz+Al9491yitRZMV98JFA+VQgW7yqsHhX4skVVpFeUyhvCsD/P+D66/YXUEsDBBQAAAAIAE+juVyMoru/kQAAAA8BAAAPAAAAd29yZC9zdHlsZXMueG1slY47DsIwDEB3ThH5ALjtwBC1meEYVhvSSvkpjhp6e4KIkBgY2Gw9vyePRXI+rGbxcNazLBOsOUeJyPOqHfE5RO0ru4fkKNc1GSwhLTGFWTNv3jiLQ9dd0NHmQZ2EGFtTFJmPqCeIlMgkiiuIhm7LBFdNS9V7UFXw5F73O9n6wBuIHlCN2Iz/w8PP8PAd/oysnlBLAwQUAAAAAABPo7lc/Jx4xyoBAAAqAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3dvcmQvZG9jdW1lbnQueG1sIiBJZD0iUjUzZGE0ZDk2MWMxMjQ2MTEiIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAACABPo7lcLA+kebYAAAAgAQAAHAAAAHdvcmQvX3JlbHMvZG9jdW1lbnQueG1sLnJlbHONzzsOwjAMBuCrRN6py1MtatqFhRVxgZA4bUWbREl4nY2BI3EFMjCAxMDo/7c/yc/7o2qu48DO5ENvDYdplgMjI63qTcvhFPWkgKaudjSImDZC17vA0okJHLoY3RoxyI5GETLryKRGWz+KmEbfohPyKFrCWZ6v0H8a8G2y/c3RP6LVupe0sfI0kok/YAzxNlAAthe+pcgBL9ard5olDdhWcdiVSpWHgopyuZgvZKmBYV3h15/1C1BLAwQUAAAACABPo7lcJ8XOvukAAADaAQAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytkTtOxDAQhq9iuUWxAwVCKMkWQAsUXMByJomFX/JMlt2zUXAkrsAkQSkQDRKl/T++Gfvz/aM5nIIXRyjoUmzlpaqlgGhT7+LYypmG6kYeuublnAEFWyO2ciLKt1qjnSAYVClDZGVIJRjiYxl1NvbVjKCv6vpa2xQJIlW0dMiuuYfBzJ7Ew4mvNyzHpbjbfAuqlSZn76whlvUx9j8gVRoGZ6FPdg4cUW+p9LkkC4g8d/BqV4Jx8WKt17+SC3j8G/p7N8XJ1YOTy7gjnvgpi+tBPJtCjyZwn16m00hnD6j+fdOtd+fr9ae6L1BLAQIUAxQAAAAIAE+juVxfRnBz0wAAAB4CAAARAAAAAAAAAAAAAACkgQAAAAB3b3JkL2RvY3VtZW50LnhtbFBLAQIUAxQAAAAIAE+juVyMoru/kQAAAA8BAAAPAAAAAAAAAAAAAACkgQIBAAB3b3JkL3N0eWxlcy54bWxQSwECFAMUAAAAAABPo7lc/Jx4xyoBAAAqAQAACwAAAAAAAAAAAAAApIHAAQAAX3JlbHMvLnJlbHNQSwECFAMUAAAACABPo7lcLA+kebYAAAAgAQAAHAAAAAAAAAAAAAAApIETAwAAd29yZC9fcmVscy9kb2N1bWVudC54bWwucmVsc1BLAQIUAxQAAAAIAE+juVwnxc6+6QAAANoBAAATAAAAAAAAAAAAAACkgQMEAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAFAAUAQAEAAB0FAAAAAA==';

test.describe('delete-range-section (WASM bridge — Issue #165)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('DeleteRange removes contiguous body blocks in one call', async ({ page }) => {
    // Mirrors DS260: delete paragraphs B..D (anchors[1]..anchors[3]) in the
    // five-paragraph fixture by passing anchors[1] as `from` and anchors[4]
    // as `toExclusive`.
    const bytes = bytesFromBase64(FIXTURE_FIVE_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        // Body paragraphs in document order — sort by their position in the
        // markdown so A..E come out in order.
        const bodyPs = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t, idx: proj.markdown.indexOf('{#' + id + '}') }))
          .filter(t => t.scope === 'body' && t.kind === 'p' && t.idx >= 0)
          .sort((a, b) => a.idx - b.idx);

        const fromId = bodyPs[1].id;
        const toExclusiveId = bodyPs[4].id;
        const r = JSON.parse(bridge.DeleteRange(handle, fromId, toExclusiveId));
        const after = JSON.parse(bridge.Project(handle));
        return {
          paragraphCount: bodyPs.length,
          success: r.success,
          errorCode: r.error?.code,
          removedCount: (r.removed ?? []).length,
          markdown: after.markdown,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.paragraphCount).toBe(5);
    expect(result.success).toBe(true);
    expect(result.errorCode).toBeUndefined();
    expect(result.removedCount).toBe(3);
    expect(result.markdown).toContain('Paragraph A');
    expect(result.markdown).not.toContain('Paragraph B');
    expect(result.markdown).not.toContain('Paragraph C');
    expect(result.markdown).not.toContain('Paragraph D');
    expect(result.markdown).toContain('Paragraph E');
  });

  test('DeleteRange rejects from > to with invalid_position', async ({ page }) => {
    // Mirrors DS261: reversed arguments (from = E, to = B) must fail with
    // EditErrorCode.InvalidPosition → snake-cased "invalid_position".
    const bytes = bytesFromBase64(FIXTURE_FIVE_B64);

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

        // from = anchors[4] (E), to = anchors[1] (B) — reversed.
        const r = JSON.parse(bridge.DeleteRange(handle, bodyPs[4].id, bodyPs[1].id));
        return { success: r.success, errorCode: r.error?.code };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.success).toBe(false);
    expect(result.errorCode).toBe('invalid_position');
  });

  test('DeleteRange rejects anchors in different parts with anchors_not_adjacent', async ({ page }) => {
    // Mirrors DS264: body anchor + fn anchor live in different package parts
    // so DeleteRange must refuse with EditErrorCode.AnchorsNotAdjacent →
    // snake-cased "anchors_not_adjacent".
    const bytes = bytesFromBase64(FIXTURE_FOOTNOTES_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const entries = Object.entries(proj.anchorIndex) as [string, any][];
        const bodyAnchor = entries.find(([, t]) => t.scope === 'body' && t.kind === 'p')?.[0];
        const fnAnchor = entries.find(([, t]) => t.scope === 'fn' && t.kind === 'fn')?.[0];
        if (!bodyAnchor || !fnAnchor) {
          return { success: null, errorCode: null, bodyAnchor, fnAnchor };
        }
        const r = JSON.parse(bridge.DeleteRange(handle, bodyAnchor, fnAnchor));
        return { success: r.success, errorCode: r.error?.code, bodyAnchor, fnAnchor };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.bodyAnchor, 'body paragraph anchor must exist').toBeTruthy();
    expect(result.fnAnchor, 'fn anchor must exist').toBeTruthy();
    expect(result.success).toBe(false);
    expect(result.errorCode).toBe('anchors_not_adjacent');
  });

  test('DeleteRange Undo restores the entire deleted range', async ({ page }) => {
    // Mirrors DS265: after a successful DeleteRange, a single Undo brings
    // every removed paragraph back.
    const bytes = bytesFromBase64(FIXTURE_FIVE_B64);

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

        const r = JSON.parse(bridge.DeleteRange(handle, bodyPs[1].id, bodyPs[4].id));
        const undidOk = bridge.Undo(handle);
        const after = JSON.parse(bridge.Project(handle));
        return {
          success: r.success,
          undidOk,
          markdown: after.markdown,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.success).toBe(true);
    expect(result.undidOk).toBe(true);
    expect(result.markdown).toContain('Paragraph A');
    expect(result.markdown).toContain('Paragraph B');
    expect(result.markdown).toContain('Paragraph C');
    expect(result.markdown).toContain('Paragraph D');
    expect(result.markdown).toContain('Paragraph E');
  });

  test('DeleteSection removes a heading-bounded section', async ({ page }) => {
    // Mirrors DS267: delete "Section One" (Heading1). Expected to take the
    // heading + "para 1.1" + "Section One.A" (Heading2 nested under it) +
    // "para 1.A.1", stopping at the next Heading1 ("Section Two").
    const bytes = bytesFromBase64(FIXTURE_HEADINGS_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const entries = Object.entries(proj.anchorIndex) as [string, any][];
        // Locate the "Section One" heading anchor by its preview text.
        const sectionOne = entries
          .map(([id, t]) => ({ id, ...t }))
          .find(t => t.scope === 'body' && t.kind === 'h' && (t.textPreview ?? '') === 'Section One');
        if (!sectionOne) {
          return { found: false, beforeMd: proj.markdown };
        }
        const r = JSON.parse(bridge.DeleteSection(handle, sectionOne.id));
        const after = JSON.parse(bridge.Project(handle));
        return {
          found: true,
          success: r.success,
          errorCode: r.error?.code,
          removedCount: (r.removed ?? []).length,
          markdown: after.markdown,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.found, 'Section One heading anchor must exist').toBe(true);
    expect(result.success).toBe(true);
    expect(result.errorCode).toBeUndefined();
    expect(result.removedCount).toBeGreaterThanOrEqual(4);
    expect(result.markdown).not.toContain('Section One.A');
    expect(result.markdown).not.toContain('para 1.A.1');
    expect(result.markdown).not.toContain('para 1.1');
    expect(result.markdown).toContain('Section Two');
    expect(result.markdown).toContain('para 2.1');
  });

  test('DeleteSection on the last heading extends to end of parent', async ({ page }) => {
    // Mirrors DS268: "Section Two" is the last Heading1. DeleteSection must
    // remove the heading + every following sibling up to the end of the body —
    // here that's "Section Two" and "para 2.1".
    const bytes = bytesFromBase64(FIXTURE_HEADINGS_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const entries = Object.entries(proj.anchorIndex) as [string, any][];
        const sectionTwo = entries
          .map(([id, t]) => ({ id, ...t }))
          .find(t => t.scope === 'body' && t.kind === 'h' && (t.textPreview ?? '') === 'Section Two');
        if (!sectionTwo) {
          return { found: false, beforeMd: proj.markdown };
        }
        const r = JSON.parse(bridge.DeleteSection(handle, sectionTwo.id));
        const after = JSON.parse(bridge.Project(handle));
        return {
          found: true,
          success: r.success,
          errorCode: r.error?.code,
          markdown: after.markdown,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.found, 'Section Two heading anchor must exist').toBe(true);
    expect(result.success).toBe(true);
    expect(result.errorCode).toBeUndefined();
    expect(result.markdown).toContain('Section One');
    expect(result.markdown).toContain('para 1.1');
    expect(result.markdown).not.toContain('Section Two');
    expect(result.markdown).not.toContain('para 2.1');
  });
});
