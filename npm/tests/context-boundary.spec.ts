import { test, expect, Page } from '@playwright/test';
import { ContextBoundary, PlaceholderKinds } from '../src/types.js';

// End-to-end Playwright spec for Issue #164 (boundary-aware Grep context windows).
// Mirrors the .NET tests DS250–DS255 in `Docxodus.Tests/DocxSessionTests.cs`:
//   - DS250 default ContextChars widened to 80
//   - DS252 boundary: Bracket stops at `[`/`]`
//   - DS253 boundary: Sentence stops at sentence terminators
//   - DS254 boundary: Comma stops at `,`
//   - DS255 FindPlaceholders honors boundary parameter end-to-end
//
// Harness pattern matches `fill-placeholders.spec.ts` (#163) — inline base64
// fixtures, raw `Docxodus.DocxSessionBridge`, per-test session opened in a
// try/finally.

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

function bytesFromBase64(b64: string): number[] {
  const bin = Buffer.from(b64, 'base64');
  const out = new Array<number>(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin[i]!;
  return out;
}

// Single combined fixture: three paragraphs covering every boundary mode the
// spec exercises. Generated via a throwaway .NET program that calls the same
// `WordprocessingDocument.Create` shape as the `BuildDocWith*` helpers in
// `Docxodus.Tests/DocxSessionTests.cs` (paragraphs lifted verbatim):
//
//   ¶1 adjacent placeholders:
//     "The address is [STREET], in the City of [CITY], County of [COUNTY], in the State of Delaware."
//   ¶2 multi-sentence:
//     "This is the first sentence. Here is some text with a [FILL] placeholder. And a third sentence after."
//   ¶3 comma-separated:
//     "Item one is [A], item two is [B], and item three is [C]."
const FIXTURE_COMBINED_B64 =
  'UEsDBBQAAAAIAEmauVyUnwmSKQEAAPMBAAARAAAAd29yZC9kb2N1bWVudC54bWyFUctOwzAQ/JVVzhAXDghVfaiEIipVINFwQFEPJt7UlmI7Wm9J+/fYKYUDSFzG8szujHc9mR9sCx9IwXg3za7yUQboaq+M202zPTeXt9l8NunHytd7i44h1rsw7qeZZu7GQoRao5Uh9x26qDWerOR4pZ3oPamOfI0hRDvbiuvR6EZYaVyWLN+9OqazS0AJeFZqBKkUxRYwAapN+bJcltsLMA44aoXhI/gGqmJVvkW68Ht3Zp5fnwbuq3TDkjEp99jKXhLmE5EiEtKA3e90M8Sm9sZQYAhx5LgPzOERCZMWvEVgPDD0hjVIqB5W6/UWulbWqH2rkHJYOBUV1obUtwXIhqP27yNWjBa8G8KqRZonEdz7gbiLhIzuJ1ITnuqK7V/G4rxk8fOBs09QSwMEFAAAAAAASZq5XAPd0LQqAQAAKgEAAAsAAABfcmVscy8ucmVsc++7vzw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9InV0Zi04Ij8+PFJlbGF0aW9uc2hpcHMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvcmVsYXRpb25zaGlwcyI+PFJlbGF0aW9uc2hpcCBUeXBlPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzL29mZmljZURvY3VtZW50IiBUYXJnZXQ9Ii93b3JkL2RvY3VtZW50LnhtbCIgSWQ9IlJhZWZkNjAwZmI0ZTE0MDVmIiAvPjwvUmVsYXRpb25zaGlwcz5QSwMEFAAAAAAASZq5XN8bAldaAQAAWgEAABMAAABbQ29udGVudF9UeXBlc10ueG1s77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48VHlwZXMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvY29udGVudC10eXBlcyI+PERlZmF1bHQgRXh0ZW5zaW9uPSJ4bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQud29yZHByb2Nlc3NpbmdtbC5kb2N1bWVudC5tYWluK3htbCIgLz48RGVmYXVsdCBFeHRlbnNpb249InJlbHMiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtcGFja2FnZS5yZWxhdGlvbnNoaXBzK3htbCIgLz48L1R5cGVzPlBLAQIUAxQAAAAIAEmauVyUnwmSKQEAAPMBAAARAAAAAAAAAAAAAACkgQAAAAB3b3JkL2RvY3VtZW50LnhtbFBLAQIUAxQAAAAAAEmauVwD3dC0KgEAACoBAAALAAAAAAAAAAAAAACkgVgBAABfcmVscy8ucmVsc1BLAQIUAxQAAAAAAEmauVzfGwJXWgEAAFoBAAATAAAAAAAAAAAAAACkgasCAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAADAAMAuQAAADYEAAAAAA==';

test.describe('context-boundary (WASM bridge — Issue #164)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('default contextChars widens to 80 (adjacent placeholders)', async ({ page }) => {
    // Mirrors DS250: with the default 80-char window, [STREET]'s ContextAfter
    // — "..., in the City of [CITY], County of [COUNTY], in the State of Delaware."
    // — exceeds the previous 40-char cap.
    const bytes = bytesFromBase64(FIXTURE_COMBINED_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // Empty options → bridge defaults (contextChars=80, boundary=Char).
        const matches = JSON.parse(bridge.Grep(handle, '\\[STREET\\]', '')) as any[];
        return {
          count: matches.length,
          contextAfterLen: matches[0]?.contextAfter?.length ?? 0,
          contextAfter: matches[0]?.contextAfter ?? '',
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.count).toBe(1);
    expect(result.contextAfterLen).toBeGreaterThan(40);
  });

  test('boundary: Bracket stops at "[" / "]" (Grep)', async ({ page }) => {
    // Mirrors DS252: contexts around [CITY] are bracket-free with Bracket mode.
    const bytes = bytesFromBase64(FIXTURE_COMBINED_B64);
    const boundary = ContextBoundary.Bracket;

    const result = await page.evaluate(async ({ bytesArray, boundary }) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const matches = JSON.parse(bridge.Grep(
          handle,
          '\\[CITY\\]',
          JSON.stringify({ scope: 1, contextChars: 80, boundary }),
        )) as any[];
        return {
          count: matches.length,
          contextBefore: matches[0]?.contextBefore ?? '',
          contextAfter: matches[0]?.contextAfter ?? '',
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: bytes, boundary });

    expect(result.count).toBe(1);
    expect(result.contextBefore).not.toContain('[');
    expect(result.contextBefore).not.toContain(']');
    expect(result.contextAfter).not.toContain('[');
    expect(result.contextAfter).not.toContain(']');
    // Lock in the exact slices from DS252.
    expect(result.contextBefore).toBe(', in the City of ');
    expect(result.contextAfter).toBe(', County of ');
  });

  test('boundary: Sentence stops at "." / ";" terminators', async ({ page }) => {
    // Mirrors DS253: a multi-sentence paragraph with [FILL] in the middle
    // sentence; Sentence mode keeps ContextBefore/After inside one sentence.
    const bytes = bytesFromBase64(FIXTURE_COMBINED_B64);
    const boundary = ContextBoundary.Sentence;

    const result = await page.evaluate(async ({ bytesArray, boundary }) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const matches = JSON.parse(bridge.Grep(
          handle,
          '\\[FILL\\]',
          JSON.stringify({ scope: 1, contextChars: 200, boundary }),
        )) as any[];
        return {
          count: matches.length,
          contextBefore: matches[0]?.contextBefore ?? '',
          contextAfter: matches[0]?.contextAfter ?? '',
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: bytes, boundary });

    expect(result.count).toBe(1);
    expect(result.contextBefore).not.toContain('.');
    expect(result.contextAfter).not.toContain('.');
    expect(result.contextBefore).not.toContain(';');
    expect(result.contextAfter).not.toContain(';');
  });

  test('boundary: Comma stops at ","', async ({ page }) => {
    // Mirrors DS254: enumeration paragraph with three comma-separated items;
    // grep [B], Comma mode keeps the context inside the middle item.
    const bytes = bytesFromBase64(FIXTURE_COMBINED_B64);
    const boundary = ContextBoundary.Comma;

    const result = await page.evaluate(async ({ bytesArray, boundary }) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const matches = JSON.parse(bridge.Grep(
          handle,
          '\\[B\\]',
          JSON.stringify({ scope: 1, contextChars: 200, boundary }),
        )) as any[];
        return {
          count: matches.length,
          contextBefore: matches[0]?.contextBefore ?? '',
          contextAfter: matches[0]?.contextAfter ?? '',
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: bytes, boundary });

    expect(result.count).toBe(1);
    expect(result.contextBefore).not.toContain(',');
    expect(result.contextAfter).not.toContain(',');
  });

  test('FindPlaceholders honors boundary parameter end-to-end', async ({ page }) => {
    // Mirrors DS255: with Bracket mode, none of the three adjacent placeholders'
    // contexts should contain bracket characters from sibling placeholders.
    const bytes = bytesFromBase64(FIXTURE_COMBINED_B64);
    const kinds = PlaceholderKinds.All; // 7
    const boundary = ContextBoundary.Bracket;

    const result = await page.evaluate(async ({ bytesArray, kinds, boundary }) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // FindPlaceholders signature: (handle, kinds, scope, contextChars, boundary)
        const placeholders = JSON.parse(
          bridge.FindPlaceholders(handle, kinds, 1, 80, boundary),
        ) as any[];
        return {
          count: placeholders.length,
          contexts: placeholders.map((p) => ({
            text: p.match?.text ?? '',
            contextBefore: p.match?.contextBefore ?? '',
            contextAfter: p.match?.contextAfter ?? '',
          })),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: bytes, kinds, boundary });

    // The combined fixture has more placeholders than the single-paragraph
    // .NET fixture (adjacent + [FILL] + [A]/[B]/[C]). Each context must still
    // be bracket-free under Bracket mode — that's the contract DS255 enforces.
    expect(result.count).toBeGreaterThanOrEqual(3);
    for (const c of result.contexts) {
      expect(c.contextBefore, `contextBefore for ${c.text}`).not.toContain('[');
      expect(c.contextBefore, `contextBefore for ${c.text}`).not.toContain(']');
      expect(c.contextAfter, `contextAfter for ${c.text}`).not.toContain('[');
      expect(c.contextAfter, `contextAfter for ${c.text}`).not.toContain(']');
    }
  });
});
