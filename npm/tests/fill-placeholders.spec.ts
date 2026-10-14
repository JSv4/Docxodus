import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Helper to decode the inline base64 fixtures into a number[] ready for
// `page.evaluate`. The fixtures are tiny (~1.2 KB) purpose-built DOCXs built
// by the .NET BuildDocWithBracketPlaceholders / BuildDocWithLongClauseBlank
// helpers used by the DocxSession.FillPlaceholders unit tests (Issue #163,
// units A/B/C); inlining them keeps this spec self-contained instead of
// requiring committed binary fixtures.
function bytesFromBase64(b64: string): number[] {
  const bin = Buffer.from(b64, 'base64');
  const out = new Array<number>(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin[i]!;
  return out;
}

// Five-paragraph fixture covering every placeholder shape `findPlaceholders`
// recognizes: a plain `[_____]`, a `$[___]` (dollar-prefix), a third blank,
// a single bracketed alternative clause, and a nested `[outer [inner] clause]`.
const FIXTURE_BRACKET_B64 =
  'UEsDBBQAAAAIAAuPuVwARQTL/gAAAO4BAAARAAAAd29yZC9kb2N1bWVudC54bWyNUcFOwzAM/RWr4gjN4IBQtXU/MHHiVk3IJN4aqYkjx6X070nQEJdJzIcXPcXvOX7Z7r/CBJ8k2XPcNY/tpgGKlp2P510z6+nhpdn326VzbOdAUaH0x9wtu2ZUTZ0x2Y4UMLecKJa7E0tALVTOZmFxSdhSzsUuTOZps3k2AX1squUHu7WeqYJU0P5tJIgYCPgEOvoMliWxoJbnQaHDe61juzW1u6L8YLpmlMTbgiSQRxSq+rvhNvmBLSo5QL2MvNQN0uGVdfE6ZsVYYyx7EJRY6MyF3QM65+s6OIGShAwBV8CUprU9/m/OcxHB4GMkOYKdcM50TWZ+8zV/f9d/A1BLAwQUAAAAAAALj7lct/en5yoBAAAqAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3dvcmQvZG9jdW1lbnQueG1sIiBJZD0iUmNlMDAyOTVlNTM1NzQ2YzciIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAAAAALj7lc3xsCV1oBAABaAQAAEwAAAFtDb250ZW50X1R5cGVzXS54bWzvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxUeXBlcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9jb250ZW50LXR5cGVzIj48RGVmYXVsdCBFeHRlbnNpb249InhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC53b3JkcHJvY2Vzc2luZ21sLmRvY3VtZW50Lm1haW4reG1sIiAvPjxEZWZhdWx0IEV4dGVuc2lvbj0icmVscyIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1wYWNrYWdlLnJlbGF0aW9uc2hpcHMreG1sIiAvPjwvVHlwZXM+UEsBAhQDFAAAAAgAC4+5XABFBMv+AAAA7gEAABEAAAAAAAAAAAAAAKSBAAAAAHdvcmQvZG9jdW1lbnQueG1sUEsBAhQDFAAAAAAAC4+5XLf3p+cqAQAAKgEAAAsAAAAAAAAAAAAAAKSBLQEAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAAAC4+5XN8bAldaAQAAWgEAABMAAAAAAAAAAAAAAKSBgAIAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAMAAwC5AAAACwQAAAAA';

// Single-paragraph fixture for the alternativeKinds borderline case: a long
// bracketed clause whose interior happens to contain a `$_______` blank.
// findPlaceholders' primary classification stays `blank_fill` for back-compat,
// but `alternativeKinds` must additionally include `alternative_clause`.
const FIXTURE_LONG_CLAUSE_B64 =
  'UEsDBBQAAAAIAAuPuVxQw8aRyAAAACcBAAARAAAAd29yZC9kb2N1bWVudC54bWxFj9FuwyAMRX/FQnvcSraHaYqS9EOqacrATZAAIxuS9u8Hqar54Rhfri5mON+Chw1ZHMVRvZ86BRgNWReXUZV8fftS52nYe0umBIwZqj9Kv49qzTn1WotZMcxyooSx3l2Jw5zryIveiW1iMihS44LXH133qcPsomqRv2TvracGbsjTJa9zhp2Kt8AoxWdwEarkcZYMLz+PauLCJAJHPlp5rZLxpe0NxIC35/A01ANujor4e0026Da034NujzbywXTwsZj+//T0B1BLAwQUAAAAAAALj7lcwruJQyoBAAAqAQAACwAAAF9yZWxzLy5yZWxz77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48UmVsYXRpb25zaGlwcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9yZWxhdGlvbnNoaXBzIj48UmVsYXRpb25zaGlwIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvb2ZmaWNlRG9jdW1lbnQiIFRhcmdldD0iL3dvcmQvZG9jdW1lbnQueG1sIiBJZD0iUmE0MjkyZDY5ZjE2MzQ5NmIiIC8+PC9SZWxhdGlvbnNoaXBzPlBLAwQUAAAAAAALj7lc3xsCV1oBAABaAQAAEwAAAFtDb250ZW50X1R5cGVzXS54bWzvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxUeXBlcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9jb250ZW50LXR5cGVzIj48RGVmYXVsdCBFeHRlbnNpb249InhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC53b3JkcHJvY2Vzc2luZ21sLmRvY3VtZW50Lm1haW4reG1sIiAvPjxEZWZhdWx0IEV4dGVuc2lvbj0icmVscyIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1wYWNrYWdlLnJlbGF0aW9uc2hpcHMreG1sIiAvPjwvVHlwZXM+UEsBAhQDFAAAAAgAC4+5XFDDxpHIAAAAJwEAABEAAAAAAAAAAAAAAKSBAAAAAHdvcmQvZG9jdW1lbnQueG1sUEsBAhQDFAAAAAAAC4+5XMK7iUMqAQAAKgEAAAsAAAAAAAAAAAAAAKSB9wAAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAAAC4+5XN8bAldaAQAAWgEAABMAAAAAAAAAAAAAAKSBSgIAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAMAAwC5AAAA1QMAAAAA';

test.describe('fill-placeholders (WASM bridge — Issue #163)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // The harness exposes only `window.Docxodus.DocxSessionBridge` (raw WASM
  // exports), matching the convention of every other session-spec in this
  // directory. `session.fillPlaceholders` is a TS-side multi-pass loop
  // (see npm/src/session.ts) over `findPlaceholders` + `replaceMatch`; we
  // reproduce it inline in `page.evaluate` so the same JS algorithm runs end-
  // to-end against the WASM bridge — every WASM primitive the wrapper uses
  // (FindPlaceholders, ReplaceTextAtSpan, ReplaceInner) is exercised.

  test('picker fills a single BlankFill', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // PlaceholderKinds.BlankFill = 1, ProjectionScopes.Body = 1
        // Mirror npm/src/session.ts fillPlaceholders for the simple case.
        let filled = 0;
        let passes = 0;
        for (let pass = 1; pass <= 8; pass++) {
          const placeholders = (JSON.parse(bridge.FindPlaceholders(handle, 1, 1, 80, 0)) as any[])
            .sort((a, b) => {
              const c = b.match.enclosingAnchor.id.localeCompare(a.match.enclosingAnchor.id);
              if (c !== 0) return c;
              return b.match.span.start - a.match.span.start;
            });
          if (placeholders.length === 0) break;
          let changes = 0;
          for (const p of placeholders) {
            const pick = p.kind === 'blank_fill' && p.match.text === '[_____]' ? 'ACME, INC.' : null;
            if (pick == null) continue;
            const r = JSON.parse(bridge.ReplaceTextAtSpan(
              handle, p.match.enclosingAnchor.id, p.match.span.start, p.match.span.length, pick));
            if (r.success) { filled++; changes++; }
          }
          if (changes > 0) passes = pass;
          if (changes === 0) break;
        }
        const proj = JSON.parse(bridge.Project(handle));
        return { filled, passes, markdown: proj.markdown };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.filled).toBe(1);
    expect(result.markdown).toContain('ACME, INC.');
    expect(result.markdown).not.toContain('[_____]');
  });

  test('picker returning null skips every placeholder', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // Test narrows to BlankFill | Instruction = 5 explicitly (the TS wrapper's
        // default is now PlaceholderKinds.All = 7; we keep 5 here so this test
        // measures the same surface area it always did).
        const placeholders = JSON.parse(bridge.FindPlaceholders(handle, 5, 1, 80, 0)) as any[];
        const unfilled: any[] = [];
        let filled = 0;
        for (const p of placeholders) {
          const pick = null;
          if (pick == null) { unfilled.push(p); continue; }
          filled++;
        }
        return { filled, unfilledCount: unfilled.length };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.filled).toBe(0);
    expect(result.unfilledCount).toBeGreaterThan(0);
  });

  test('preserves $ prefix when picker returns a plain number', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // BlankFill kind only.
        const placeholders = JSON.parse(bridge.FindPlaceholders(handle, 1, 1, 80, 0)) as any[];
        let filled = 0;
        for (const p of placeholders) {
          if (!p.match.text.includes('$[___]')) continue;
          // Mirror the preserveDollarPrefix=true branch of session.fillPlaceholders:
          // picker returns "0.20", the match starts with "$" so we prepend "$".
          let replacement = '0.20';
          if (p.match.text.startsWith('$') && !replacement.startsWith('$')) {
            replacement = '$' + replacement;
          }
          const r = JSON.parse(bridge.ReplaceTextAtSpan(
            handle, p.match.enclosingAnchor.id, p.match.span.start, p.match.span.length, replacement));
          if (r.success) filled++;
        }
        const proj = JSON.parse(bridge.Project(handle));
        return { filled, markdown: proj.markdown };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.filled).toBe(1);
    expect(result.markdown).toContain('$0.20');
    expect(result.markdown).not.toContain('$$0.20');
  });

  test('multi-pass strips nested brackets', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // AlternativeClause = 2.
        const kinds = 2;
        const scope = 1;
        let passes = 0;
        let filled = 0;
        for (let pass = 1; pass <= 8; pass++) {
          const placeholders = (JSON.parse(bridge.FindPlaceholders(handle, kinds, scope, 80, 0)) as any[])
            .sort((a, b) => {
              const c = b.match.enclosingAnchor.id.localeCompare(a.match.enclosingAnchor.id);
              if (c !== 0) return c;
              return b.match.span.start - a.match.span.start;
            });
          if (placeholders.length === 0) break;
          let changes = 0;
          for (const p of placeholders) {
            // Picker strips the outer brackets of the match, keeping any
            // prefix/suffix outside (same algorithm as DS244).
            const t: string = p.match.text;
            const lb = t.indexOf('[');
            const rb = t.lastIndexOf(']');
            if (lb < 0 || rb <= lb) continue;
            const pick = t.slice(0, lb) + t.slice(lb + 1, rb) + t.slice(rb + 1);
            const r = JSON.parse(bridge.ReplaceTextAtSpan(
              handle, p.match.enclosingAnchor.id, p.match.span.start, p.match.span.length, pick));
            if (r.success) { filled++; changes++; }
          }
          if (changes > 0) passes = pass;
          if (changes === 0) break;
        }
        const leftover = JSON.parse(bridge.FindPlaceholders(handle, kinds, scope, 80, 0)) as any[];
        return { passes, filled, leftoverCount: leftover.length };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.passes).toBeGreaterThanOrEqual(2);
    expect(result.leftoverCount).toBe(0);
  });

  test('replaceInner directly substitutes inside brackets and preserves $ prefix', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_BRACKET_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // BlankFill kind, body scope.
        const placeholders = JSON.parse(bridge.FindPlaceholders(handle, 1, 1, 80, 0)) as any[];
        const target = placeholders.find((p: any) => p.match.text.startsWith('$['));
        if (!target) return { error: 'no $[...] placeholder found' };

        const r = JSON.parse(bridge.ReplaceInner(
          handle,
          target.match.text,
          target.match.enclosingAnchor.id,
          target.match.span.start,
          target.match.span.length,
          '0.20',
        ));
        const proj = JSON.parse(bridge.Project(handle));
        return { success: r.success, errorCode: r.error?.code, markdown: proj.markdown };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.error).toBeUndefined();
    expect(result.success).toBe(true);
    expect(result.markdown).toContain('$0.20');
    expect(result.markdown).not.toContain('$[___]');
  });

  test('alternativeKinds includes alternative_clause for long-clause-with-blanks', async ({ page }) => {
    const bytes = bytesFromBase64(FIXTURE_LONG_CLAUSE_B64);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // PlaceholderKinds.All = 7.
        const placeholders = JSON.parse(bridge.FindPlaceholders(handle, 7, 1, 80, 0)) as any[];
        return {
          count: placeholders.length,
          first: placeholders[0] ?? null,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, bytes);

    expect(result.count).toBe(1);
    expect(result.first).not.toBeNull();
    expect(result.first.kind).toBe('blank_fill');
    expect(Array.isArray(result.first.alternativeKinds)).toBe(true);
    expect(result.first.alternativeKinds).toContain('alternative_clause');
  });
});
