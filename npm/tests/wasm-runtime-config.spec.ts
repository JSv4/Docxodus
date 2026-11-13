// Issue #811. Mono's interpreter can mark its stack for the GC "precisely", skipping slots it
// knows hold no object references. To find those slots, every nursery collection walks the
// thread's chain of interpreter-to-native transition records (LMFs, "last managed frame").
// Mono's AOT compiler drops the record pop on a `ret` reached only by a backward branch, so the
// OpenXml SDK's GetAllParts() iterator leaves its record behind on every `yield return` and the
// next MoveNext links it to itself. WebAssembly does not fault on that walk the way a native
// process would, so a collection inside such a loop never returns: previewBatch,
// proveRedlineReversibility and verifyDeliverable were all seen spinning at 100% CPU.
//
// The walk runs only while the interpreter's precise-GC option is on, which is the .NET 10
// default. The build turns it off through the boot config, and this spec pins that the
// runtime a page actually loads carries the setting. It cannot reproduce the hang itself —
// that needs a collection to land inside one of those loops, a few runs in a hundred — so it
// checks the one switch that makes the loop unreachable. See docs/architecture/wasm-packaging.md.
import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('WASM runtime configuration (#811, #1704)', () => {
  test('the loaded runtime runs the interpreter without precise stack marking', async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);

    const env = await page.evaluate(
      () => (globalThis as any).getDotnetRuntime(0).getConfig().environmentVariables ?? {});

    // Mono reads MONO_INTERPRETER_OPTIONS once, when the interpreter initializes, as a
    // comma-separated list in which a leading '-' clears an option.
    const options = String(env.MONO_INTERPRETER_OPTIONS ?? '').split(',');
    expect(options, 'MONO_INTERPRETER_OPTIONS in the loaded boot config').toContain('-precise');
  });

  // Issue #1704. Once an interpreted method has been called often enough from compiled (AOT)
  // code, the jiterpreter swaps the runtime's built-in entry path for a generated "interp-entry"
  // thunk. With those thunks on, a nursery collection landing while AOT and interpreted frames
  // alternate over freshly allocated objects left a stale reference behind, and compareDocuments
  // threw InvalidCastException about once in a hundred calls after warm-up.
  // wasm-gc-stress.spec.ts provokes that failure; this checks the switch that closes it, on
  // the runtime a page actually loads. The second boot adds a runtime option of its own, as an
  // embedder may, to show the build's setting survives that.
  for (const query of ['', '?jiterpStats=1']) {
    test(`interp-entry thunks are off, traces and jit-call thunks on (${query || 'plain boot'})`, async ({ page }) => {
      await page.goto('/test-harness.html' + query);
      await waitForDocxodus(page);

      const jiterpreter = await page.evaluate(() => {
        const o = (window as any).DocxodusRuntime.INTERNAL.jiterpreter_get_options();
        return { traces: !!o.enableTraces, interpEntry: !!o.enableInterpEntry, jitCall: !!o.enableJitCall };
      });

      expect(jiterpreter).toEqual({ traces: true, interpEntry: false, jitCall: true });
    });
  }
});
