// Issue #811. Mono's interpreter can mark its stack for the GC "precisely", skipping slots it
// knows hold no object references. To find those slots, every nursery collection walks the
// thread's chain of interpreter-to-native transition records (LMFs, "last managed frame").
// In this mixed AOT + interpreter build that chain can end up cyclic, and WebAssembly does not
// fault on a bad pointer the way a native process would (compare dotnet/runtime#123573), so
// the walk never ends: proveRedlineReversibility and verifyDeliverable were sampled spinning
// at the same loop header, at 100% CPU with no exception, and the hung previewBatch in the
// issue shares their innermost frames.
//
// The walk runs only while the interpreter's precise-GC option is on, which is the .NET 10
// default. The build turns it off through the boot config, and this spec pins that the
// runtime a page actually loads carries the setting. It cannot reproduce the hang itself —
// that depends on heap and binary layout, a few runs in a hundred — so it checks the one
// switch that makes the loop unreachable. See docs/architecture/wasm-packaging.md.
import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('WASM runtime configuration (#811)', () => {
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
});
