// Steady-state speed of the shipped WASM bundle, and the runtime-configuration canary
// behind it (issue #652).
//
// The browser build runs the engine on the Mono interpreter, tiered by the jiterpreter
// (hot interpreter traces compiled to WebAssembly at runtime). Two things can silently
// take that speed away: a build knob or feature switch that disables the jiterpreter,
// and the jiterpreter's own budget of generated code (`jiterpreter-wasm-bytes-limit`),
// past which no new trace is compiled. The first test pins both.
//
// The second test is a MEASUREMENT harness, not a fidelity pin: it runs the
// representative workload in wasm-workload.ts (warm, N iterations) and prints
// median/p95 per operation so every runtime configuration — interpreter, jiterpreter
// knobs, profiled or full AOT — is scored the same way. Assertions are limited to
// "the operations succeeded"; wall-clock numbers vary too much across machines to gate
// CI on. Set DOCXODUS_BENCH_OUT=<dir> to also write steady-state.json and the generated
// heavy-variant .docx there, for an out-of-browser counterpart run on the same inputs.
import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { loadWorkloadInput, runWasmWorkload, WorkloadResult } from './wasm-workload';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

const fmt = (ms: number) => (ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${ms.toFixed(1)} ms`);

test.describe('WASM steady-state (issue #652)', () => {
  test('the jiterpreter is active and has headroom under its code budget', async ({ page }) => {
    test.setTimeout(300000);
    const logs: string[] = [];
    page.on('console', (m) => logs.push(m.text()));
    // Stats printing is a boot-time runtime option (see test-harness.html); the timing test
    // below boots the plain configuration.
    await page.goto('/test-harness.html?jiterpStats=1');
    await waitForDocxodus(page);

    // Enough work for the tiering heuristics to compile traces.
    const input = loadWorkloadInput({
      warmup: 0,
      iterations: { compareSmall: 3, compareHeavy: 0, convert: 1, convertHeavy: 0, sessionRefresh: 0 },
    });
    await page.evaluate(runWasmWorkload, input);

    const status = await page.evaluate(() => {
      const I = (window as any).DocxodusRuntime.INTERNAL;
      const o = I.jiterpreter_get_options();
      I.jiterpreter_dump_stats(/* concise */ true);
      return {
        enableTraces: o.enableTraces,
        enableInterpEntry: o.enableInterpEntry,
        enableJitCall: o.enableJitCall,
        wasmBytesLimit: o.wasmBytesLimit,
      };
    });

    // The runtime's concise stats line:
    //   // jitted <bytes>b; <traces> traces (<pct>%) (<n> rejected); <n> jit_calls; <n> interp_entries
    const line = logs.find((l) => /jitted \d+b; \d+ traces/.test(l));
    expect(line, 'jiterpreter stats line (runtime log format changed?)').toBeDefined();
    const m = /jitted (\d+)b; (\d+) traces .*?(\d+) jit_calls; (\d+) interp_entries/.exec(line!)!;
    const stats = { jittedBytes: +m[1], traces: +m[2], jitCalls: +m[3], interpEntries: +m[4] };
    console.log(`jiterpreter: ${JSON.stringify({ ...status, ...stats })}`);

    expect(status.enableTraces).toBeTruthy();
    expect(status.enableInterpEntry).toBeTruthy();
    expect(status.enableJitCall).toBeTruthy();
    expect(stats.traces).toBeGreaterThan(0);
    // Once jitted bytes reach the limit the jiterpreter stops compiling; a workload this
    // small must sit well inside it.
    expect(stats.jittedBytes).toBeLessThan(status.wasmBytesLimit / 2);
  });

  // The AOT bundle prints a handful of "Jiterpreter table N is not yet initialized" lines
  // while booting: the runtime allocates the 38 function-table slices in
  // jiterpreter_allocate_tables(), which start_runtime() calls on the line AFTER
  // mono_wasm_load_runtime() returns, and runtime startup itself crosses the AOT->interp
  // boundary a few times before that. Each such crossing builds its entry thunk eagerly, finds
  // the table still zeroed, and falls back to the runtime's generic entry wrapper. Cosmetic —
  // see docs/architecture/wasm-packaging.md.
  //
  // A jiterpreter that is genuinely broken logs the SAME message, so this pins the difference
  // rather than the noise: the allocation must have completed, the fallbacks must be confined
  // to boot, and they must never name the trace (0) or do_jit_call (1) tables — those two carry
  // the whole tier, and losing them would drop the engine back to the plain interpreter with
  // nothing but console output to say so.
  test('the jiterpreter function tables are allocated, and only boot-time interp_entry thunks fall back', async ({ page }) => {
    test.setTimeout(300000);
    const logs: { afterBoot: boolean; text: string }[] = [];
    let afterBoot = false;
    page.on('console', (m) => logs.push({ afterBoot, text: m.text() }));

    await page.goto('/test-harness.html?jiterpStats=1');
    await waitForDocxodus(page);
    afterBoot = true;

    // jiterpreter_allocate_tables() reserves traceTableSize + jitCallTableSize +
    // 36 * interpEntryTableSize + 1 entries and logs the total under --jiterpreter-stats-enabled.
    // Its absence means the call threw and NO table was initialized.
    const allocated = logs
      .map((l) => /Allocated (\d+) function table entries for jiterpreter/.exec(l.text))
      .find((m) => m !== null);
    expect(allocated, 'jiterpreter function-table allocation line (did allocate_tables throw?)')
      .toBeTruthy();
    // 38 slices at minimum one entry each, plus the spare. Anything smaller means the loop
    // over the interp_entry tables did not run to completion.
    expect(Number(allocated![1])).toBeGreaterThan(38);

    // Enough work to drive the tiering heuristics well past startup.
    const input = loadWorkloadInput({
      warmup: 0,
      iterations: { compareSmall: 3, compareHeavy: 0, convert: 2, convertHeavy: 0, sessionRefresh: 20 },
    });
    await page.evaluate(runWasmWorkload, input);

    const uninitialized = logs
      .map((l) => ({ l, m: /Jiterpreter table (\d+) is not yet initialized/.exec(l.text) }))
      .filter((x) => x.m !== null)
      .map((x) => ({ afterBoot: x.l.afterBoot, table: Number(x.m![1]) }));
    console.log(`interp_entry fallbacks: ${JSON.stringify(uninitialized)}`);

    // Tables 0 (traces) and 1 (do_jit_call) are allocated first and must never be reported.
    expect(uninitialized.filter((u) => u.table <= 1)).toEqual([]);
    // Every fallback belongs to startup; once the tables exist, allocation always succeeds.
    expect(uninitialized.filter((u) => u.afterBoot)).toEqual([]);
    // 36 interp_entry tables; a handful of boot crossings is expected (9 on the 2026-09 AOT
    // bundle), a flood is not.
    expect(uninitialized.length).toBeLessThanOrEqual(36);
  });

  test('steady-state latency of the representative workload', async ({ page }) => {
    test.setTimeout(900000);
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
    const outDir = process.env.DOCXODUS_BENCH_OUT;
    const input = loadWorkloadInput({
      warmup: 2,
      iterations: { compareSmall: 15, compareHeavy: 5, convert: 15, convertHeavy: 5, sessionRefresh: 30 },
      returnVariant: !!outDir,
    });
    const result: WorkloadResult = await page.evaluate(runWasmWorkload, input);

    console.log(`heavy variant edits: ${JSON.stringify(result.variantEdits)}`);
    console.log('op                    n   median      p95      min');
    for (const t of result.timings) {
      console.log(
        `${t.op.padEnd(20)} ${String(t.n).padStart(3)} ${fmt(t.medianMs).padStart(9)} ` +
        `${fmt(t.p95Ms).padStart(9)} ${fmt(t.minMs).padStart(9)}`);
    }

    expect(result.variantEdits.replaced + result.variantEdits.inserted + result.variantEdits.deleted)
      .toBeGreaterThan(0);
    for (const t of result.timings) {
      expect(result.outputSizes[t.op], `${t.op} produced output`).toBeGreaterThan(0);
    }

    if (outDir) {
      fs.mkdirSync(outDir, { recursive: true });
      const { variant, ...rest } = result;
      fs.writeFileSync(path.join(outDir, 'steady-state.json'), JSON.stringify(rest, null, 2));
      fs.writeFileSync(path.join(outDir, 'heavy-variant.docx'), Buffer.from(variant!));
    }
  });
});
