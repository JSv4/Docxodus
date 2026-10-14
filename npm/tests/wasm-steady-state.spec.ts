// Steady-state speed of the shipped WASM bundle, and the runtime-configuration canary
// behind it (issue #652).
//
// The browser build runs the engine on the Mono interpreter, tiered by the jiterpreter
// (hot interpreter traces compiled to WebAssembly at runtime). Three things can silently
// take that speed away: a build knob or feature switch that disables the jiterpreter,
// the jiterpreter's own budget of generated code (`jiterpreter-wasm-bytes-limit`), past
// which no new trace is compiled, and a comparison entry point that runs the engine's
// cold path on a real document instead of warming first. The first two tests pin those.
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
import { loadWorkloadInput, runWasmWorkload, WorkloadResult, WORKLOAD_TEST_FILES_DIR } from './wasm-workload';

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

  // Issues #695, #696. A comparison that is the FIRST one a module instance runs executes the
  // engine's cold path — much of it on the interpreter, which Mono's GC pins conservatively
  // from — while allocating a package's worth of XML. On a heap an earlier operation has
  // already filled, the two together stop the comparison making meaningful progress: this exact
  // sequence ran for over ten minutes before ComparisonEngine.EnsureWarm() became an invariant
  // of every comparison export, against ~25 ms natively and ~500 ms warm in the browser.
  //
  // The order is the whole test. A session read must come first and the comparison must be the
  // first of the page: warming the engine by comparing first makes it pass trivially.
  test('a comparison that follows a session read is not the engine cold path', async ({ page }) => {
    test.setTimeout(180000);
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);

    const read = (rel: string) => Array.from(fs.readFileSync(path.join(WORKLOAD_TEST_FILES_DIR, rel)));
    const revisionDoc = read('FA/RevTracking/014-MovedParagraph.docx');
    const left = read('WC/WC007-Unmodified.docx');
    const right = read('WC/WC007-Moved-into-Table.docx');

    const revisions = await page.evaluate(
      (bytes) => (window as any).DocxodusTests.getRevisions(new Uint8Array(bytes)),
      revisionDoc);
    expect(revisions.error, 'session revision read').toBeUndefined();
    expect(revisions.revisions.length, 'the read has to do real work to fill the heap')
      .toBeGreaterThan(0);

    const t0 = Date.now();
    const compared = await page.evaluate(
      ([l, r]) => {
        const result = (window as any).DocxodusTests.compareDocuments(new Uint8Array(l), new Uint8Array(r));
        return { len: result.docxBytes ? result.docxBytes.length : -1, error: result.error };
      },
      [left, right]);
    const elapsed = Date.now() - t0;

    expect(compared.error).toBeUndefined();
    expect(compared.len, 'the comparison produced a package').toBeGreaterThan(0);
    // Warm in this browser is ~0.5 s including the warm-up; the collapse this pins is three
    // orders of magnitude away, so the budget is deliberately loose enough to survive a slow
    // CI machine and still fail the moment the cold path is back.
    console.log(`first comparison after a session read: ${fmt(elapsed)}`);
    expect(elapsed, 'first comparison after a session read').toBeLessThan(30000);
  });
});
