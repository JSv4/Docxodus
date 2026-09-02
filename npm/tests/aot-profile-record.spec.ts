// Records wasm/DocxodusWasm/docxodus.aotprofile — the list of methods the shipped build
// compiles ahead of time (RunAOTCompilation + AOTProfilePath in DocxodusWasm.csproj);
// everything not in it stays on the interpreter/jiterpreter.
//
// Opt-in, and only meaningful against a profiler build: scripts/record-aot-profile.sh
// builds the bundle with <WasmProfilers>aot</WasmProfilers> and AOT off (AOT-compiled
// methods are invisible to the profiler), runs this spec, then rebuilds the shipped
// configuration. The Mono AOT profiler records every method the runtime compiles, so
// the workload — shared with wasm-steady-state.spec.ts — is exactly what ends up AOT'd.
// The profile is written once, when the write-at method (DocumentComparer.Warmup, armed
// by test-harness.html?aotProfile=1) is first compiled, into INTERNAL.aotProfileData.
import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { loadWorkloadInput, runWasmWorkload } from './wasm-workload';

const PROFILE_PATH = path.join(
  path.dirname(fileURLToPath(import.meta.url)), '../../wasm/DocxodusWasm/docxodus.aotprofile');

test.describe('AOT profile recording', () => {
  test.skip(process.env.DOCXODUS_RECORD_AOT_PROFILE !== '1',
    'opt-in: run through scripts/record-aot-profile.sh');

  test('records the AOT profile of the representative workload', async ({ page }) => {
    test.setTimeout(900000);
    // The runtime reports a failed dump on the console, not as an exception
    // (e.g. "Cannot find method in loaded assemblies: 'Interop/Runtime::DumpAotProfileData'"
    // when the trimmer removed the profiler's send-to method).
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(String(e)));
    page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
    await page.goto('/test-harness.html?aotProfile=1');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 60000 });

    // Coverage is what matters here, not repetition: one pass compiles every method.
    const input = loadWorkloadInput({
      warmup: 0,
      iterations: { compareSmall: 1, compareHeavy: 1, convert: 1, convertHeavy: 1, sessionRefresh: 5 },
    });
    const result = await page.evaluate(runWasmWorkload, input);
    for (const t of result.timings) {
      expect(result.outputSizes[t.op], `${t.op} produced output`).toBeGreaterThan(0);
    }

    const profile = await page.evaluate(() => {
      (window as any).Docxodus.DocumentComparer.Warmup();
      const data = (window as any).DocxodusRuntime.INTERNAL.aotProfileData as Uint8Array | undefined;
      return data ? Array.from(data) : null;
    });
    expect(errors).toEqual([]);
    expect(profile, 'INTERNAL.aotProfileData (is this a <WasmProfilers>aot</WasmProfilers> build?)')
      .not.toBeNull();
    expect(profile!.length).toBeGreaterThan(1000);

    fs.writeFileSync(PROFILE_PATH, Buffer.from(profile!));
    console.log(`wrote ${profile!.length} bytes to ${PROFILE_PATH}`);
  });
});
