// Records wasm/DocxodusWasm/docxodus.aotprofile — the list of methods the shipped build
// compiles ahead of time (RunAOTCompilation + AOTProfilePath in DocxodusWasm.csproj);
// everything not in it stays on the interpreter/jiterpreter.
//
// Opt-in, and only meaningful against a profiler build: scripts/record-aot-profile.sh
// builds the bundle with <WasmProfilers>aot</WasmProfilers> and AOT off (AOT-compiled
// methods are invisible to the profiler), runs this spec, then rebuilds the shipped
// configuration. The Mono AOT profiler records every method the runtime compiles, so
// the workload — the steady-state suite plus a dense-text editor sample below —
// is exactly what ends up AOT'd.
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

    // Dense formatted text is also an ordinary editor workload: code listings,
    // terminal captures and text diagrams can have thousands of short runs.
    // Exercise the generic mutation + formatting-template path so newly added
    // hot methods do not silently remain on the WASM interpreter.
    await page.evaluate(() => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bridge.CreateBlankDocx(), '');
      try {
        const blocks = JSON.parse(bridge.ListBlocks(handle));
        const anchor = blocks.body[0].id;
        const seed = bridge.RawGetXml(handle, anchor) as string;
        const open = seed.slice(0, seed.indexOf('>') + 1).replace(/\/>$/, '>');
        for (let frame = 0; frame < 3; frame++) {
          const runs = Array.from({ length: 1200 }, (_, i) =>
            '<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>' +
            `<w:color w:val="${i % 2 ? 'FF8844' : 'FFFFFF'}"/><w:spacing w:val="-2"/>` +
            '<w:sz w:val="6"/></w:rPr><w:t xml:space="preserve">' +
            `| ${frame}: terminal text ${i}   </w:t>${i % 10 === 0 ? '<w:br/>' : ''}</w:r>`).join('');
          const changed = JSON.parse(bridge.RawReplaceXml(handle, anchor, open +
            '<w:pPr><w:spacing w:line="48" w:lineRule="exact"/></w:pPr>' + runs + '</w:p>'));
          if (!changed.success) throw new Error(JSON.stringify(changed));
          bridge.RenderBlockHtml(handle, anchor, 'profile-', false);
        }
      } finally { bridge.CloseSession(handle); }
    });

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
