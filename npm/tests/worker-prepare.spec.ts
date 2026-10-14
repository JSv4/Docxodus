/**
 * Tests for WorkerDocxodus.prepare() — pre-warming the comparison code path.
 *
 * prepare() runs a real comparison against in-memory seed documents inside the
 * worker, forcing the .NET WASM runtime to fully resolve and JIT the comparison
 * code path. After it resolves, the first real compareDocuments() (a) triggers
 * no further .wasm fetches and (b) runs meaningfully faster than a cold one.
 *
 * We monitor network requests for ".wasm" at the page level (Playwright
 * surfaces dedicated-worker requests on the owning page) and time compares
 * inside the worker via performance.now().
 */

import { test, expect, type Request } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const testFilesDir = path.join(__dirname, "../../TestFiles");

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(testFilesDir, relativePath)));
}

/** Last path segment of a .wasm URL, e.g. ".../Docxodus.wasm" -> "Docxodus.wasm". */
function wasmName(url: string): string {
  return url.split("/").pop() ?? url;
}

test.describe("WorkerDocxodus.prepare()", () => {
  let wasmRequests: string[];

  test.beforeEach(async ({ page }) => {
    wasmRequests = [];
    page.on("request", (req: Request) => {
      const url = req.url();
      if (url.endsWith(".wasm")) {
        wasmRequests.push(url);
      }
    });

    await page.goto("/worker-test-harness.html");
    await page.waitForFunction(
      () => (window as any).DocxodusWorkerTests !== undefined,
      { timeout: 10000 }
    );
  });

  test("after prepare(), a real compare triggers no additional .wasm fetches", async ({
    page,
  }) => {
    // Bring up the worker and warm the comparison path.
    await page.evaluate(() => (window as any).createDocxodusWorker());
    const prepareResult = await page.evaluate(() =>
      (window as any).DocxodusWorkerTests.prepare()
    );
    // Let any in-flight 'request' events drain before snapshotting.
    await page.waitForTimeout(300);

    expect(prepareResult.error).toBeUndefined();
    expect(prepareResult.ok).toBe(true);
    // The monitor must actually observe the worker's .wasm requests, otherwise
    // the "no new fetches" assertion below would be a meaningless 0 === 0.
    expect(wasmRequests.length).toBeGreaterThan(0);

    const beforeCompare = [...wasmRequests];

    // A real comparison after prepare() must trigger NO new .wasm fetches —
    // everything the engine needs is already loaded.
    const originalBytes = readTestFile("WC/WC001-Digits.docx");
    const modifiedBytes = readTestFile("WC/WC001-Digits-Mod.docx");

    const compareResult = await page.evaluate(
      async ([original, modified]) =>
        (window as any).DocxodusWorkerTests.compareDocuments(original, modified),
      [Array.from(originalBytes), Array.from(modifiedBytes)]
    );
    await page.waitForTimeout(300);

    expect(compareResult.error).toBeUndefined();
    expect(compareResult.docxBytes.length).toBeGreaterThan(0);

    const loadedByCompare = wasmRequests
      .filter((u) => !beforeCompare.includes(u))
      .map(wasmName);

    console.log(
      `Assemblies loaded by compareDocuments() after prepare(): ${
        loadedByCompare.length === 0 ? "(none)" : loadedByCompare.join(", ")
      }`
    );

    expect(loadedByCompare).toEqual([]);
  });

  test("prepare() makes the first real compare meaningfully faster", async ({
    page,
  }) => {
    const originalBytes = readTestFile("WC/WC001-Digits.docx");
    const modifiedBytes = readTestFile("WC/WC001-Digits-Mod.docx");
    const original = Array.from(originalBytes);
    const modified = Array.from(modifiedBytes);

    // Throwaway worker to absorb one-time page-level costs (module eval, first
    // dynamic import of the .NET runtime), so the cold/warm gap reflects
    // prepare() rather than page warmup.
    await page.evaluate(async () => {
      await (window as any).createDocxodusWorker();
      (window as any).DocxodusWorker.terminate();
    });

    // COLD: fresh worker, no prepare(), time the first compare.
    const coldFirstMs = await page.evaluate(
      async ([o, m]) => {
        await (window as any).createDocxodusWorker();
        const t = performance.now();
        await (window as any).DocxodusWorker.compareDocuments(
          new Uint8Array(o),
          new Uint8Array(m)
        );
        const elapsed = performance.now() - t;
        (window as any).DocxodusWorker.terminate();
        return elapsed;
      },
      [original, modified]
    );

    // WARM: fresh worker, prepare() first, then time the first compare.
    const warm = await page.evaluate(
      async ([o, m]) => {
        await (window as any).createDocxodusWorker();
        const tp = performance.now();
        await (window as any).DocxodusWorker.prepare();
        const prepMs = performance.now() - tp;
        const t = performance.now();
        await (window as any).DocxodusWorker.compareDocuments(
          new Uint8Array(o),
          new Uint8Array(m)
        );
        const firstMs = performance.now() - t;
        return { prepMs, firstMs };
      },
      [original, modified]
    );

    console.log(
      `cold first compare=${coldFirstMs.toFixed(0)}ms | ` +
        `prepare=${warm.prepMs.toFixed(0)}ms, warm first compare=${warm.firstMs.toFixed(0)}ms`
    );

    // prepare() must do real warmup work (not a no-op).
    expect(warm.prepMs).toBeGreaterThan(50);
    // The whole point: a warmed first compare is substantially cheaper than a
    // cold one. Observed ~1.9x; assert a conservative 15% improvement.
    expect(warm.firstMs).toBeLessThan(coldFirstMs * 0.85);
  });

  test("prepare() is idempotent — second call resolves in <50ms", async ({
    page,
  }) => {
    await page.evaluate(() => (window as any).createDocxodusWorker());

    const first = await page.evaluate(() =>
      (window as any).DocxodusWorkerTests.prepare()
    );
    expect(first.error).toBeUndefined();
    expect(first.ok).toBe(true);

    const second = await page.evaluate(() =>
      (window as any).DocxodusWorkerTests.prepare()
    );
    expect(second.error).toBeUndefined();
    expect(second.ok).toBe(true);

    console.log(
      `prepare() first=${first.durationMs.toFixed(1)}ms, second=${second.durationMs.toFixed(1)}ms`
    );

    expect(second.durationMs).toBeLessThan(50);

    // Idempotent second call must not re-fetch any assembly.
    await page.waitForTimeout(200);
    const unique = new Set(wasmRequests);
    expect(unique.size).toBe(wasmRequests.length);
  });

  test("compareDocuments() while prepare() is in flight does not double-load assemblies", async ({
    page,
  }) => {
    await page.evaluate(() => (window as any).createDocxodusWorker());
    await page.waitForTimeout(200);

    const originalBytes = readTestFile("WC/WC001-Digits.docx");
    const modifiedBytes = readTestFile("WC/WC001-Digits-Mod.docx");

    // Fire prepare() and compareDocuments() concurrently — do NOT await prepare
    // before issuing the compare.
    const result = await page.evaluate(
      async ([original, modified]) => {
        const worker = (window as any).DocxodusWorker;
        const preparePromise = worker.prepare();
        const comparePromise = worker.compareDocuments(
          new Uint8Array(original),
          new Uint8Array(modified)
        );
        const [, docxBytes] = await Promise.all([preparePromise, comparePromise]);
        return { docxLength: docxBytes.length };
      },
      [Array.from(originalBytes), Array.from(modifiedBytes)]
    );

    await page.waitForTimeout(300);

    expect(result.docxLength).toBeGreaterThan(0);

    // No .wasm URL may be fetched more than once — the runtime de-duplicates
    // concurrent loads, so there is no double-load even under contention.
    const duplicates = wasmRequests.filter(
      (u, i) => wasmRequests.indexOf(u) !== i
    );
    console.log(
      `Total .wasm requests: ${wasmRequests.length}, unique: ${new Set(wasmRequests).size}`
    );
    expect(duplicates.map(wasmName)).toEqual([]);
  });
});
