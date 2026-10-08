/**
 * The worker proxy's request path against a scripted stand-in worker (issue #972): deadlines,
 * per-request cancellation, responses that cannot be deserialized, and success responses that
 * omit their payload. The real engine is never loaded; `Worker` is replaced before the page runs,
 * so each case can make the "worker" hang, answer without a payload, or fail to deliver.
 */

import { test, expect, type Page } from "@playwright/test";

/** Installed before any page script: a Worker whose replies are scripted by `window.fakeMode`. */
function installFakeWorker() {
  type Mode = "ok" | "hang" | "missing" | "messageerror";
  const w = window as any;
  w.fakeMode = "ok" as Mode;
  w.fakeInitMode = "ok" as Mode;
  w.fakeWorkers = [];
  class FakeWorker {
    onmessage: ((event: { data: unknown }) => void) | null = null;
    onerror: ((event: unknown) => void) | null = null;
    onmessageerror: ((event: unknown) => void) | null = null;
    posted: Array<{ id: string; type: string }> = [];
    terminated = false;
    constructor() {
      w.fakeWorkers.push(this);
    }
    postMessage(request: { id: string; type: string }) {
      this.posted.push(request);
      const mode: Mode = request.type === "init" ? w.fakeInitMode : w.fakeMode;
      setTimeout(() => {
        if (this.terminated || mode === "hang") return;
        if (mode === "messageerror") {
          this.onmessageerror?.({ data: null });
          return;
        }
        const payload = mode === "missing" ? {} : {
          html: "<p>ok</p>",
          version: { library: "fake" },
          handle: 7,
        };
        this.onmessage?.({ data: { id: request.id, type: request.type, success: true, ...payload } });
      }, 0);
    }
    terminate() {
      this.terminated = true;
    }
  }
  w.Worker = FakeWorker;
}

async function open(page: Page) {
  await page.addInitScript(installFakeWorker);
  await page.goto("/worker-test-harness.html");
}

/**
 * Settle `script` in the page and report how it ended: the value, the rejection's message and
 * worker error code, or "pending" if it has not settled within `ms`, which is how a request
 * with no deadline shows up.
 */
function outcome(page: Page, script: string, ms = 2000) {
  return page.evaluate(async ({ script, ms }) => {
    const run = new Function(`return (async () => { ${script} })();`)() as Promise<unknown>;
    const pending = new Promise((resolve) => setTimeout(() => resolve({ state: "pending" }), ms));
    return Promise.race([
      run.then(
        (value) => ({ state: "resolved", value }),
        (error: any) => ({ state: "rejected", message: String(error?.message), code: error?.code }),
      ),
      pending,
    ]);
  }, { script, ms });
}

const CREATE = `const { createWorkerDocxodus } = await import("./worker-proxy.js");`;

test.describe("worker proxy request deadlines and payloads (#972)", () => {
  test.beforeEach(async ({ page }) => open(page));

  test("a request that never answers times out instead of staying pending", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/", requestTimeoutMs: 150 });
      window.fakeMode = "hang";
      try { await dx.getVersion(); } finally { window.stillActive = dx.isActive(); }`);
    expect(result).toMatchObject({ state: "rejected", code: "timeout" });
    // A timed-out request does not stop the worker: only terminate() does.
    expect(await page.evaluate(() => (window as any).stillActive)).toBe(true);
  });

  test("an initialization that never answers times out and stops the worker", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      window.fakeInitMode = "hang";
      await createWorkerDocxodus({ wasmBasePath: "/", requestTimeoutMs: 150 });`);
    expect(result).toMatchObject({ state: "rejected", code: "timeout" });
    expect(await page.evaluate(() => (window as any).fakeWorkers[0].terminated)).toBe(true);
  });

  test("a request through an aborted view is rejected, and the worker keeps serving others", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      const controller = new AbortController();
      window.fakeMode = "hang";
      const call = dx.withRequestOptions({ signal: controller.signal }).getVersion();
      setTimeout(() => controller.abort(), 20);
      try { await call; } finally {
        window.fakeMode = "ok";
        window.afterAbort = await dx.getVersion();
      }`);
    expect(result).toMatchObject({ state: "rejected", code: "aborted" });
    expect(await page.evaluate(() => (window as any).afterAbort)).toEqual({ library: "fake" });
  });

  test("an already-aborted signal rejects before anything is posted", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      const controller = new AbortController();
      controller.abort();
      try { await dx.withRequestOptions({ signal: controller.signal }).getVersion(); }
      finally { window.postedTypes = window.fakeWorkers[0].posted.map((r) => r.type); }`);
    expect(result).toMatchObject({ state: "rejected", code: "aborted" });
    expect(await page.evaluate(() => (window as any).postedTypes)).toEqual(["init"]);
  });

  test("a view's timeout overrides the instance default", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      window.fakeMode = "hang";
      await dx.withRequestOptions({ timeoutMs: 100 }).getVersion();`);
    expect(result).toMatchObject({ state: "rejected", code: "timeout" });
  });

  test("a long-lived signal does not accumulate abort listeners", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      const controller = new AbortController();
      let live = 0;
      const add = controller.signal.addEventListener.bind(controller.signal);
      const remove = controller.signal.removeEventListener.bind(controller.signal);
      controller.signal.addEventListener = (...args) => { live++; return add(...args); };
      controller.signal.removeEventListener = (...args) => { live--; return remove(...args); };
      const view = dx.withRequestOptions({ signal: controller.signal });
      for (let i = 0; i < 20; i++) await view.getVersion();
      return live;`);
    expect(result).toEqual({ state: "resolved", value: 0 });
  });

  test("a session opened through a view still closes after the view's signal fired", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      const controller = new AbortController();
      const session = await dx.withRequestOptions({ signal: controller.signal })
        .openDocxSession(new Uint8Array([80, 75]));
      controller.abort();
      await session.close();
      return window.fakeWorkers[0].posted.map((r) => r.type);`);
    expect(result).toEqual({ state: "resolved", value: ["init", "sessionOpen", "sessionClose"] });
  });

  test("a success response without its payload rejects instead of resolving undefined", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      window.fakeMode = "missing";
      return await dx.convertDocxToHtml(new Uint8Array([80, 75]));`);
    expect(result).toMatchObject({ state: "rejected", code: "missing_result" });
  });

  test("a response that cannot be deserialized rejects the request in flight", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      const dx = await createWorkerDocxodus({ wasmBasePath: "/" });
      window.fakeMode = "messageerror";
      try { await dx.getVersion(); } finally {
        window.fakeMode = "ok";
        window.afterMessageError = await dx.getVersion();
      }`);
    expect(result).toMatchObject({ state: "rejected", code: "message_error" });
    expect(await page.evaluate(() => (window as any).afterMessageError)).toEqual({ library: "fake" });
  });

  test("a timeout that is not a positive number is refused up front", async ({ page }) => {
    const result = await outcome(page, `${CREATE}
      await createWorkerDocxodus({ wasmBasePath: "/", requestTimeoutMs: 0 });`);
    expect(result).toMatchObject({ state: "rejected", message: expect.stringContaining("requestTimeoutMs") });
  });
});
