import assert from "node:assert/strict";
import { stat } from "node:fs/promises";
import { setImmediate } from "node:timers/promises";
import { test } from "node:test";
import { chromium } from "playwright-core";
import { openOwnedExportBrowserSession } from "../dist/browser-session.js";
import { renderDocxArtifacts } from "../dist/index.js";
import { generateBrowserCleanupDocx } from "./mixed-section-fixture.mjs";

async function fakeSession(t, close) {
  let directory;
  t.mock.method(chromium, "launch", async (options) => {
    directory = options.env.HOME;
    return { close };
  });
  const session = await openOwnedExportBrowserSession(process.execPath, 120_000);
  return { session, directory };
}

test("owned shutdown waits for Playwright's force-close window before removing its runtime", async (t) => {
  const { session, directory } = await fakeSession(t,
    () => new Promise((resolve) => setTimeout(resolve, 31_000)));
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const closing = session.close().then(() => undefined, (error) => error);
  t.mock.timers.tick(30_000);
  await setImmediate();
  assert.equal((await stat(directory)).isDirectory(), true);
  t.mock.timers.tick(1_000);
  assert.equal(await closing, undefined);
  await assert.rejects(stat(directory), { code: "ENOENT" });
});

test("owned shutdown remains bounded if Playwright never finishes", async (t) => {
  const { session } = await fakeSession(t, () => new Promise(() => {}));
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const closing = session.close().then(() => undefined, (error) => error);
  t.mock.timers.tick(40_000);
  const error = await closing;
  assert.equal(error.code, "filesystem_failure");
  assert.equal(error.phase, "cleanup");
  assert.match(error.cause.errors[0].message, /40000ms cleanup limit/);
});

test("owned shutdown still reports a browser close error", async (t) => {
  const cause = new Error("Synthetic browser transport failure");
  const { session, directory } = await fakeSession(t, async () => { throw cause; });
  await assert.rejects(session.close(), (error) => error.code === "filesystem_failure"
    && error.phase === "cleanup" && error.cause.errors[0] === cause);
  await assert.rejects(stat(directory), { code: "ENOENT" });
});

test("a verified PDF survives Playwright force-closing a stopped owned browser", {
  skip: process.platform === "win32" ? "SIGSTOP requires a POSIX host" : false,
  timeout: 120_000,
}, async (t) => {
  const launch = chromium.launch.bind(chromium);
  let browser, browserPid, close, directory;
  t.mock.method(chromium, "launch", async (options) => {
    directory = options.env.HOME;
    browser = await launch(options);
    close = browser.close.bind(browser);
    t.mock.method(browser, "close", async () => {
      const session = await browser.newBrowserCDPSession();
      const info = await session.send("SystemInfo.getProcessInfo");
      browserPid = info.processInfo.find((entry) => entry.type === "browser").id;
      await session.detach();
      process.kill(browserPid, "SIGSTOP");
      return close();
    });
    return browser;
  });
  try {
    const result = await renderDocxArtifacts(generateBrowserCleanupDocx(), {
      outputs: ["pdf"], reviewProfile: "final", commentProfile: "hidden", timeoutMs: 120_000,
      ...(process.env.DOCXODUS_CHROMIUM_PATH
        ? { browserExecutablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
    });
    const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const task = getDocument({ data: new Uint8Array(result.pdf), verbosity: 0 });
    try {
      const pdf = await task.promise;
      assert.equal(pdf.numPages, 1);
      const content = await (await pdf.getPage(1)).getTextContent();
      assert.equal(content.items.filter((item) => "str" in item).map((item) => item.str)
        .filter(Boolean).join(" ").replace(/\s+/g, " "),
      "CEDAR WORKSHOP INSPECTION Seven tools were returned to cabinet 24. Inspection complete. Signed: Rowan Vale.");
    } finally { await task.destroy(); }
    assert.equal(browser.isConnected(), false);
    assert.throws(() => process.kill(browserPid, 0), { code: "ESRCH" });
    await assert.rejects(stat(directory), { code: "ENOENT" });
  } finally {
    if (browserPid) {
      try { process.kill(browserPid, "SIGCONT"); } catch (error) {
        if (error.code !== "ESRCH") throw error;
      }
    }
    await close?.();
  }
});
