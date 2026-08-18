import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { createSharedAbortableLoader, runtimeAssetGraphDigest } from "../dist/assets.js";
import { PINNED_CHROMIUM_BUILD } from "../dist/browser-session.js";
import { canonicalJson } from "../dist/canonical.js";

const here = dirname(fileURLToPath(import.meta.url));
const repositoryRoot = dirname(dirname(here));

test("Node and browser use one canonical runtime asset-graph identity", async () => {
  const manifest = JSON.parse(await readFile(
    join(repositoryRoot, "npm", "dist", "export-assets.json"),
    "utf8",
  ));
  const expected = createHash("sha256").update(canonicalJson({
    schemaVersion: manifest.schemaVersion,
    packageVersion: manifest.packageVersion,
    assets: manifest.assets,
  })).digest("hex");
  assert.equal(runtimeAssetGraphDigest(manifest), expected);
});

test("pinned Chromium build matches Playwright's installed metadata", async () => {
  const metadata = JSON.parse(await readFile(
    join(repositoryRoot, "npm-export", "node_modules", "playwright-core", "browsers.json"),
    "utf8",
  ));
  assert.equal(
    metadata.browsers.find(({ name }) => name === "chromium")?.browserVersion,
    PINNED_CHROMIUM_BUILD,
  );
});

test("shared asset loading cancels per waiter and never caches an all-aborted load", async () => {
  let calls = 0;
  const loader = createSharedAbortableLoader((signal) => new Promise((resolve, reject) => {
    calls++;
    const timer = setTimeout(() => resolve(`value-${calls}`), 20);
    signal.addEventListener("abort", () => {
      clearTimeout(timer);
      reject(new Error("shared operation aborted"));
    }, { once: true });
  }));
  const firstController = new AbortController();
  const first = loader(firstController.signal);
  const second = loader();
  firstController.abort();
  await assert.rejects(first, (error) => error?.code === "operation_cancelled");
  assert.equal(await second, "value-1");
  assert.equal(await loader(), "value-1");
  assert.equal(calls, 1);

  let abortedCalls = 0;
  const allAbortedLoader = createSharedAbortableLoader((signal) => new Promise((resolve, reject) => {
    abortedCalls++;
    const timer = setTimeout(() => resolve(`retry-${abortedCalls}`), 20);
    signal.addEventListener("abort", () => {
      clearTimeout(timer);
      reject(new Error("all waiters aborted"));
    }, { once: true });
  }));
  const onlyController = new AbortController();
  const abandoned = allAbortedLoader(onlyController.signal);
  onlyController.abort();
  await assert.rejects(abandoned, (error) => error?.code === "operation_cancelled");
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(await allAbortedLoader(), "retry-2");
  assert.equal(abortedCalls, 2);
});
