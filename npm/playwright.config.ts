import { defineConfig, devices } from '@playwright/test';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { serialLaneSpecs } from './scripts/playwright-lanes.mjs';

// The visual-parity benchmark pins font substitution for BOTH renderers (issue #379): the spec
// passes the same file to LibreOffice via FONTCONFIG_FILE, and Chromium must be LAUNCHED under
// it, which only the config can do. Scoped to the benchmark opt-in so ordinary specs and their
// committed snapshots keep the host's default fonts.
const visualParityLaunchEnv = process.env.DOCXODUS_VISUAL_PARITY === '1'
  ? {
      env: {
        ...process.env as Record<string, string>,
        FONTCONFIG_FILE: resolve(
          dirname(fileURLToPath(import.meta.url)), 'tests/visual-parity/fonts.conf'),
      },
    }
  : {};

// Sandboxed dev environments often pre-install one pinned Chromium build and
// block downloads; DOCXODUS_CHROMIUM_PATH points the chromium-based projects
// at it instead of the per-version browser Playwright would fetch. Unset (the
// normal case, including CI), Playwright resolves its own browsers.
const chromiumExecutable = process.env.DOCXODUS_CHROMIUM_PATH
  ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH }
  : {};

// The suite runs as two lanes, chosen by DOCXODUS_PLAYWRIGHT_LANE (scripts/playwright-lanes.mjs
// sets it for `npm test`, lane by lane; each CI lane job sets it the same way):
//   pool   — every spec except the serial list, on several workers at once. Each test boots
//            its own page and nothing is shared across spec files, so running them side by
//            side costs CPU and nothing else.
//   serial — the specs listed below, on one worker, so that no other worker's WASM boot
//            inflates the latency, throughput or frame-rate figures they assert. Firefox runs
//            in this lane too, so only its CI job installs that browser.
//   unset  — every project on one worker: the pre-lane behaviour, for an ad-hoc
//            `npx playwright test` that should still reach everything.
// A new spec that asserts a wall-clock bound goes in serialLaneSpecs (scripts/playwright-lanes.mjs);
// scripts/playwright-lanes.test.mjs fails when a spec file is in both lanes or in neither.
const lane = process.env.DOCXODUS_PLAYWRIGHT_LANE;
if (lane !== undefined && lane !== 'pool' && lane !== 'serial') {
  throw new Error(`DOCXODUS_PLAYWRIGHT_LANE must be "pool" or "serial", not "${lane}"`);
}
const serialLaneMatch = new RegExp(
  `[\\\\/](${serialLaneSpecs.map(name => name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})$`);

const chromiumUse = {
  ...devices['Desktop Chrome'],
  launchOptions: { ...visualParityLaunchEnv, ...chromiumExecutable },
};
const poolProjects = [
  {
    name: 'chromium',
    testIgnore: [/demo-arcade-mobile\.spec\.ts/, serialLaneMatch],
    use: chromiumUse,
  },
  {
    // Phone-shaped rig for the mobile arcade garble (device-inflated text
    // re-breaking lines the document authored at a fixed column width):
    // mobile viewport, touch, and the fit-to-width CSS zoom that a phone
    // actually exercises. Its spec recreates the platform's text inflation
    // itself, so it runs here only and the desktop chromium project skips it.
    name: 'chromium-pixel5',
    testMatch: /demo-arcade-mobile\.spec\.ts/,
    use: {
      ...devices['Pixel 5'],
      launchOptions: { ...chromiumExecutable },
    },
  },
];
const serialProjects = [
  {
    name: 'chromium-serial',
    testMatch: serialLaneMatch,
    use: chromiumUse,
  },
  {
    // Firefox is the canary for cross-contenteditable drag feedback: its native selection
    // update runs after mousemove dispatch and used to erase the bridged Range until mouseup.
    name: 'firefox-cross-block-selection',
    testMatch: /editor-multiblock-format\.spec\.ts/,
    use: { ...devices['Desktop Firefox'] },
  },
];

export default defineConfig({
  testDir: './tests',
  fullyParallel: false, // a spec file's tests run in order on one worker; files spread across workers
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  // Pool lane: four workers on a 4-vCPU hosted runner, Playwright's default (half the cores)
  // elsewhere — four workers on an 8-core host ran the pool 3.3x faster than one, and three on
  // the hosted runner 2.25x. Any other lane, or none, keeps one worker: see the lane comment above.
  workers: lane === 'pool' ? (process.env.CI ? 4 : undefined) : 1,
  // Lane-specific folders, so `npm test` running the two lanes back to back keeps both
  // reports and both sets of traces; CI merges the lanes' blob reports into one HTML report.
  reporter: [['html', { outputFolder: lane === 'serial' ? 'playwright-report-serial' : 'playwright-report' }]],
  outputDir: lane === 'serial' ? 'test-results-serial' : 'test-results',
  timeout: 60000, // WASM loading can be slow
  use: {
    baseURL: 'http://localhost:8082',
    trace: 'on-first-retry',
  },
  // Snapshot configuration for visual testing
  expect: {
    toHaveScreenshot: {
      maxDiffPixelRatio: 0.05,
      threshold: 0.2,
    },
  },
  snapshotDir: './tests/__snapshots__',
  snapshotPathTemplate: '{snapshotDir}/{testFilePath}/{arg}{ext}',
  projects: lane === 'pool' ? poolProjects
    : lane === 'serial' ? serialProjects
    : [...poolProjects, ...serialProjects],
  webServer: [
    {
      command: 'python3 -m http.server 8084 --directory dist/site',
      url: 'http://localhost:8084/demo/app.html',
      reuseExistingServer: !process.env.CI,
      timeout: 30000,
    },
    {
      command: 'python3 -m http.server 8082 --directory dist/wasm',
      url: 'http://localhost:8082',
      reuseExistingServer: !process.env.CI,
      timeout: 30000,
    },
    {
      // Second origin with CORS headers — emulates a CDN (jsDelivr/unpkg) so
      // the cdn-embed spec exercises genuinely cross-origin module + WASM loads.
      command: 'python3 tests/cors-server.py 8083 dist',
      url: 'http://localhost:8083/embed.bundle.js',
      reuseExistingServer: !process.env.CI,
      timeout: 30000,
    },
  ],
});
