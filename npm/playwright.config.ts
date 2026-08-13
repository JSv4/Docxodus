import { defineConfig, devices } from '@playwright/test';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

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

export default defineConfig({
  testDir: './tests',
  fullyParallel: false, // WASM tests need sequential execution
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: 1, // Single worker for WASM
  reporter: 'html',
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
  projects: [
    {
      name: 'chromium',
      testIgnore: /demo-arcade-mobile\.spec\.ts/,
      use: {
        ...devices['Desktop Chrome'],
        launchOptions: { ...visualParityLaunchEnv, ...chromiumExecutable },
      },
    },
    {
      // Firefox is the canary for cross-contenteditable drag feedback: its native selection
      // update runs after mousemove dispatch and used to erase the bridged Range until mouseup.
      name: 'firefox-cross-block-selection',
      testMatch: /editor-multiblock-format\.spec\.ts/,
      use: { ...devices['Desktop Firefox'] },
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
  ],
  webServer: [
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
