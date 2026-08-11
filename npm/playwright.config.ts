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
      use: { ...devices['Desktop Chrome'], launchOptions: visualParityLaunchEnv },
    },
    {
      // Firefox is the canary for cross-contenteditable drag feedback: its native selection
      // update runs after mousemove dispatch and used to erase the bridged Range until mouseup.
      name: 'firefox-cross-block-selection',
      testMatch: /editor-multiblock-format\.spec\.ts/,
      use: { ...devices['Desktop Firefox'] },
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
