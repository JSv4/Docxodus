import { defineConfig, devices } from '@playwright/test';

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
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
  ],
  webServer: {
    command: 'python3 -m http.server 8082 --directory dist/wasm',
    url: 'http://localhost:8082',
    reuseExistingServer: !process.env.CI,
    timeout: 30000,
  },
});
