import { test as base, expect } from '@playwright/test';
import { build } from 'esbuild';
import type * as HistoryApi from '../src/index.js';

declare global { interface Window { historyApi: typeof HistoryApi } }

export const test = base.extend<{ historyPage: void }>({
  historyPage: [async ({ page }, use) => {
    const bundle = await build({ entryPoints: ['dist/index.js'], bundle: true, format: 'esm', write: false });
    await page.route('**/history-controls.html', route => route.fulfill({ contentType: 'text/html',
      body: '<!doctype html><html lang="en"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Document history</title><main id="controls"></main></html>' }));
    await page.route('**/history-controls-api.js', route => route.fulfill({ contentType: 'text/javascript', body: bundle.outputFiles[0].text }));
    await page.goto('http://localhost:8083/history-controls.html');
    await page.evaluate(async () => {
      const url = '/history-controls-api.js';
      window.historyApi = await import(url);
      await window.historyApi.initialize('http://localhost:8083/wasm/');
    });
    await use();
  }, { auto: true }],
});
export { expect };
