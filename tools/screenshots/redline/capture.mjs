// Capture the converter output at the README asset's historical 1854x1037 dimensions.
// The 1.4 device scale preserves the original 1324x741 CSS viewport and makes the legal
// text comfortably legible when GitHub scales the image responsively.

import { pathToFileURL } from 'node:url';
import path from 'node:path';

const playwrightModule = process.env.PLAYWRIGHT_MODULE
  || new URL('../../../npm/node_modules/playwright/index.mjs', import.meta.url).href;
const { chromium } = await import(playwrightModule);

if (process.argv.length !== 4) {
  console.error('usage: node capture.mjs <redline.html> <redline.png>');
  process.exit(1);
}

const htmlPath = path.resolve(process.argv[2]);
const outputPath = path.resolve(process.argv[3]);
const browser = await chromium.launch({
  headless: true,
  executablePath: process.env.CHROME_PATH || '/usr/bin/google-chrome',
});

try {
  const context = await browser.newContext({
    viewport: { width: 1324, height: 741 },
    deviceScaleFactor: 1.4,
  });
  const page = await context.newPage();
  await page.goto(pathToFileURL(htmlPath).href);

  const section = page
    .locator('h1, h2, h3, h4, h5, h6')
    .filter({ hasText: 'Voting Provisions Regarding the Board' })
    .first();
  await section.evaluate(element => element.scrollIntoView({ block: 'start' }));
  await page.evaluate(() => window.scrollBy(0, -8));
  await page.screenshot({ path: outputPath });
  await context.close();
} finally {
  await browser.close();
}

console.log(`screenshot: ${outputPath}`);
