import { chromium } from '@playwright/test';
import { createServer } from 'vite';

const server = await createServer({
  root: '../..',
  server: { host: '127.0.0.1', port: 5177, strictPort: true },
  logLevel: 'error',
});

await server.listen();

const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage();
  page.on('console', (message) => console.log(message.text()));
  page.on('pageerror', (error) => console.error(error));
  await page.goto('http://127.0.0.1:5177/packages/browser-wasm/smoke/index.html');
  const result = await page.waitForFunction(() => window.__docxSaxSmokeResult, undefined, { timeout: 30_000 });
  console.log(JSON.stringify(await result.jsonValue()));
} finally {
  await browser.close();
  await server.close();
}
