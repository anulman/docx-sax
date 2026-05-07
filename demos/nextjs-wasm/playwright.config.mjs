import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  timeout: 60_000,
  webServer: {
    command: 'npm run dev -- --hostname 127.0.0.1 --port 3016',
    url: 'http://127.0.0.1:3016',
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
  use: {
    ...devices['Desktop Chrome'],
    baseURL: 'http://127.0.0.1:3016',
  },
});
