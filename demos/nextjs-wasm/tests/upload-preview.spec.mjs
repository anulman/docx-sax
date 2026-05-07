import { expect, test } from '@playwright/test';
import path from 'node:path';

const fixturePath = path.resolve('public/fixtures/simple.docx');

test('uploads public fixture, loads WASM, parses events, and renders preview', async ({ page }) => {
  const errors = [];
  page.on('pageerror', (error) => errors.push(error));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(new Error(message.text()));
  });

  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'Next.js DOCX event preview' })).toBeVisible();

  await page.getByLabel('DOCX file').setInputFiles(fixturePath);

  await expect(page.getByRole('status')).toContainText(/Parsed \d+ events from simple\.docx/, { timeout: 45_000 });
  await expect(page.getByLabel('Rendered DOCX preview')).toContainText('Hello DOCX SAX');
  await expect(page.getByText('/word/document.xml')).toBeVisible();
  await expect(page.getByText('Text events')).toBeVisible();

  expect(errors).toEqual([]);
});
