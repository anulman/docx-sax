import { expect, test } from '@playwright/test';
import path from 'node:path';
import { createGeneratedDocxBytes } from '../../../test/js/generated-docx.mjs';

const simpleFixturePath = path.resolve('public/fixtures/simple.docx');
const chartFixturePath = path.resolve('public/fixtures/chart-2d-column.docx');

async function collectBrowserErrors(page) {
  const errors = [];
  page.on('pageerror', (error) => errors.push(error));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(new Error(message.text()));
  });
  return errors;
}

async function openReadyDemo(page) {
  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'Next.js DOCX event preview' })).toBeVisible();
  await expect(page.getByRole('status')).toContainText('WASM runtime ready', { timeout: 45_000 });
}

test('uploads public fixture, loads WASM, parses events, and renders preview', async ({ page }) => {
  const errors = await collectBrowserErrors(page);

  await openReadyDemo(page);
  await page.getByLabel('DOCX file').setInputFiles(simpleFixturePath);

  await expect(page.getByRole('status')).toContainText(/Parsed \d+ events from simple\.docx/, { timeout: 45_000 });
  await expect(page.getByLabel('Rendered DOCX preview')).toContainText('Hello DOCX SAX');
  await expect(page.getByRole('heading', { name: '/word/document.xml' })).toBeVisible();
  await expect(page.getByText('Text events')).toBeVisible();

  expect(errors).toEqual([]);
});

test('reports partial parse progress before final completion for a large generated DOCX', async ({ page }) => {
  const errors = await collectBrowserErrors(page);

  await openReadyDemo(page);
  await page.evaluate(() => {
    const status = document.querySelector('[role="status"]');
    window.__docxSaxStatuses = status ? [status.textContent] : [];
    if (status) {
      const observer = new MutationObserver(() => window.__docxSaxStatuses.push(status.textContent));
      observer.observe(status, { childList: true, characterData: true, subtree: true });
      window.__docxSaxStatusObserver = observer;
    }
  });

  await page.getByLabel('DOCX file').setInputFiles({
    name: 'large-generated.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: Buffer.from(createGeneratedDocxBytes({ paragraphs: 6000 })),
  });

  await expect(page.getByRole('status')).toContainText(/Parsed \d+ events from large-generated\.docx\./, { timeout: 60_000 });
  await expect(page.getByLabel('Rendered DOCX preview')).toContainText('Generated DOCX SAX paragraph 0');
  await expect(page.getByLabel('Rendered DOCX preview')).toContainText('Generated DOCX SAX paragraph 5999');
  await expect(page.getByRole('heading', { name: '/word/document.xml' })).toBeVisible();

  const statuses = await page.evaluate(() => window.__docxSaxStatuses ?? []);
  const partialStatuses = statuses.filter((status) => /Parsed \d+ events from large-generated\.docx…/.test(status));
  const finalStatus = statuses.findLast((status) => /Parsed \d+ events from large-generated\.docx\./.test(status));

  expect(partialStatuses.length, `statuses: ${JSON.stringify(statuses)}`).toBeGreaterThanOrEqual(2);
  expect(finalStatus, `statuses: ${JSON.stringify(statuses)}`).toBeTruthy();

  const firstPartialCount = Number(partialStatuses[0].match(/Parsed (\d+) events/)?.[1]);
  const lastPartialCount = Number(partialStatuses.at(-1).match(/Parsed (\d+) events/)?.[1]);
  const finalCount = Number(finalStatus.match(/Parsed (\d+) events/)?.[1]);

  expect(firstPartialCount).toBeGreaterThan(0);
  expect(lastPartialCount).toBeGreaterThan(firstPartialCount);
  expect(finalCount).toBeGreaterThan(lastPartialCount);

  expect(errors).toEqual([]);
});

test('renders representative text in the right-hand preview for DOCX chart parts', async ({ page }) => {
  const errors = await collectBrowserErrors(page);

  await openReadyDemo(page);
  await page.getByLabel('DOCX file').setInputFiles(chartFixturePath);

  await expect(page.getByRole('status')).toContainText(/Parsed \d+ events from chart-2d-column\.docx/, { timeout: 45_000 });

  const previewCard = page.locator('.card.preview');
  const preview = page.getByLabel('Rendered DOCX preview');
  await expect(preview).toBeVisible();
  await expect(preview).toContainText('Series 1');
  await expect(preview).toContainText('Category 1');
  await expect(preview).toContainText('4.3');
  await expect(preview).not.toContainText('Parsed document text will appear here.');

  await expect(page.getByRole('heading', { name: '/word/charts/chart1.xml' })).toBeVisible();
  await expect(page.getByRole('heading', { name: '/word/charts/chart2.xml' })).toBeVisible();
  await expect(page.getByRole('heading', { name: '/word/charts/chart3.xml' })).toBeVisible();
  await expect(page.locator('dl').filter({ hasText: 'Text events' })).toContainText('117');

  const [uploadBox, previewBox, articleBox] = await Promise.all([
    page.locator('.grid > .card').first().boundingBox(),
    previewCard.boundingBox(),
    preview.boundingBox(),
  ]);

  expect(uploadBox, 'upload card should have a bounding box').not.toBeNull();
  expect(previewBox, 'right-hand preview card should have a bounding box').not.toBeNull();
  expect(articleBox, 'rendered preview article should have a bounding box').not.toBeNull();
  expect(previewBox.x, 'preview card should be positioned to the right of the upload card').toBeGreaterThan(uploadBox.x + uploadBox.width - 1);
  expect(articleBox.x, 'visible preview content should render inside the right-hand preview card').toBeGreaterThanOrEqual(previewBox.x);
  expect(articleBox.width, 'visible preview content should occupy horizontal space').toBeGreaterThan(100);
  expect(articleBox.height, 'visible preview content should occupy vertical space').toBeGreaterThan(20);
  await expect(previewCard).not.toHaveCSS('overflow', 'auto');

  expect(errors).toEqual([]);
});
