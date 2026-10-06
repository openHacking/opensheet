import { spawn } from 'node:child_process';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const server = spawn(
  process.execPath,
  [
    fileURLToPath(new URL('../node_modules/vite/bin/vite.js', import.meta.url)),
    'preview',
    '--config',
    'apps/playground/vite.config.ts',
    '--host',
    '127.0.0.1',
    '--port',
    '5185',
    '--strictPort',
  ],
  {
    env: { ...process.env, OPENSHEET_BASE_PATH: '/opensheet/' },
    stdio: 'ignore',
  },
);
let browser;
try {
  let ready = false;
  for (let attempt = 0; attempt < 60; attempt++) {
    if (server.exitCode !== null)
      throw new Error('Production preview exited before becoming ready');
    try {
      ready = (await fetch('http://127.0.0.1:5185/opensheet/')).ok;
    } catch {}
    if (ready) break;
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  assert.ok(ready, 'Production preview must start');
  browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('response', (response) => {
    if (response.status() >= 400) errors.push(`${response.status()} ${response.url()}`);
  });
  for (const id of ['budget', 'sales', 'planner', 'code']) {
    await page.goto(`http://127.0.0.1:5185/opensheet/#${id}`);
    await page.waitForSelector('.os-grid');
    await page.reload();
    await page.waitForSelector('.os-grid');
    assert.equal(await page.locator(`[data-scene="${id}"]`).getAttribute('aria-current'), 'page');
  }
  const { utils, write } = await import('xlsx');
  const workbook = utils.book_new(
    utils.aoa_to_sheet([
      ['Name', 'Count'],
      ['Production', 7],
    ]),
    'Imported',
  );
  await page.locator('#file').setInputFiles({
    name: 'production.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: write(workbook, { type: 'buffer', bookType: 'xlsx' }),
  });
  await page.getByRole('tab', { name: 'Imported', exact: true }).waitFor();
  assert.equal(
    await page.evaluate(
      () => window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues()[0][0],
    ),
    7,
  );
  assert.deepEqual(errors, []);
  console.log('PASS: production Pages paths, all scene reloads, assets and XLSX Worker import.');
} finally {
  await browser?.close();
  server.kill('SIGTERM');
}
