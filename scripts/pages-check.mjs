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
  await page.goto('http://127.0.0.1:5185/opensheet/');
  const entry = page.getByRole('link', { name: 'Try all features' });
  assert.equal(await entry.getAttribute('href'), '#features');
  await page.waitForFunction(() => document.querySelector('#scene-title')?.textContent);
  await entry.click();
  await page.getByRole('heading', { name: 'All Features' }).waitFor();
  await page.getByRole('button', { name: 'Rates!B2', exact: true }).click();
  assert.equal(
    await page.getByRole('tab', { name: 'Rates', exact: true }).getAttribute('aria-selected'),
    'true',
  );
  for (const id of ['features', 'budget', 'sales', 'planner', 'code', 'performance']) {
    await page.goto(`http://127.0.0.1:5185/opensheet/#${id}`);
    await page.waitForSelector('.os-grid');
    await page.waitForFunction(() => !document.querySelector('#scene-select')?.disabled);
    await page.reload();
    await page.waitForSelector('.os-grid');
    await page.waitForFunction(() => !document.querySelector('#scene-select')?.disabled);
    assert.equal(await page.locator('#scene-select').inputValue(), id);
  }
  await page.goto('http://127.0.0.1:5185/opensheet/#code');
  await page.waitForFunction(() => !document.querySelector('#scene-select')?.disabled);
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
  const imported = page.getByRole('tab', { name: 'Imported', exact: true });
  const convert = page.getByRole('button', { name: 'Import converted workbook', exact: true });
  await imported.or(convert).first().waitFor();
  if (await convert.isVisible()) await convert.click();
  await imported.waitFor();
  assert.equal(
    await page.evaluate(
      async () =>
        (await window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues())[0][0],
    ),
    7,
  );
  assert.deepEqual(errors, []);
  console.log('PASS: production Pages paths, all scene reloads, assets and XLSX Worker import.');
} finally {
  await browser?.close();
  server.kill('SIGTERM');
}
