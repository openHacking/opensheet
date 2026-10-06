import { test, expect } from '@playwright/test';

test('budget indicators reflect edits and undo', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await expect(page.locator('.metric strong').first()).toHaveText('$13,097');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('D2')
      .setValues([[2]]),
  );
  await expect(page.locator('.metric strong').first()).toHaveText('$15,497');
  await page.evaluate(() => window.opensheet.getWorkbook().undo());
  await expect(page.locator('.metric strong').first()).toHaveText('$13,097');
});

test('sales chart follows workbook data and resets', async ({ page }) => {
  await page.goto('/#sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await expect(page.locator('.metric strong').first()).toHaveText('$179,400');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('B2')
      .setValues([[36000]]),
  );
  await expect(page.locator('.metric strong').first()).toHaveText('$197,400');
  await expect(page.locator('.revenue-bar').first()).toHaveAttribute('title', '$36,000');
  page.once('dialog', (dialog) => dialog.accept());
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.locator('.metric strong').first()).toHaveText('$179,400');
  await expect(page.getByRole('button', { name: 'Σ Sum', exact: true })).toHaveCount(1);
});

test('planner timeline follows dates and completion', async ({ page }) => {
  await page.goto('/#planner');
  await expect(page.locator('.timeline-row')).toHaveCount(6);
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('C2:E2')
      .setValues([['2026-10-05', 7, 0.5]]),
  );
  await expect(page.locator('.timeline-row').first()).toContainText('50%');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('C2')
      .setValues([['invalid']]),
  );
  await expect(page.locator('.timeline-row').first()).toContainText('Enter a valid start date');
});

test('table code updates and downloads a complete export', async ({ page }) => {
  await page.goto('/#code');
  await expect(page.locator('#source')).toContainText('Baseline');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A2')
      .setValues([['Updated model']]),
  );
  await page.getByRole('tab', { name: 'Markdown', exact: true }).click();
  await expect(page.locator('#source')).toContainText('Updated model');
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  expect((await downloaded).suggestedFilename()).toBe('Table to Code.md');
});

test('demo links, history and reload preserve the chosen scene', async ({ page }) => {
  await page.goto('/#sales');
  await page.locator('[data-scene="planner"]').click();
  await expect(page.locator('#scene-title')).toHaveText('Project Planner');
  await page.goBack();
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
});

test('switching a modified workbook requires an explicit discard', async ({ page }) => {
  await page.goto('/#budget');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A2')
      .setValues([['Unsaved change']]),
  );
  page.once('dialog', (dialog) => dialog.dismiss());
  await page.locator('[data-scene="sales"]').click();
  await expect(page).toHaveURL(/#budget$/);
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  page.once('dialog', (dialog) => dialog.accept());
  await page.locator('[data-scene="sales"]').click();
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
});
