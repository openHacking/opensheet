import { test, expect } from '@playwright/test';
test('budget indicators reflect edits and undo', async ({ page }) => {
  await page.goto('/#budget');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await expect(page.locator('.metric strong').first()).toHaveText('$13,097');
  await page.evaluate(
    async () =>
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('D2')
        .setValues([[2]]),
  );
  await expect(page.locator('.metric strong').first()).toHaveText('$15,497');
  await page.evaluate(async () => await window.opensheet.getWorkbook().undo());
  await expect(page.locator('.metric strong').first()).toHaveText('$13,097');
});
test('sales chart follows workbook data and resets', async ({ page }) => {
  await page.goto('/#sales');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await expect(page.locator('.metric strong').first()).toHaveText('$179,400');
  await page.evaluate(
    async () =>
      await window.opensheet
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
  await expect(
    page.getByRole('button', { name: 'Σ Sum', exact: true, includeHidden: true }),
  ).toHaveCount(1);
});
test('planner timeline follows dates and completion', async ({ page }) => {
  await page.goto('/#planner');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await expect(page.locator('.timeline-row')).toHaveCount(6);
  await page.evaluate(
    async () =>
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('C2:E2')
        .setValues([['2026-10-05', 7, 0.5]]),
  );
  await expect(page.locator('.timeline-row').first()).toContainText('50%');
  await page.evaluate(
    async () =>
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('C2')
        .setValues([['invalid']]),
  );
  await expect(page.locator('.timeline-row').first()).toContainText('Enter a valid start date');
});
test('table code updates and downloads a complete export', async ({ page }) => {
  await page.goto('/#code');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await expect(page.getByRole('tab', { name: 'Code export' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(page.locator('#source')).toContainText('Baseline');
  await expect(page.locator('#source')).toContainText('\\begin{tabular}{llll}');
  await page.evaluate(
    async () =>
      await window.opensheet
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
test('demo selector, history and reload preserve the chosen scene', async ({ page }) => {
  await page.goto('/#sales');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await page.getByRole('combobox', { name: 'Choose demo' }).selectOption('planner');
  await expect(page.locator('#scene-title')).toHaveText('Project Planner');
  await page.goBack();
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await expect(page.getByRole('combobox', { name: 'Choose demo' })).toHaveValue('sales');
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
});
test('switching a modified workbook requires an explicit discard', async ({ page }) => {
  await page.goto('/#budget');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await page.evaluate(
    async () =>
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A2')
        .setValues([['Unsaved change']]),
  );
  page.once('dialog', (dialog) => dialog.dismiss());
  await page.getByRole('combobox', { name: 'Choose demo' }).selectOption('sales');
  await expect(page).toHaveURL(/#budget$/);
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await expect(page.getByRole('combobox', { name: 'Choose demo' })).toHaveValue('budget');
  page.once('dialog', (dialog) => dialog.accept());
  await page.getByRole('combobox', { name: 'Choose demo' }).selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
});
test('desktop workbench fits the viewport and resizes both panes', async ({ page }) => {
  for (const viewport of [
    { width: 1280, height: 800 },
    { width: 1440, height: 900 },
  ]) {
    await page.setViewportSize(viewport);
    for (const scene of ['budget', 'sales', 'planner', 'code']) {
      await page.goto(`/#${scene}`);
      await page.waitForFunction(
        () =>
          !!window.opensheet?.selection.get() &&
          !!document.querySelector('#scene-title')?.textContent,
      );
      await expect(page.locator('#spreadsheet .os-grid')).toBeVisible();
      await expect(page.locator('.side-panel')).toBeVisible();
      const dimensions = await page.evaluate(() => ({
        width: document.documentElement.scrollWidth,
        height: document.documentElement.scrollHeight,
      }));
      expect(dimensions.width).toBeLessThanOrEqual(viewport.width);
      expect(dimensions.height).toBeLessThanOrEqual(viewport.height);
      for (const selector of ['.workbench', '.document', '.side-panel']) {
        const bounds = await page.locator(selector).boundingBox();
        expect(bounds!.y + bounds!.height).toBeCloseTo(viewport.height, 0);
      }
    }
  }
  const splitter = page.getByRole('separator', { name: 'Resize spreadsheet and results' });
  const before = await page.locator('.document').boundingBox();
  const handle = await splitter.boundingBox();
  await page.mouse.move(handle!.x + handle!.width / 2, handle!.y + handle!.height / 2);
  await page.mouse.down();
  await page.mouse.move(handle!.x - 60, handle!.y + handle!.height / 2, { steps: 4 });
  await page.mouse.up();
  const dragged = await page.locator('.document').boundingBox();
  expect(dragged!.width).toBeLessThan(before!.width);
  await splitter.focus();
  await splitter.press('ArrowRight');
  const after = await page.locator('.document').boundingBox();
  expect(after!.width).toBeGreaterThan(dragged!.width);
  await splitter.dblclick();
  await expect(splitter).toHaveAttribute('aria-valuenow', '67');
});
test('result tabs expose the code preview', async ({ page }) => {
  await page.goto('/#budget');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await page.getByRole('tab', { name: 'Code export', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Code export' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(page.locator('#source')).toBeVisible();
  await page.getByRole('tab', { name: 'Data view' }).click();
  await expect(page.locator('.metric strong').first()).toBeVisible();
});
test('developer pages use main content and preserve workbook edits when returning', async ({
  page,
}) => {
  await page.goto('/#sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('B2')
      .setValues([[36000]]);
  });
  await expect(page.locator('.metric strong').first()).toHaveText('$197,400');
  await expect(page.getByRole('button', { name: 'Table generator' })).toHaveCount(0);
  await page.getByRole('button', { name: 'API quick start', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'API quick start' })).toBeVisible();
  await expect(page.locator('#api-example')).toContainText('createOpenSheet');
  await expect(page.locator('#nav-api')).toHaveAttribute('aria-current', 'page');
  await expect(page.locator('#workbench')).toBeHidden();
  await expect(page.getByRole('dialog')).toBeHidden();
  await page.getByRole('button', { name: 'Plugin example', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'Plugin example' })).toBeVisible();
  await expect(page.locator('#plugin-example')).toContainText('definePlugin');
  await page.goBack();
  await expect(page.getByRole('heading', { name: 'API quick start' })).toBeVisible();
  await page.getByRole('button', { name: /Spreadsheet/ }).click();
  await expect(page.locator('#workbench')).toBeVisible();
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await expect(page.locator('.metric strong').first()).toHaveText('$197,400');
});
test('developer pages can be opened directly and reloaded', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/#plugins');
  await expect(page.getByRole('heading', { name: 'Plugin example' })).toBeVisible();
  await page.reload();
  await expect(page.getByRole('heading', { name: 'Plugin example' })).toBeVisible();
  await page.getByRole('button', { name: 'Back to spreadsheet', exact: false }).click();
  await expect(page.locator('#workbench')).toBeVisible();
});
