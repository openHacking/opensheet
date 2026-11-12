import { resolve } from 'node:path';
import { expect, test } from '@playwright/test';

test('activity example observes real edits, restores and removes cleanly', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const moduleUrl = `/@fs/${resolve('examples/plugin-activity.ts')}`;
  await page.evaluate(async (url) => {
    const { activityCounter } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    app.use(activityCounter);
    await app
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[42]]);
  }, moduleUrl);
  await page.getByRole('button', { name: 'More tools' }).click();
  await page.getByRole('button', { name: 'Changes', exact: true }).click();
  await expect(page.locator('.os-toast')).toHaveText('1 committed changes in this session');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const old = app.plugins.get<{ getCount(): number }>('example.activity-counter')!;
    await app.open(app.getWorkbook().id);
    const fresh = app.plugins.get<{ getCount(): number }>('example.activity-counter')!;
    const initial = fresh.getCount();
    await app
      .getWorkbook()
      .getSheets()[0]
      .range('A2')
      .setValues([[17]]);
    const after = fresh.getCount();
    app.plugins.remove('example.activity-counter');
    await app
      .getWorkbook()
      .getSheets()[0]
      .range('A3')
      .setValues([[9]]);
    const removedCount = fresh.getCount();
    const values = await app.getWorkbook().getSheets()[0].range('A1:A3').getValues();
    return { old: old.getCount(), initial, after, removedCount, values };
  });
  expect(result).toEqual({
    old: 1,
    initial: 0,
    after: 1,
    removedCount: 1,
    values: [[42], [17], [9]],
  });
  await expect(page.locator('[data-action="example.activity-counter.show"]')).toHaveCount(0);
});

test('conformance helper removes actual workbook subscriptions', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async ({ sdk, example }) => {
      const { checkPluginConformance } = await import(/* @vite-ignore */ sdk);
      const { activityCounter } = await import(/* @vite-ignore */ example);
      const book = window.opensheet.getWorkbook();
      const original = book.onCommit;
      let subscriptions = 0,
        removals = 0,
        callbacks = 0;
      book.onCommit = (listener) => {
        subscriptions++;
        const remove = original.call(book, (commit) => {
          callbacks++;
          listener(commit);
        });
        return () => {
          removals++;
          remove();
        };
      };
      const report = checkPluginConformance(() => [activityCounter], {
        getWorkbook: () => book,
        getSelection: () => null,
        addToolbar: () => () => {},
        notify() {},
        onError() {},
      });
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[79]]);
      book.onCommit = original;
      return { report, subscriptions, removals, callbacks };
    },
    {
      sdk: `/@fs/${resolve('packages/plugin-sdk/src/index.ts')}`,
      example: `/@fs/${resolve('examples/plugin-activity.ts')}`,
    },
  );
  expect(result.report.passed).toBe(true);
  expect(result).toMatchObject({ subscriptions: 2, removals: 2, callbacks: 0 });
});
