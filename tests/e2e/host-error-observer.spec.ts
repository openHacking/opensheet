import { expect, test } from '@playwright/test';

test('a throwing host error callback does not suppress later commit observers', async ({
  page,
}) => {
  const uncaught: string[] = [];
  page.on('pageerror', (error) => uncaught.push(error.message));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook();
    const errors: string[] = [],
      commits: number[] = [];
    (app as any).options.onError = (error: Error) => {
      errors.push(error.message);
      throw new Error('host error handler failed');
    };
    const stopBad = app.on('workbook:committed', () => {
      throw new Error('commit observer failed');
    });
    const stopGood = app.on('workbook:committed', (commit) => commits.push(commit.revision));
    const revision = book.revision;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[1234]]);
    stopBad();
    stopGood();
    return {
      revision,
      after: book.revision,
      errors,
      commits,
      value: await book.getSheets()[0].range('A1').getValues(),
      state: book.persistenceState.state,
    };
  });
  expect(result.after).toBe(result.revision + 1);
  expect(result.errors).toEqual(['commit observer failed']);
  expect(result.commits).toEqual([result.after]);
  expect(result.value).toEqual([[1234]]);
  expect(result.state).toBe('saved');
  expect(uncaught).toEqual([]);
});

test('asynchronous observer failures do not become unhandled rejections', async ({ page }) => {
  const uncaught: string[] = [];
  page.on('pageerror', (error) => uncaught.push(error.message));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook();
    const errors: string[] = [];
    let delivered = 0;
    (app as any).options.onError = async (error: Error) => {
      errors.push(error.message);
      throw new Error('async host failure');
    };
    const stopBad = app.on('workbook:committed', async () => {
      throw new Error('async listener failure');
    });
    const stopGood = app.on('workbook:committed', () => {
      delivered++;
    });
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[5678]]);
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    stopBad();
    stopGood();
    return { errors, delivered, value: await book.getSheets()[0].range('A1').getValues() };
  });
  expect(result).toEqual({ errors: ['async listener failure'], delivered: 1, value: [[5678]] });
  expect(uncaught).toEqual([]);
});
