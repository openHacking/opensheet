import { expect, test } from '@playwright/test';

test('a throwing progress callback does not block other observers or generation', async ({
  page,
}) => {
  const pageErrors: string[] = [];
  page.on('pageerror', (error) => pageErrors.push(error.message));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    let optionCalls = 0,
      listenerCalls = 0;
    (book as any).options.onProgress = () => {
      optionCalls++;
      throw new Error('Injected progress observer');
    };
    const stop = book.onProgress(() => {
      listenerCalls++;
    });
    const revision = book.revision;
    await book.generate(1000);
    stop();
    return {
      optionCalls,
      listenerCalls,
      revision,
      after: book.revision,
      state: book.persistenceState.state,
    };
  });
  expect(result.optionCalls).toBeGreaterThan(0);
  expect(result.listenerCalls).toBe(result.optionCalls);
  expect(result.after).toBe(result.revision + 1);
  expect(result.state).toBe('saved');
  expect(pageErrors).toEqual([]);
});

test('throwing and rejecting workbook observers do not disrupt prefetch or commits', async ({
  page,
}) => {
  const pageErrors: string[] = [];
  page.on('pageerror', (error) => pageErrors.push(error.message));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    let data = 0,
      commits = 0,
      persistence = 0,
      progress = 0;
    (book as any).options.onProgress = async () => {
      throw new Error('progress option');
    };
    const stops = [
      book.onData(() => {
        throw new Error('data sync');
      }),
      book.onData(async () => {
        throw new Error('data async');
      }),
      book.onData(() => {
        data++;
      }),
      book.onCommit(async () => {
        throw new Error('commit async');
      }),
      book.onCommit(() => {
        commits++;
      }),
      book.onPersistence(async () => {
        throw new Error('persistence async');
      }),
      book.onPersistence(() => {
        persistence++;
      }),
      book.onProgress(async () => {
        throw new Error('progress async');
      }),
      book.onProgress(() => {
        progress++;
      }),
    ];
    await sheet.range('A65').setValues([[4321]]);
    await book.prefetch(sheet.id, sheet.range('A65').bounds);
    const value = await sheet.range('A65').getValues();
    await book.generate(1000);
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    stops.forEach((stop) => stop());
    return { data, commits, persistence, progress, value, state: book.persistenceState.state };
  });
  expect(result.data).toBeGreaterThan(0);
  expect(result.commits).toBeGreaterThan(0);
  expect(result.persistence).toBeGreaterThan(0);
  expect(result.progress).toBeGreaterThan(0);
  expect(result.value).toEqual([[4321]]);
  expect(result.state).toBe('saved');
  expect(pageErrors).toEqual([]);
});
