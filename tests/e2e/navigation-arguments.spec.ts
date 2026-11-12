import { expect, test } from '@playwright/test';

test('navigation rejects malformed arguments locally and at the Worker boundary', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[42]]);
    const revision = book.revision,
      original = (book as any).rpc;
    let calls = 0;
    (book as any).rpc = function (...args: unknown[]) {
      calls++;
      return original.apply(this, args);
    };
    const code = async (fn: () => unknown) => {
      try {
        await fn();
        return 'success';
      } catch (error) {
        return (error as any).code ?? 'uncoded';
      }
    };
    const invalid = [
      await code(() => book.navigateData('', 0, 0, 'row', 1)),
      await code(() => book.navigateData(sheet.id, NaN, 0, 'row', 1)),
      await code(() => book.navigateData(sheet.id, -1, 0, 'row', 1)),
      await code(() => book.navigateData(sheet.id, 0, Number.MAX_SAFE_INTEGER + 1, 'row', 1)),
      await code(() => book.navigateData(sheet.id, 0, 0, 'bad' as any, 1)),
      await code(() => book.navigateData(sheet.id, 0, 0, 'row', 0 as any)),
      await code(() => book.navigateData(sheet.id, 0, 0, 'row', 1, null as any)),
      await code(() => book.navigateData(sheet.id, 0, 0, 'row', 1, { signal: {} as any })),
    ];
    const localCalls = calls;
    const direct = [];
    for (const args of [
      ['', 0, 0, 'row', 1],
      [sheet.id, -1, 0, 'row', 1],
      [sheet.id, 0, 0, 'row', 0],
      [sheet.id, sheet.rowCount, 0, 'row', 1],
    ])
      direct.push(await code(() => original.call(book, 'navigateData', args)));
    const valid = await book.navigateData(sheet.id, 0, 0, 'row', 1);
    (book as any).rpc = original;
    await app.open(book.id);
    const reopened = app.getWorkbook();
    return {
      invalid,
      localCalls,
      direct,
      valid,
      revision,
      after: reopened.revision,
      value: await reopened.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result.invalid).toEqual(Array(8).fill('INVALID_ARGUMENT'));
  expect(result.localCalls).toBe(0);
  expect(result.direct).toEqual(Array(4).fill('INVALID_ARGUMENT'));
  expect(result.valid.revision).toBe(result.revision);
  expect(result.after).toBe(result.revision);
  expect(result.value).toEqual([[42]]);
});
