import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('cancels only the requested search while preserving queued work', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      for (let start = 0; start < 30000; start += 10000)
        await sheet
          .range(`A${start + 1}:A${start + 10000}`)
          .setValues(Array.from({ length: 10000 }, () => ['needle']));
      const already = new AbortController();
      already.abort();
      const preAborted = await book.search(sheet.id, 'needle', -1, { signal: already.signal }).then(
        () => 'completed',
        (error) => error.code,
      );
      const active = new AbortController();
      const cancelledActive = book.search(sheet.id, 'needle', -1, { signal: active.signal }).then(
        () => 'completed',
        (error) => error.code,
      );
      active.abort();
      const activeResult = await cancelledActive;
      const first = book.search(sheet.id, 'needle');
      const write = sheet.range('B1').setValues([['preserved']]);
      const queued = new AbortController();
      const cancelledQueued = book.search(sheet.id, 'needle', -1, { signal: queued.signal }).then(
        () => 'completed',
        (error) => error.code,
      );
      queued.abort();
      const firstResult = await first;
      await write;
      const queuedResult = await cancelledQueued;
      const id = book.id;
      await app.open(id);
      book = app.getWorkbook();
      return {
        preAborted,
        activeResult,
        queuedResult,
        firstResult,
        reopened: await book.getSheets()[0].range('B1').getValues(),
      };
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Search', rows: 30000, columns: 26 }] })),
  );
  expect(result).toEqual({
    preAborted: 'ABORTED',
    activeResult: 'ABORTED',
    queuedResult: 'ABORTED',
    firstResult: { count: 30000, match: { row: 0, column: 0 } },
    reopened: [['preserved']],
  });
});
