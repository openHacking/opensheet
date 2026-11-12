import { expect, test } from '@playwright/test';

test('caller mutation of a returned commit does not poison retry results or cache accounting', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    const command = {
      protocolVersion: 1 as const,
      commandId: 'isolated-retry',
      workbookId: book.id,
      baseRevision: book.revision,
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 7 } }],
      },
    };
    const first = (await book.execute(command))!;
    const expected = structuredClone(first);
    const bytes = book.stats.viewportCacheBytes;
    first.revision = 987654;
    first.commands[0].payload = { altered: 'x'.repeat(1000000) };
    first.changedRanges.length = 0;
    const second = (await book.execute(command))!;
    const secondMatches = JSON.stringify(second) === JSON.stringify(expected);
    second.commands.length = 0;
    second.changedRanges[0].startRow = 999;
    const third = await book.execute(command);
    return {
      secondMatches,
      thirdMatches: JSON.stringify(third) === JSON.stringify(expected),
      revision: book.revision,
      expectedRevision: expected.revision,
      bytes,
      afterBytes: book.stats.viewportCacheBytes,
      values: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    secondMatches: true,
    thirdMatches: true,
    revision: result.expectedRevision,
    expectedRevision: result.expectedRevision,
    bytes: result.bytes,
    afterBytes: result.bytes,
    values: [[7]],
  });
});
