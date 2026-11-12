import { expect, test } from '@playwright/test';

test('transaction lookups reflect renamed and removed sheets without changing the parent', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(
      createWorkbookFile({
        sheets: [
          { name: 'First', rows: 2, columns: 2 },
          { name: 'Second', rows: 2, columns: 2 },
        ],
      }),
    );
    let book = app.getWorkbook();
    const [first, second] = book.getSheets();
    let inside: any;
    await book.transaction({}, async (tx) => {
      await tx.renameSheet(first.id, 'Renamed');
      await tx.removeSheet(second.id);
      const replacement = await tx.addSheet('Second', { rows: 1, columns: 1 });
      await tx
        .getSheetByName('renamed')!
        .range('A1')
        .setValues([[12]]);
      await tx
        .getSheetByName('Second')!
        .range('A1')
        .setValues([[34]]);
      inside = {
        names: tx.getSheets().map((sheet) => sheet.name),
        renamedId: tx.getSheetByName('Renamed')?.id,
        removedId: tx.getSheetById(second.id)?.id ?? null,
        replacementId: replacement.id,
        lookupId: tx.getSheetByName('second')?.id,
        parent: book.getSheets().map((sheet) => sheet.name),
        committedName: tx.sheetData(first.id).name,
      };
    });
    const read = async () => ({
      names: book.getSheets().map((sheet) => sheet.name),
      values: await Promise.all(book.getSheets().map((sheet) => sheet.range('A1').getValues())),
    });
    const committed = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return {
      firstId: first.id,
      secondId: second.id,
      inside,
      committed,
      undone,
      reopened: await read(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.inside.names).toEqual(['Renamed', 'Second']);
  expect(result.inside.renamedId).toBe(result.firstId);
  expect(result.inside.removedId).toBeNull();
  expect(result.inside.lookupId).toBe(result.inside.replacementId);
  expect(result.inside.replacementId).not.toBe(result.secondId);
  expect(result.inside.parent).toEqual(['First', 'Second']);
  expect(result.inside.committedName).toBe('First');
  expect(result.committed).toEqual({ names: ['Renamed', 'Second'], values: [[[12]], [[34]]] });
  expect(result.undone).toEqual({ names: ['First', 'Second'], values: [[[null]], [[null]]] });
  expect(result.reopened).toEqual(result.committed);
});

test('a removed sheet ID can be re-added with new staged dimensions', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(
      createWorkbookFile({
        sheets: [
          { name: 'Keep', rows: 2, columns: 2 },
          { name: 'Old', rows: 3, columns: 2 },
        ],
      }),
    );
    let book = app.getWorkbook();
    const old = book.getSheetByName('Old')!;
    await old.range('A1').setValues([[8]]);
    const replacement = structuredClone(book.sheetData(old.id));
    replacement.name = 'New';
    replacement.rowOrder = [0];
    replacement.columnOrder = [0];
    replacement.nextRowId = 1;
    replacement.nextColumnId = 1;
    let inside: any;
    await book.transaction({}, async (tx) => {
      await tx.removeSheet(old.id);
      await tx.execute({ type: 'core.sheet.add', payload: { sheet: replacement } });
      const sheet = tx.getSheetById(old.id)!;
      inside = {
        names: tx.getSheets().map((item) => item.name),
        dimensions: [sheet.rowCount, sheet.columnCount],
        committed: await sheet.range('A1').getValues(),
      };
      await sheet.range('A1').setValues([[15]]);
    });
    const read = async () => ({
      names: book.getSheets().map((item) => item.name),
      value: await book.getSheetById(old.id)!.range('A1').getValues(),
    });
    const committed = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { inside, committed, undone, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.inside).toEqual({ names: ['Keep', 'New'], dimensions: [1, 1], committed: [[8]] });
  expect(result.committed).toEqual({ names: ['Keep', 'New'], value: [[15]] });
  expect(result.undone).toEqual({ names: ['Keep', 'Old'], value: [[8]] });
  expect(result.reopened).toEqual(result.committed);
});
