import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('column edits around a filter and merge follow an independent model through history', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Columns', rows: 4, columns: 6 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    type Model = { rows: (number | null)[][]; key: number; start: number; end: number };
    const app = window.opensheet;
    for (const seed of [37, 819, 4231]) {
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      let model: Model = {
        rows: Array.from({ length: 4 }, (_, r) =>
          Array.from({ length: 6 }, (_, c) =>
            r === 3 && c >= 3 && c < 5 ? null : seed + r * 10 + c,
          ),
        ),
        key: 1,
        start: 3,
        end: 5,
      };
      await sheet
        .range({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 6 })
        .setValues(model.rows);
      const filter = (key: number) => ({
        mode: 'all' as const,
        headerRows: 0,
        criteria: [{ column: key, operator: 'gt' as const, value: seed + 15 }],
      });
      await sheet.setFilter(filter(model.key));
      await sheet.range({ startRow: 3, endRow: 4, startColumn: 3, endColumn: 5 }).merge();
      const clone = (value: Model): Model => ({
        ...value,
        rows: value.rows.map((row) => [...row]),
      });
      const verify = async (label: string) => {
        const actual = await sheet
          .range({ startRow: 0, endRow: 4, startColumn: 0, endColumn: model.rows[0].length })
          .getValues();
        const metadata = book.sheetData(sheet.id);
        const expectedHidden = model.rows.flatMap((row, i) =>
          row[model.key] === null || row[model.key]! <= seed + 15 ? [i] : [],
        );
        const expectedMerge = [
          { startRow: 3, startColumn: model.start, endRow: 4, endColumn: model.end },
        ];
        if (
          JSON.stringify(actual) !== JSON.stringify(model.rows) ||
          JSON.stringify(book.filteredRows(sheet.id)) !== JSON.stringify(expectedHidden) ||
          JSON.stringify(metadata.filter) !== JSON.stringify(filter(model.key)) ||
          JSON.stringify(metadata.merges) !== JSON.stringify(expectedMerge) ||
          sheet.columnCount !== model.rows[0].length
        )
          throw new Error(
            JSON.stringify({
              seed,
              label,
              model,
              actual,
              hidden: book.filteredRows(sheet.id),
              filter: metadata.filter,
              merges: metadata.merges,
            }),
          );
      };
      await verify('initial');
      const positions = [
        () => model.key,
        () => model.key + 1,
        () => model.start,
        () => model.end,
        () => model.rows[0].length,
      ];
      const inserted: number[] = [];
      for (const position of positions) {
        const index = position();
        const before = clone(model);
        await sheet.insertColumns(index);
        for (const row of model.rows) row.splice(index, 0, null);
        if (index <= model.key) model.key++;
        if (index <= model.start) {
          model.start++;
          model.end++;
        }
        inserted.push(index);
        await verify(`insert ${index}`);
        if (!(await book.undo())) throw new Error(`seed ${seed}: undo failed`);
        const after = model;
        model = before;
        await verify(`undo insert ${index}`);
        if (!(await book.redo())) throw new Error(`seed ${seed}: redo failed`);
        model = after;
        await verify(`redo insert ${index}`);
        await app.open(book.id);
        book = app.getWorkbook();
        sheet = book.getSheets()[0];
        await verify(`reopen insert ${index}`);
      }
      const revision = book.revision;
      for (const action of [
        () => sheet.insertColumns(model.start + 1),
        () => sheet.deleteColumns(model.start),
      ]) {
        let rejected = false;
        try {
          await action();
        } catch {
          rejected = true;
        }
        if (!rejected || book.revision !== revision)
          throw new Error(`seed ${seed}: merge intersection was not rejected atomically`);
        await verify('rejected merge intersection');
      }
      // Rejected commands must not create history entries: undo still removes the
      // last successful append, and redo restores it.
      const afterRejections = clone(model);
      if (!(await book.undo())) throw new Error(`seed ${seed}: undo after rejection failed`);
      for (const row of model.rows) row.pop();
      await verify('undo after rejections');
      if (!(await book.redo())) throw new Error(`seed ${seed}: redo after rejection failed`);
      model = afterRejections;
      await verify('redo after rejections');
      for (const index of inserted.reverse()) {
        const before = clone(model);
        await sheet.deleteColumns(index);
        for (const row of model.rows) row.splice(index, 1);
        if (index < model.key) model.key--;
        if (index < model.start) {
          model.start--;
          model.end--;
        }
        await verify(`delete ${index}`);
        if (!(await book.undo())) throw new Error(`seed ${seed}: undo failed`);
        const after = model;
        model = before;
        await verify(`undo delete ${index}`);
        if (!(await book.redo())) throw new Error(`seed ${seed}: redo failed`);
        model = after;
        await verify(`redo delete ${index}`);
        await app.open(book.id);
        book = app.getWorkbook();
        sheet = book.getSheets()[0];
        await verify(`reopen delete ${index}`);
      }
    }
    return null;
  }, file);
  expect(result).toBeNull();
});
