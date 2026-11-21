import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('seeded formula filtering and row edits match an independent model through history', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Formula model', rows: 6, columns: 5 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    type Row = { value: number | null; formula: boolean; tag: number | null };
    type Model = { rows: Row[]; filter: boolean };
    const app = window.opensheet;
    for (const seed of [43, 419, 8161]) {
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      let model: Model = {
        rows: [
          ...[3, 8, 3, 12, 1].map(
            (value, index): Row => ({ value, formula: true, tag: 101 + index }),
          ),
          { value: null, formula: false, tag: null },
        ],
        filter: false,
      };
      await sheet.range('A1:A5').setValues(model.rows.slice(0, 5).map((row) => [row.value]));
      await sheet
        .range('B1:B5')
        .setFormulas(model.rows.slice(0, 5).map((_, index) => [`A${index + 1}*2`]));
      await sheet.range('C1:C5').setValues(model.rows.slice(0, 5).map((row) => [row.tag]));
      await sheet.range('D6:E6').merge();
      const clone = (m: Model): Model => ({
        filter: m.filter,
        rows: m.rows.map((row) => ({ ...row })),
      });
      const past: Model[] = [],
        future: Model[] = [],
        trace: string[] = [];
      const covered = new Set<string>();
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      const save = () => {
        past.push(clone(model));
        future.length = 0;
      };
      const verify = async (step: number) => {
        const count = model.rows.length;
        const values = await sheet
          .range({ startRow: 0, endRow: count, startColumn: 0, endColumn: 5 })
          .getValues();
        const formulas = await sheet
          .range({ startRow: 0, endRow: count, startColumn: 1, endColumn: 2 })
          .getFormulas();
        const expectedValues = model.rows.map((row) => [
          row.value,
          row.formula ? row.value! * 2 : null,
          row.tag,
          null,
          null,
        ]);
        const expectedFormulas = model.rows.map((row, index) => [
          row.formula ? `A${index + 1}*2` : null,
        ]);
        // Structural transforms may add grouping and an explicit current-sheet qualifier.
        const normalizedFormulas = formulas.map(([expression]) => [
          expression === null
            ? null
            : expression.replace(/[()]/g, '').replaceAll("'Formula model'!", ''),
        ]);
        const hidden = model.filter
          ? model.rows.flatMap((row, index) => (row.formula && row.value! * 2 > 15 ? [] : [index]))
          : [];
        const filter = model.filter
          ? { mode: 'all', headerRows: 0, criteria: [{ column: 1, operator: 'gt', value: 15 }] }
          : null;
        const meta = book.sheetData(sheet.id);
        const merges = [{ startRow: count - 1, startColumn: 3, endRow: count, endColumn: 5 }];
        if (
          JSON.stringify(values) !== JSON.stringify(expectedValues) ||
          JSON.stringify(normalizedFormulas) !== JSON.stringify(expectedFormulas) ||
          JSON.stringify(book.filteredRows(sheet.id)) !== JSON.stringify(hidden) ||
          JSON.stringify(meta.filter) !== JSON.stringify(filter) ||
          JSON.stringify(meta.merges) !== JSON.stringify(merges) ||
          sheet.rowCount !== count
        )
          throw Error(
            JSON.stringify({
              seed,
              step,
              trace,
              model,
              values,
              expectedValues,
              formulas,
              expectedFormulas,
              hidden: book.filteredRows(sheet.id),
              expectedHidden: hidden,
              filter: meta.filter,
              merges: meta.merges,
            }),
          );
      };
      for (let step = 0; step < 42; step++) {
        const action = step < 8 ? [4, 3, 0, 1, 2, 3, 5, 6][step] : random(7);
        const last = model.rows.length - 1;
        if (action === 0) {
          const row = random(last),
            value = seed * 100 + step;
          save();
          await sheet
            .range({ startRow: row, endRow: row + 1, startColumn: 0, endColumn: 1 })
            .setValues([[value]]);
          model.rows[row].value = value;
          trace.push(`edit:${row}`);
          covered.add('edit');
        } else if (action === 1 && last < 9) {
          const row = random(last + 1);
          save();
          await sheet.insertRows(row);
          model.rows.splice(row, 0, { value: null, formula: false, tag: null });
          trace.push(`insert:${row}`);
          covered.add('insert');
        } else if (action === 2 && last > 2) {
          const row = random(last);
          save();
          await sheet.deleteRows(row);
          model.rows.splice(row, 1);
          trace.push(`delete:${row}`);
          covered.add('delete');
        } else if (action === 3) {
          const keys = model.rows
            .slice(0, last)
            .filter((row) => row.formula)
            .map((row) => row.value);
          if (new Set(keys).size < keys.length) covered.add('sort-tie');
          const direction = random(2) ? 'asc' : 'desc';
          save();
          await sheet.sort(
            { startRow: 0, endRow: last, startColumn: 0, endColumn: 3 },
            1,
            direction,
          );
          const ordered = model.rows
            .slice(0, last)
            .map((row, index) => ({ row, index }))
            .sort((a, b) => {
              if (!a.row.formula && !b.row.formula) return a.index - b.index;
              if (!a.row.formula) return 1;
              if (!b.row.formula) return -1;
              return (
                (direction === 'asc' ? a.row.value! - b.row.value! : b.row.value! - a.row.value!) ||
                a.index - b.index
              );
            })
            .map(({ row }) => row);
          model.rows = [...ordered, model.rows[last]];
          trace.push(`sort:${direction}`);
          covered.add('sort');
        } else if (action === 4) {
          save();
          model.filter = !model.filter;
          await sheet.setFilter(
            model.filter
              ? { mode: 'all', headerRows: 0, criteria: [{ column: 1, operator: 'gt', value: 15 }] }
              : null,
          );
          trace.push(model.filter ? 'filter' : 'clear filter');
          covered.add('filter');
        } else if (action === 5 && past.length) {
          future.push(clone(model));
          model = past.pop()!;
          if (!(await book.undo())) throw Error('undo failed');
          trace.push('undo');
          covered.add('undo');
        } else if (action === 6 && future.length) {
          past.push(clone(model));
          model = future.pop()!;
          if (!(await book.redo())) throw Error('redo failed');
          trace.push('redo');
          covered.add('redo');
        } else trace.push('skip');
        if (model.filter && step % 8 === 7) {
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          trace.push('reopen');
          covered.add('reopen-filter');
        }
        await verify(step);
      }
      for (const action of [
        'edit',
        'insert',
        'delete',
        'sort',
        'sort-tie',
        'filter',
        'undo',
        'redo',
        'reopen-filter',
      ])
        if (!covered.has(action)) throw Error(JSON.stringify({ seed, trace, missing: action }));
    }
    return 'ok';
  }, file);
  expect(result).toBe('ok');
});
