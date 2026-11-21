import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('generated column changes preserve formula references, filters and merges', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const input = JSON.stringify(
    createWorkbookFile({
      sheets: [
        { name: 'Data', rows: 4, columns: 8 },
        { name: 'Other', rows: 4, columns: 3 },
      ],
    }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    type Model = { tokens: string[]; ids: number[]; next: number };
    const app = window.opensheet;
    for (const seed of [103, 3011, 9199]) {
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook(),
        [sheet, other] = book.getSheets();
      await sheet.range('A1:G3').setValues([
        [2, 3, 5, 4, 5, 6, 7],
        [4, 5, 6, 7, 8, 9, 10],
        [7, 8, 9, 10, 11, 12, 13],
      ]);
      await sheet.range('H1').setFormulas([['$A1+B$1+$C$1']]);
      await other.range('A1').setFormulas([['Data!$A1+Data!B$1+Data!$C$1']]);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 2, operator: 'gt', value: 5 }],
      });
      await sheet.range('E4:F4').merge();
      let model: Model = {
        tokens: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
        ids: [0, 1, 2, 3, 4, 5, 6, 7],
        next: 8,
      };
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      const clone = (value: Model): Model => ({
        tokens: [...value.tokens],
        ids: [...value.ids],
        next: value.next,
      });
      const trace: string[] = [],
        covered = new Set<string>();
      const verify = async () => {
        const column = (token: string) => model.tokens.indexOf(token);
        const letter = (index: number) => String.fromCharCode(65 + index);
        const expectedRefs = [
          `$${letter(column('A'))}1`,
          `${letter(column('B'))}$1`,
          `$${letter(column('C'))}$1`,
        ];
        const local = sheet.range({
          startRow: 0,
          endRow: 1,
          startColumn: column('H'),
          endColumn: column('H') + 1,
        });
        const formulas = [
          (await local.getFormulas())[0][0],
          (await other.range('A1').getFormulas())[0][0],
        ];
        const values = [
          (await local.getValues())[0][0],
          (await other.range('A1').getValues())[0][0],
        ];
        const matrix = await sheet
          .range({ startRow: 0, endRow: 4, startColumn: 0, endColumn: model.tokens.length })
          .getValues();
        const expectedMatrix = Array.from({ length: 4 }, (_, row) =>
          model.tokens.map((token) => {
            if (row === 3 || token.startsWith('new')) return null;
            if (token === 'H') return row === 0 ? 10 : null;
            return row === 0
              ? [2, 3, 5, 4, 5, 6, 7][token.charCodeAt(0) - 65]
              : row * 3 + token.charCodeAt(0) - 64;
          }),
        );
        const meta = book.sheetData(sheet.id);
        const filter = {
          mode: 'all',
          headerRows: 0,
          criteria: [{ column: column('C'), operator: 'gt', value: 5 }],
        };
        const merges = [
          { startRow: 3, startColumn: column('E'), endRow: 4, endColumn: column('F') + 1 },
        ];
        if (
          formulas.some(
            (formula) =>
              JSON.stringify(formula?.match(/\$?[A-Z]+\$?1/g)) !== JSON.stringify(expectedRefs),
          ) ||
          formulas[0]?.replace(/[()]/g, '').replace(/(?:'Data'|Data)!/g, '') !==
            expectedRefs.join('+') ||
          formulas[1]?.replace(/[()]/g, '').replace(/(?:'Data'|Data)!/g, '') !==
            expectedRefs.join('+') ||
          (formulas[1]?.match(/(?:'Data'|Data)!/g)?.length ?? 0) !== 3 ||
          JSON.stringify(values) !== JSON.stringify([10, 10]) ||
          JSON.stringify(matrix) !== JSON.stringify(expectedMatrix) ||
          JSON.stringify(book.filteredRows(sheet.id)) !== JSON.stringify([0, 3]) ||
          JSON.stringify(meta.filter) !== JSON.stringify(filter) ||
          JSON.stringify(meta.merges) !== JSON.stringify(merges) ||
          JSON.stringify(meta.columnOrder) !== JSON.stringify(model.ids) ||
          meta.nextColumnId !== model.next
        )
          throw new Error(
            JSON.stringify({
              seed,
              trace,
              model,
              formulas,
              expectedRefs,
              values,
              matrix,
              expectedMatrix,
              hidden: book.filteredRows(sheet.id),
              filter: meta.filter,
              merges: meta.merges,
              ids: meta.columnOrder,
              next: meta.nextColumnId,
            }),
          );
      };
      await verify();
      for (let step = 0; step < 30; step++) {
        const start = model.tokens.indexOf('E'),
          end = model.tokens.indexOf('F') + 1;
        const inserted = model.tokens.flatMap((token, index) =>
          token.startsWith('new') ? [index] : [],
        );
        const insert =
          step < 5 || !inserted.length || (model.tokens.length < 16 && random(2) === 0);
        const allowed = Array.from({ length: model.tokens.length + 1 }, (_, index) => index).filter(
          (index) => index <= start || index >= end,
        );
        const scripted = [0, model.tokens.indexOf('C'), model.tokens.indexOf('C') + 1, start, end][
          step
        ];
        const index = insert
          ? step < 5
            ? scripted
            : allowed[random(allowed.length)]
          : inserted[random(inserted.length)];
        const before = clone(model);
        if (insert) {
          covered.add(
            index === 0
              ? 'left'
              : index === start
                ? 'merge-start'
                : index === end
                  ? 'merge-end'
                  : index < start
                    ? 'before-merge'
                    : 'after-merge',
          );
          await sheet.insertColumns(index);
          model.tokens.splice(index, 0, `new${model.next}`);
          model.ids.splice(index, 0, model.next++);
        } else {
          covered.add('delete-inserted');
          await sheet.deleteColumns(index);
          model.tokens.splice(index, 1);
          model.ids.splice(index, 1);
        }
        trace.push(`${insert ? 'insert' : 'delete'}:${index}`);
        await verify();
        const after = clone(model);
        if (!(await book.undo())) throw new Error(`seed ${seed}: undo failed`);
        model = before;
        await verify();
        if (!(await book.redo())) throw new Error(`seed ${seed}: redo failed`);
        model = after;
        await verify();
        if (step % 6 === 5) {
          await app.open(book.id);
          book = app.getWorkbook();
          [sheet, other] = book.getSheets();
          await verify();
        }
      }
      for (const feature of [
        'left',
        'merge-start',
        'merge-end',
        'before-merge',
        'after-merge',
        'delete-inserted',
      ])
        if (!covered.has(feature)) throw new Error(JSON.stringify({ seed, feature, trace }));
    }
    return 'ok';
  }, input);
  expect(result).toBe('ok');
});
