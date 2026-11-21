import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

const evidence = JSON.parse(
  readFileSync('benchmarks/calendar-series-reader-2026-11-21.json', 'utf8'),
);
for (const system of ['1900', '1904'] as const) {
  test(`fills independently recalculated calendar cases in ${system}`, async ({ page }) => {
    const cases = evidence.cases.filter((entry: { system: string }) => entry.system === system);
    const file = {
      ...createWorkbookFile({ sheets: [{ name: 'Independent dates', rows: 100, columns: 26 }] }),
      dateSystem: system,
    };
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(
      async ({ fileJSON, cases }) => {
        const app = window.opensheet;
        await app.load(JSON.parse(fileJSON));
        let book = app.getWorkbook();
        let sheet = book.getSheets()[0];
        for (let column = 0; column < cases.length; column++) {
          const item = cases[column];
          const name = String.fromCharCode(65 + column);
          await sheet
            .range(`${name}1:${name}${item.values.length}`)
            .fillSeries({ start: item.start, step: item.step, calendar: item.unit });
        }
        const read = async () =>
          Promise.all(
            cases.map(async (item: { values: number[] }, column: number) => {
              const name = String.fromCharCode(65 + column);
              return (await sheet.range(`${name}1:${name}${item.values.length}`).getValues()).map(
                (row) => row[0],
              );
            }),
          );
        const before = await read();
        await book.undo();
        const undone = await read();
        await book.redo();
        await app.open(book.id);
        book = app.getWorkbook();
        sheet = book.getSheets()[0];
        return { before, undone, after: await read() };
      },
      { fileJSON: JSON.stringify(file), cases },
    );
    const expected = cases.map((item: { values: number[] }) => item.values);
    expect(result.before).toEqual(expected);
    expect(result.after).toEqual(expected);
    expect(result.undone.slice(0, -1)).toEqual(expected.slice(0, -1));
    expect(result.undone.at(-1)).toEqual(Array(expected.at(-1).length).fill(null));
  });
}
