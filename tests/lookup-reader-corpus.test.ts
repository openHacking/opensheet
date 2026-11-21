import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { evaluateFormula, parseFormula, type FormulaValue } from '@opensheetjs/formula';

const fixtureBytes = readFileSync(new URL('./fixtures/lookup-error-corpus.json', import.meta.url));
const fixture = JSON.parse(fixtureBytes.toString('utf8')) as {
  inputs: (FormulaValue | { formula: string })[];
  results: number[];
  cases: { formula: string; expected: FormulaValue }[];
};
const evidence = JSON.parse(
  readFileSync(
    new URL('../benchmarks/lookup-error-reader-2026-11-22.json', import.meta.url),
    'utf8',
  ),
);
const inputs: FormulaValue[] = [1, { error: '#DIV/0!' }, 3, null, null, ''];

it('matches the documented lookup profile in the pure evaluator', () => {
  for (const item of fixture.cases) {
    expect(
      evaluateFormula(parseFormula(item.formula), (ref) =>
        ref.column === 0 ? (inputs[ref.row] ?? null) : (fixture.results[ref.row] ?? null),
      ),
      item.formula,
    ).toEqual(item.expected);
  }
});

it('pins the independent reader observations and all workbook bytes', () => {
  expect(createHash('sha256').update(fixtureBytes).digest('hex')).toBe(evidence.fixture.sha256);
  expect(fixture.inputs).toEqual([1, { error: '#DIV/0!' }, 3, null, null, { formula: '""' }]);
  const source = XLSX.read(
    readFileSync(new URL(`../benchmarks/${evidence.artifacts.source.file}`, import.meta.url)),
    { type: 'buffer', sheetStubs: true },
  ).Sheets.Lookup;
  expect(source.A1.v).toBe(1);
  expect(source.A2.w).toBe('#DIV/0!');
  expect(source.A3.v).toBe(3);
  expect(source.A4).toBeUndefined();
  expect(source.A5).toBeUndefined();
  expect(source.A6.f).toBe('""');
  for (const [index, item] of fixture.cases.entries())
    expect(source[`D${index + 1}`].f).toBe(item.formula);
  for (const [index, value] of fixture.results.entries())
    expect(source[`B${index + 1}`].v).toBe(value);
  for (const artifact of Object.values(evidence.artifacts) as { file: string; sha256: string }[]) {
    const bytes = readFileSync(new URL(`../benchmarks/${artifact.file}`, import.meta.url));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
  }
  const saved = XLSX.read(
    readFileSync(new URL(`../benchmarks/${evidence.artifacts.saved.file}`, import.meta.url)),
    { type: 'buffer' },
  ).Sheets.Lookup;
  expect(evidence.observations).toHaveLength(fixture.cases.length);
  for (const [index, item] of fixture.cases.entries()) {
    const observed = evidence.observations[index];
    const cell = saved[`D${index + 1}`];
    const actual = cell.t === 'e' ? { error: cell.w } : cell.v;
    expect(observed).toMatchObject({
      ...item,
      reader: actual,
      agrees: JSON.stringify(item.expected) === JSON.stringify(actual),
    });
  }
  expect(evidence.observations.filter((item: { agrees: boolean }) => !item.agrees)).toHaveLength(6);
});

it.each([0, 65536])(
  'evaluates lookup cases through Worker and reopen with cache %i',
  async (cacheBytes) => {
    let book = await createWorkbook(
      { sheets: [{ name: 'Lookup', rows: 20, columns: 5 }] },
      { cacheBytes },
    );
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1:A6').setValues([[1], [null], [3], [null], [null], [null]]);
      await book.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [{ row: 1, column: 0, input: { type: 'error', code: '#DIV/0!' } }],
        },
      });
      await sheet.range('A6').setFormulas([['""']]);
      await sheet.range('B1:B6').setValues(fixture.results.map((value) => [value]));
      await sheet
        .range(`D1:D${fixture.cases.length}`)
        .setFormulas(fixture.cases.map((item) => [item.formula]));
      const expected = fixture.cases.map((item) => [item.expected]);
      expect(await sheet.range(`D1:D${fixture.cases.length}`).getValues()).toEqual(expected);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes });
      sheet = book.getSheets()[0];
      expect(await sheet.range(`D1:D${fixture.cases.length}`).getValues()).toEqual(expected);
    } finally {
      await book.deleteStorage();
    }
  },
);
