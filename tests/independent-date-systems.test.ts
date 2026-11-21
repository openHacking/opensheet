import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { FileReader } from '@opensheetjs/core';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';

it.each([
  ['1900', [1463, 45351.520833333336, 73110]],
  ['1904', [1, 43889.520833333336, 71648]],
] as const)('imports independently authored %s dates and formulas', (system, serials) => {
  const root = new URL('../benchmarks/date-system-openpyxl-2026-11-22/', import.meta.url);
  const bytes = readFileSync(new URL(`dates-${system}.xlsx`, root));
  const evidence = JSON.parse(readFileSync(new URL(`dates-${system}.xlsx.json`, root), 'utf8'));
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sha256);
  expect(evidence.serials).toEqual(serials);
  expect(evidence.rawDate1904).toBe(system === '1904');
  expect(evidence.xlsxSerials[1]).not.toBe(serials[1]);
  expect(evidence.xlsxSerials[1]).toBeCloseTo(serials[1], 8);
  const manifest = JSON.parse(
    readFileSync(new URL('./fixtures/interop/manifest.json', import.meta.url), 'utf8'),
  );
  const entry = manifest.fixtures.find(
    (item: { id: string }) => item.id === `dates-${system}-openpyxl`,
  );
  expect(entry.sha256).toBe(evidence.sha256);
  expect(Object.values(entry.expectedValues)).toEqual(evidence.xlsxSerials);
  const parsed = XLSX.read(bytes, {
    type: 'buffer',
    cellNF: true,
    cellStyles: true,
    sheetStubs: true,
    bookFiles: true,
  });
  const result = fromSheetJS(parsed);
  expect(result.file.dateSystem).toBe(system);
  for (const issue of entry.expectedImportIssues)
    expect(result.report.issues).toContainEqual(expect.objectContaining(issue));
  expect(() => fromSheetJS(parsed, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
  const dates = fromSheetJS(
    XLSX.read(bytes, {
      type: 'buffer',
      cellNF: true,
      cellDates: true,
      sheetStubs: true,
      bookFiles: true,
    }),
  ).file;
  expect(dates.dateSystem).toBe(system);
  const dateReader = new FileReader(dates);
  try {
    const dateValues = dateReader.getSheetById(dates.sheets[0].id)!.range('A1:A3').getValues();
    for (const [index, [value]] of dateValues.entries())
      expect(value).toBeCloseTo(evidence.xlsxSerials[index], 8);
  } finally {
    dateReader.dispose();
  }
  const reader = new FileReader(result.file);
  try {
    const sheet = reader.getSheetById(result.file.sheets[0].id)!;
    expect(sheet.range('A1:A3').getValues()).toEqual(
      evidence.xlsxSerials.map((value: number) => [value]),
    );
    expect(reader.getCell(result.file.sheets[0].id, 0, 1)?.input).toEqual({
      type: 'formula',
      expression: 'A2-A1',
    });
    expect(sheet.range('B1').getValues()[0][0]).toBeCloseTo(serials[1] - serials[0], 8);
    const cells = [0, 1, 2].map((row) => reader.getCell(result.file.sheets[0].id, row, 0)!);
    expect(
      cells.map((cell) => cell.numberFormat ?? reader.getStyle(cell.styleId).numberFormat),
    ).toEqual(['yyyy-mm-dd', 'yyyy-mm-dd hh:mm', 'yyyy-mm-dd']);
  } finally {
    reader.dispose();
  }
});
