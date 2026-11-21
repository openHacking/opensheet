import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs';

it('reports independently observed lookup interoperability risks on import and export', () => {
  const bytes = readFileSync(
    new URL('../benchmarks/lookup-error-reader-2026-11-22.source.xlsx', import.meta.url),
  );
  const input = XLSX.read(bytes, { type: 'buffer', sheetStubs: true, bookFiles: true });
  const imported = fromSheetJS(input);
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'XLSX_LOOKUP_IMPORT_INTEROP',
      action: 'approximated',
      address: 'D1',
      count: 12,
    }),
  );
  expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
  const exported = toSheetJS(imported.file);
  expect(exported.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'XLSX_LOOKUP_EXPORT_INTEROP',
      action: 'approximated',
      address: 'D1',
      count: 12,
    }),
  );
  expect(exported.report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LOOKUP_IMPORT_INTEROP', count: 12 }),
  );
  expect(() => toSheetJS(imported.file, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('recognizes function calls rather than similar names or quoted text', () => {
  const expressions = [
    'MATCH(1,A1:A2,0)',
    '_xlfn.VLOOKUP(1,A1:B2,2,0)',
    'IF(TRUE,MATCH(1,A1:A2,0),0)',
    '"MATCH("',
    "'MATCH('!A1",
    'MISMATCH(1)',
    '汉MATCH(1)',
    '𝒜MATCH(1)',
    '\\MATCH(1)',
    'XLOOKUP(1,A1:A2,B1:B2)',
    '"VLOOKUP("',
  ];
  const sheet = XLSX.utils.aoa_to_sheet(expressions.map(() => [0]));
  for (const [index, expression] of expressions.entries())
    sheet[`A${index + 1}`] = { t: 'n', f: expression, v: 0 };
  const imported = fromSheetJS(XLSX.utils.book_new(sheet, 'Functions'));
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LOOKUP_IMPORT_INTEROP', count: 3, address: 'A1' }),
  );
  expect(toSheetJS(imported.file).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LOOKUP_EXPORT_INTEROP', count: 3, address: 'A1' }),
  );
});
