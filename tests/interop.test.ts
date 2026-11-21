import { decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { readFileSync } from 'node:fs';
import { describe, it, expect, vi } from 'vitest';
import * as XLSX from 'xlsx';
import {
  fromSheetJS,
  toSheetJS,
  writeXlsx,
  unsupportedXlsxSheetName,
} from '@opensheetjs/adapter-sheetjs';
import { createWorkbook, FileReader, Workbook } from '@opensheetjs/core';
import { exportRange, parseDelimited, stringifyDelimited } from '@opensheetjs/formats';
import { FormulaCalculation } from '../packages/core/src/calculation.js';
import { mayDifferInPowerPrecedence } from '../packages/adapter-sheetjs/src/formula.js';
import { buildSheetJS } from '../packages/adapter-sheetjs/src/export.js';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { parseSpreadsheetFormula } from '../packages/formula/src/parser.js';

it('identifies the pinned XLSX writer sheet-name restrictions without rejecting valid internal apostrophes', () => {
  for (const name of [
    '',
    'x'.repeat(32),
    "'prefix",
    "suffix'",
    'HiStOrY',
    'bad:name',
    'a/b',
    'a\\b',
    'a?b',
    'a*b',
    'a[b]',
    null as any,
  ])
    expect(unsupportedXlsxSheetName(name), String(name)).toBe(true);
  for (const name of ['Sheet', "O'Brien", 'History 2', 'x'.repeat(31)])
    expect(unsupportedXlsxSheetName(name), name).toBe(false);
});

it('counts independent workbook- and sheet-scoped defined names as dropped', () => {
  const source = XLSX.read(readFileSync('benchmarks/xlsx-defined-names-openpyxl-2026-11-12.xlsx'), {
    type: 'buffer',
    bookFiles: true,
  });
  const imported = fromSheetJS(source);
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({ code: 'NAMES_UNSUPPORTED', action: 'dropped', count: 3 }),
  );
  expect(imported.report.summary.dropped).toBe(4);
  expect(toSheetJS(imported.file).report.issues).toContainEqual(
    expect.objectContaining({ code: 'NAMES_UNSUPPORTED', action: 'dropped', count: 3 }),
  );
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
  for (const names of ['invalid', { length: 2 }, 5, [null], [{}], [{ Name: '' }], new Array(1)]) {
    const invalid = { ...source, Workbook: { ...source.Workbook, Names: names } };
    expect(() => fromSheetJS(invalid)).toThrow('Invalid defined-name collection');
  }
});

it('reports formulas that reference independently authored defined names without flagging quoted text', () => {
  const source = XLSX.read(
    readFileSync('benchmarks/xlsx-defined-name-formulas-openpyxl-2026-11-12.xlsx'),
    {
      type: 'buffer',
      bookFiles: true,
      sheetStubs: true,
      cellStyles: true,
    },
  );
  const imported = fromSheetJS(source);
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'DEFINED_NAME_REFERENCE_DROPPED',
      action: 'dropped',
      address: 'B1',
      count: 2,
    }),
  );
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({ code: 'NAMES_UNSUPPORTED', count: 3 }),
  );
  const output = toSheetJS(imported.file);
  expect(output.report.issues).toContainEqual(
    expect.objectContaining({ code: 'DEFINED_NAME_REFERENCE_DROPPED', count: 2 }),
  );
});

it('limits local defined-name warnings to their sheet and discloses an incomplete index', () => {
  const source = {
    SheetNames: ['First', 'Second'],
    Sheets: {
      First: { '!ref': 'A1', A1: { t: 'n', f: 'LocalOnly+1', v: 0 } },
      Second: {
        '!ref': 'A1:A2',
        A1: { t: 'n', f: 'LocalOnly+1', v: 0 },
        A2: { t: 'n', f: 'GlobalName+1', v: 0 },
      },
    },
    Workbook: {
      Names: [{ Name: 'LocalOnly', Sheet: 0 }, { Name: 'GlobalName' }, { Name: 'not valid' }],
    },
  };
  const result = fromSheetJS(source);
  const references = result.report.issues.filter(
    (issue) => issue.code === 'DEFINED_NAME_REFERENCE_DROPPED',
  );
  expect(references.map((issue) => [issue.sheetId, issue.address, issue.count])).toEqual([
    [result.file.sheets[0].id, 'A1', 1],
    [result.file.sheets[1].id, 'A2', 1],
  ]);
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'DEFINED_NAME_SCAN_INCOMPLETE' }),
  );
});

it('reports an independently authored sheet-qualified local name but not external or cell references', () => {
  const source = XLSX.read(
    readFileSync('benchmarks/xlsx-qualified-names-openpyxl-2026-11-12.xlsx'),
    {
      type: 'buffer',
      bookFiles: true,
      sheetStubs: true,
      cellStyles: true,
    },
  );
  const result = fromSheetJS(source);
  const issues = result.report.issues.filter(
    (issue) => issue.code === 'DEFINED_NAME_REFERENCE_DROPPED',
  );
  expect(issues).toHaveLength(1);
  expect(issues[0]).toMatchObject({ sheetId: result.file.sheets[1].id, address: 'A1', count: 1 });
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
      address: 'A3',
    }),
  );
});

it('reports an independently authored sheet-qualified global name without flagging controls', () => {
  const source = XLSX.read(
    readFileSync('benchmarks/xlsx-qualified-global-openpyxl-2026-11-12.xlsx'),
    {
      type: 'buffer',
      bookFiles: true,
      sheetStubs: true,
      cellStyles: true,
    },
  );
  const result = fromSheetJS(source);
  const issues = result.report.issues.filter(
    (issue) => issue.code === 'DEFINED_NAME_REFERENCE_DROPPED',
  );
  expect(issues).toHaveLength(1);
  expect(issues[0]).toMatchObject({ sheetId: result.file.sheets[1].id, address: 'A1', count: 1 });
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
      address: 'A4',
    }),
  );
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('reports independently authored Unicode name references by sheet without flagging controls', () => {
  const source = XLSX.read(readFileSync('benchmarks/xlsx-unicode-names-openpyxl-2026-11-12.xlsx'), {
    type: 'buffer',
    bookFiles: true,
    sheetStubs: true,
    cellStyles: true,
  });
  const result = fromSheetJS(source);
  const issues = result.report.issues.filter(
    (issue) => issue.code === 'DEFINED_NAME_REFERENCE_DROPPED',
  );
  expect(issues.map(({ sheetId, address, count }) => [sheetId, address, count])).toEqual([
    [result.file.sheets[0].id, 'B1', 2],
    [result.file.sheets[1].id, 'A1', 2],
  ]);
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'DEFINED_NAME_SCAN_INCOMPLETE' }),
  );
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
      address: 'A4',
    }),
  );
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('reports independently authored backslash-prefixed names without flagging controls', () => {
  const source = XLSX.read(
    readFileSync('benchmarks/xlsx-backslash-names-openpyxl-2026-11-12.xlsx'),
    {
      type: 'buffer',
      bookFiles: true,
      sheetStubs: true,
      cellStyles: true,
    },
  );
  const result = fromSheetJS(source);
  const issues = result.report.issues.filter(
    (issue) => issue.code === 'DEFINED_NAME_REFERENCE_DROPPED',
  );
  expect(issues.map(({ sheetId, address, count }) => [sheetId, address, count])).toEqual([
    [result.file.sheets[0].id, 'B1', 2],
    [result.file.sheets[1].id, 'A1', 2],
  ]);
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'DEFINED_NAME_SCAN_INCOMPLETE' }),
  );
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
      address: 'A4',
    }),
  );
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('reports long serialized formulas at the observed reader boundary', () => {
  const formulas = [8191, 8192, 8193].map(
    (length) => (length % 2 ? '1' : '+1') + '+0'.repeat(Math.floor((length - 1) / 2)),
  );
  const source = {
    SheetNames: ['Length'],
    Sheets: {
      Length: {
        '!ref': 'A1:A3',
        A1: { t: 'n', f: formulas[0], v: 1 },
        A2: { t: 'n', f: formulas[1], v: 1 },
        A3: { t: 'n', f: formulas[2], v: 1 },
      },
    },
  };
  const file = fromSheetJS(source).file;
  const converted = toSheetJS(file);
  expect(converted.report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_FORMULA_LENGTH_INTEROP', count: 2, address: 'A2' }),
  );
  for (let i = 0; i < formulas.length; i++)
    expect(converted.workbook.Sheets.Length[`A${i + 1}`].f).toBe(formulas[i]);
  expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('reports and blocks XLSX text beyond the writer limit before invoking the writer', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Text', rows: 3, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet
      .range('A1:A3')
      .setValues([['a'.repeat(32767)], ['b'.repeat(32768)], ['😀'.repeat(16384)]]);
    const file = await book.toJSON();
    const report = toSheetJS(file).report;
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'XLSX_TEXT_LENGTH_BLOCKED',
        action: 'blocked',
        address: 'A2',
        count: 2,
      }),
    );
    expect(report.summary.blocked).toBe(2);
    const writer = { ...XLSX, write: vi.fn() };
    expect(() => writeXlsx(file, writer)).toThrow(
      expect.objectContaining({
        code: 'UNSUPPORTED_FEATURE',
        message: expect.stringContaining('First location: Text!A2'),
        details: report,
      }),
    );
    expect(() => writeXlsx(file, writer, { unsupported: 'strict' })).toThrow(
      expect.objectContaining({ code: 'UNSUPPORTED_FEATURE', details: report }),
    );
    expect(writer.write).not.toHaveBeenCalled();
    await sheet.range('A2:A3').clear({ all: true });
    const bytes = writeXlsx(await book.toJSON(), XLSX).bytes;
    expect(XLSX.read(bytes, { type: 'array' }).Sheets.Text.A1.v).toBe('a'.repeat(32767));
  } finally {
    await book.deleteStorage();
  }
});

it('discards retained worksheets after an XLSX text blocker without changing the report', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'First' }, { name: 'Second' }] });
  try {
    const [first, second] = book.getSheets();
    await first.range('A1:B1').setValues([[42, '_x0041_']]);
    await second.range('A1:B1').setValues([['x'.repeat(32768), 7]]);
    const file = await book.toJSON();
    const full = buildSheetJS(file, {}, true);
    const rejected = buildSheetJS(file, {}, true, false, true);
    expect(rejected.report).toEqual(full.report);
    expect(rejected.xlsxTextTooLong).toEqual({ sheetName: 'Second', address: 'A1' });
    expect(
      Object.values(rejected.workbook.Sheets).flatMap((sheet) => Object.keys(sheet)),
    ).not.toContain('A1');
    expect(full.workbook.Sheets.First.A1.v).toBe(42);
    expect(full.workbook.Sheets.Second.B1.v).toBe(7);
    await second.range('A1').setValues([['valid']]);
    const valid = buildSheetJS(await book.toJSON(), {}, true, false, true);
    expect(valid.xlsxTextTooLong).toBeUndefined();
    expect(valid.workbook.Sheets.First.A1.v).toBe(42);
    expect(valid.workbook.Sheets.Second.B1.v).toBe(7);
  } finally {
    await book.deleteStorage();
  }
});

it('reports an independently authored unfrozen active worksheet and selection', () => {
  const input = XLSX.read(readFileSync('benchmarks/xlsx-unfrozen-view-openpyxl-2026-11-12.xlsx'), {
    type: 'buffer',
    bookFiles: true,
  });
  const converted = fromSheetJS(input);
  expect(converted.file.sheets[0].freeze).toEqual({ rows: 0, columns: 0 });
  expect(converted.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_ACTIVE_TAB_DROPPED' }),
  );
  expect(converted.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_SELECTION_DROPPED' }),
  );
  expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('reports and blocks names rejected by the XLSX writer without renaming references', async () => {
  const book = await createWorkbook({
    sheets: [
      { name: "'Prefix" },
      { name: "Suffix'" },
      { name: 'HiStOrY' },
      { name: "O'Brien" },
      { name: 'History 2' },
    ],
  });
  try {
    const sheets = book.getSheets();
    await sheets[2].range('A1').setValues([[6]]);
    await sheets[3].range('A1').setFormulas([["'HiStOrY'!A1+1"]]);
    expect(await sheets[3].range('A1').getValues()).toEqual([[7]]);
    const file = await book.toJSON(),
      before = structuredClone(file);
    const converted = toSheetJS(file);
    const blocked = converted.report.issues.filter(
      (issue) => issue.code === 'XLSX_SHEET_NAME_BLOCKED',
    );
    expect(converted.report.summary.blocked).toBe(3);
    expect(blocked.map((issue) => issue.sheetId)).toEqual(
      sheets.slice(0, 3).map((sheet) => sheet.id),
    );
    expect(converted.workbook.SheetNames).toEqual(sheets.map((sheet) => sheet.name));
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const writer = {
      ...XLSX,
      write: vi.fn(() => {
        throw new Error('writer must not run');
      }),
    };
    let failure: unknown;
    try {
      writeXlsx(file, writer);
    } catch (error) {
      failure = error;
    }
    expect(failure).toMatchObject({
      code: 'UNSUPPORTED_FEATURE',
      message: expect.stringContaining('Rename sheets'),
      details: { summary: { blocked: 3 } },
    });
    expect(writer.write).not.toHaveBeenCalled();
    expect(file).toEqual(before);
    await book.renameSheet(sheets[0].id, 'Prefix');
    await book.renameSheet(sheets[1].id, 'Suffix');
    await book.renameSheet(sheets[2].id, 'Archive');
    expect(await sheets[3].range('A1').getValues()).toEqual([[7]]);
    const exported = writeXlsx(await book.toJSON(), XLSX);
    expect(exported.report.summary.blocked).toBe(0);
    const parsed = XLSX.read(exported.bytes, { type: 'array' });
    expect(parsed.SheetNames).toEqual(['Prefix', 'Suffix', 'Archive', "O'Brien", 'History 2']);
    expect(parsed.Sheets["O'Brien"].A1.f).toContain('Archive');
  } finally {
    await book.deleteStorage();
  }
});

it('blocks an unsupported sheet name alongside grouped cell warnings', async () => {
  const book = await createWorkbook({ sheets: [{ name: "'Blocked", rows: 201, columns: 1 }] });
  try {
    await book
      .getSheets()[0]
      .range('A1:A201')
      .setValues(Array.from({ length: 201 }, () => ['x\0y']));
    const file = await book.toJSON(),
      report = toSheetJS(file).report;
    expect(report.issues).toHaveLength(2);
    expect(report.issues.some((issue) => issue.code === 'XLSX_SHEET_NAME_BLOCKED')).toBe(true);
    expect(report.summary.approximated).toBe(201);
    expect(report.summary.blocked).toBe(1);
    expect(report.omittedIssueOccurrences).toBe(0);
    const writer = {
      ...XLSX,
      write: vi.fn(() => {
        throw new Error('writer must not run');
      }),
    };
    const expectedReport = buildSheetJS(file, {}, true).report;
    const reportOnly = buildSheetJS(file, {}, true, true);
    expect(reportOnly.report).toEqual(expectedReport);
    expect(
      Object.keys(reportOnly.workbook.Sheets["'Blocked"]).filter((key) => !key.startsWith('!')),
    ).toHaveLength(0);
    expect(() => writeXlsx(file, writer)).toThrow(
      expect.objectContaining({ code: 'UNSUPPORTED_FEATURE', details: expectedReport }),
    );
    expect(() => writeXlsx(file, writer, { unsupported: 'strict' })).toThrow(
      expect.objectContaining({
        code: 'UNSUPPORTED_FEATURE',
        message: 'Strict compatibility check failed',
        details: expectedReport,
      }),
    );
    expect(writer.write).not.toHaveBeenCalled();
  } finally {
    await book.deleteStorage();
  }
});
describe('SheetJS', () => {
  it('reports potentially different power precedence on import and export', () => {
    for (const expression of [
      '2^3^2',
      '-2^2',
      '+1^++1',
      '2^(-3^2)',
      '2^3e-2^2',
      '2^+3^2',
      'SUM(2^-3,1)',
    ])
      expect(mayDifferInPowerPrecedence(expression), expression).toBe(true);
    for (const expression of [
      '2^3',
      '2^(3^2)',
      'SUM(2^3,4^5)',
      '2^3+4^5',
      '1e-3+2^3',
      '"-2^2"&"2^3^2"',
      "'a^-b'!A1+2^3",
    ])
      expect(mayDifferInPowerPrecedence(expression), expression).toBe(false);
    const source = {
      SheetNames: ['Data'],
      Sheets: {
        Data: {
          A1: { t: 'n', f: '2^3^2', v: 64 },
          A2: { t: 'n', f: '-2^2', v: 4 },
          A3: { t: 'n', f: '2^(3^2)', v: 512 },
          A4: { t: 'n', f: '2^+3^2', v: 64 },
          '!ref': 'A1:A4',
        },
      },
    };
    const imported = fromSheetJS(source);
    expect(imported.report.issues).toContainEqual(
      expect.objectContaining({ code: 'IMPORT_FORMULA_PRECEDENCE_NORMALIZED', count: 3 }),
    );
    expect(() => fromSheetJS(source, { unsupported: 'strict' })).not.toThrow();
    const snapshot = decodeWorkbookFile(imported.file);
    for (const [index, expected] of [64, 4, 512, 64].entries()) {
      const input = snapshot.sheets[0].cells[`${index}:0`].input;
      expect(input.type).toBe('formula');
      if (input.type === 'formula')
        expect(evaluateFormula(parseFormula(input.expression), () => null)).toBe(expected);
    }
    const exported = toSheetJS(imported.file);
    expect(exported.report.issues).toContainEqual(
      expect.objectContaining({ code: 'EXPORT_FORMULA_PRECEDENCE_NORMALIZED' }),
    );
    for (const [index, expected] of [64, 4, 512, 64].entries())
      expect(
        evaluateFormula(
          parseSpreadsheetFormula(exported.workbook.Sheets.Data[`A${index + 1}`].f),
          () => null,
        ),
      ).toBe(expected);
    const unsupported = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', f: 'UNKNOWN(1)^-2' }, '!ref': 'A1' } },
    };
    const fallback = fromSheetJS(unsupported);
    expect(fallback.report.issues).toContainEqual(
      expect.objectContaining({ code: 'IMPORT_FORMULA_POWER_PRECEDENCE' }),
    );
    expect(() => fromSheetJS(unsupported, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
  it('preserves independently recalculated spreadsheet precedence on import', () => {
    const input = XLSX.read(
      readFileSync('benchmarks/formula-precedence-2026-11-11-libreoffice.xlsx'),
      {
        type: 'buffer',
      },
    );
    const { file } = fromSheetJS(input);
    const snapshot = decodeWorkbookFile(file);
    for (const [row, expected] of [64, 4, 512, 4].entries()) {
      const cell = snapshot.sheets[0].cells[`${row}:0`].input;
      expect(cell.type).toBe('formula');
      if (cell.type === 'formula')
        expect(evaluateFormula(parseFormula(cell.expression), () => null)).toBe(expected);
    }
  });
  it('preserves unary-plus precedence in an independently authored workbook', () => {
    const input = XLSX.read(readFileSync('benchmarks/plus-precedence-2026-11-11-source.xlsx'), {
      type: 'buffer',
      sheetStubs: true,
    });
    const { file } = fromSheetJS(input);
    const snapshot = decodeWorkbookFile(file);
    for (const [row, expected] of [64, 1].entries()) {
      const cell = snapshot.sheets[0].cells[`${row}:0`].input;
      expect(cell.type).toBe('formula');
      if (cell.type === 'formula')
        expect(evaluateFormula(parseFormula(cell.expression), () => null)).toBe(expected);
    }
    expect(snapshot.sheets[0].cells['0:0'].cached).toBeUndefined();
  });
  it('does not mistake a formula stub value for a cached calculation', () => {
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'z', f: 'UNKNOWN(1)', v: 0 }, '!ref': 'A1' } },
    };
    const cell = decodeWorkbookFile(fromSheetJS(source).file).sheets[0].cells['0:0'];
    expect(cell.input).toEqual({ type: 'formula', expression: 'UNKNOWN(1)' });
    expect(cell.cached).toBeUndefined();
    const file = XLSX.read(readFileSync('benchmarks/formula-uncached-openpyxl-2026-11-11.xlsx'), {
      type: 'buffer',
      sheetStubs: true,
    });
    const independent = decodeWorkbookFile(fromSheetJS(file).file).sheets[0].cells['2:0'];
    expect(independent.input).toEqual({ type: 'formula', expression: 'UNSUPPORTED(1)' });
    expect(independent.cached).toBeUndefined();
  });
  it('inspects worksheet markup without treating comments or quoted attributes as rules', () => {
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
      files: {
        'xl/worksheets/sheet1.xml': {
          content: new TextEncoder().encode(
            `<?xml version="1.0"?><worksheet><!-- <dataValidation/> --><![CDATA[<conditionalFormatting/>]]><x attr="> <dataValidation/>"></x><ns:dataValidation/><conditionalFormatting/></worksheet>`,
          ),
        },
      },
    };
    const { report } = fromSheetJS(source);
    expect(report.issues.filter((issue) => issue.feature === 'data-validation')).toEqual([
      expect.objectContaining({ count: 1 }),
    ]);
    expect(report.issues.filter((issue) => issue.feature === 'conditional-formatting')).toEqual([
      expect.objectContaining({ count: 1 }),
    ]);
  });
  it('finds declared worksheet parts with quoted greater-than signs in the manifest', () => {
    const encode = (text: string) => new TextEncoder().encode(text);
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
      files: {
        '[Content_Types].xml': {
          content: encode(
            '<Types><Override Extra=">" PartName="/xl/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
          ),
        },
        'xl/custom.xml': { content: encode('<worksheet><dataValidation/></worksheet>') },
      },
    };
    expect(fromSheetJS(source).report.issues).toContainEqual(
      expect.objectContaining({ feature: 'data-validation', count: 1 }),
    );
  });
  it('matches worksheet and manifest part names using the parser case rules', () => {
    const encode = (text: string) => new TextEncoder().encode(text);
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
      files: {
        '[content_types].xml': {
          content: encode(
            '<Types><Override PartName="/XL/Custom.XML" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
          ),
        },
        'xl/custom.xml': {
          content: encode('<worksheet><dataValidation/><conditionalFormatting/></worksheet>'),
        },
      },
    };
    const { report } = fromSheetJS(source);
    expect(report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ feature: 'data-validation', count: 1 }),
        expect.objectContaining({ feature: 'conditional-formatting', count: 1 }),
      ]),
    );
    expect(() =>
      fromSheetJS({
        ...source,
        files: { ...source.files, '/XL/CUSTOM.XML': { content: encode('<worksheet/>') } },
      }),
    ).toThrow('Ambiguous XLSX part names');
  });
  it('scans a long unassigned manifest attribute without hiding declared worksheet rules', () => {
    const encode = (text: string) => new TextEncoder().encode(text);
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
      files: {
        '[Content_Types].xml': {
          content: encode(
            `<Types><Override ${'A'.repeat(100_000)} PartName="/xl/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
          ),
        },
        'xl/custom.xml': { content: encode('<worksheet><dataValidation/></worksheet>') },
      },
    };
    expect(fromSheetJS(source).report.issues).toContainEqual(
      expect.objectContaining({ feature: 'data-validation', count: 1 }),
    );
  });
  it.each([
    '<!--'.repeat(20_000),
    '<![CDATA['.repeat(10_000),
    '<?'.repeat(20_000),
    '<tag "unterminated',
  ])('rejects unterminated worksheet markup', (xml) => {
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
      files: { 'xl/worksheets/sheet1.xml': { content: new TextEncoder().encode(xml) } },
    };
    expect(() => fromSheetJS(source)).toThrow('Unterminated XLSX XML markup');
  });
  it('reports unknown formula functions while excluding quoted text and supported prefixes', () => {
    const source = {
      SheetNames: ['Data'],
      Sheets: {
        Data: {
          A1: { t: 'n', f: 'SUM(1,UNKNOWN(2))', v: 3 },
          A2: { t: 'n', f: '_xlfn.LET(1,2,3)', v: 3 },
          A3: { t: 's', f: '"UNKNOWN(2)"', v: 'UNKNOWN(2)' },
          A4: { t: 'n', f: '_xlfn.IFNA(1,2)', v: 1 },
          A5: { t: 'n', f: "'Sheet (Q1)'!A1+SUM(1,2)", v: 3 },
          '!ref': 'A1:A5',
        },
      },
    };
    const { report } = fromSheetJS(source);
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'UNSUPPORTED_FORMULA_FUNCTION',
        action: 'dropped',
        address: 'A1',
        count: 2,
      }),
    );
    expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
  it('scans long non-call identifiers while still finding nested and prefixed calls', () => {
    const source = {
      SheetNames: ['Data'],
      Sheets: {
        Data: {
          A1: { t: 'n', f: 'A'.repeat(32_000), v: 0 },
          A2: { t: 'n', f: 'SUM(1,_xlfn.UNKNOWN (2))', v: 0 },
          A3: { t: 's', f: '"say ""UNKNOWN(1)"""', v: '' },
          A4: { t: 'n', f: "'Unknown (sheet)'!A1+_xlfn.IFNA(1,2)", v: 0 },
          '!ref': 'A1:A4',
        },
      },
    };
    const { report } = fromSheetJS(source);
    expect(report.issues.filter((issue) => issue.code === 'UNSUPPORTED_FORMULA_FUNCTION')).toEqual([
      expect.objectContaining({ address: 'A2', count: 1 }),
    ]);
    expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
  it('rejects oversized formulas before conversion while accepting the exact text limit', () => {
    const source = (length: number) => ({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', f: 'A'.repeat(length), v: 0 }, '!ref': 'A1' } },
    });
    expect(() => fromSheetJS(source(32_768))).not.toThrow();
    expect(() => fromSheetJS(source(32_769))).toThrow('Formula exceeds text limit');
  });
  it('reports external and structured formula references without matching quoted text', () => {
    const source = {
      SheetNames: ['Data'],
      Sheets: {
        Data: {
          A1: { t: 'n', f: "'[other.xlsx]Sheet 1'!A1", v: 42 },
          A2: { t: 'n', f: 'SUM(Table1[Amount])', v: 4 },
          A3: { t: 's', f: '"[literal]"', v: '[literal]' },
          A4: { t: 's', f: '"say ""[literal]"""', v: 'say "[literal]"' },
          '!ref': 'A1:A4',
        },
      },
    };
    const imported = fromSheetJS(source);
    expect(imported.report.issues).toContainEqual(
      expect.objectContaining({
        code: 'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
        action: 'dropped',
        address: 'A1',
        count: 2,
      }),
    );
    const snapshot = decodeWorkbookFile(imported.file);
    expect(snapshot.sheets[0].cells['0:0'].input).toEqual({
      type: 'formula',
      expression: "'[other.xlsx]Sheet 1'!A1",
    });
    expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    expect(() => toSheetJS(imported.file, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
  it('scans long unmatched brackets and quoted bracket text without losing real references', () => {
    const source = {
      SheetNames: ['Data'],
      Sheets: {
        Data: {
          A1: { t: 'n', f: '['.repeat(32_000), v: 0 },
          A2: { t: 's', f: '"[literal]"&"say ""[still literal]"""', v: '' },
          A3: { t: 'n', f: 'SUM(Table1[[#All],[Amount]])', v: 0 },
          '!ref': 'A1:A3',
        },
      },
    };
    const { report } = fromSheetJS(source);
    expect(
      report.issues.filter(
        (issue) => issue.code === 'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
      ),
    ).toEqual([expect.objectContaining({ address: 'A3', count: 1 })]);
    expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
  it('round-trips notes and approved hyperlinks and reports malformed links', () => {
    const imported = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: {
        Data: {
          A1: {
            t: 's',
            v: 'Source',
            l: { Target: 'https://example.com', Tooltip: 'Reference' },
            c: [{ a: 'Author', t: '备注\nline two' }],
          },
          B1: { t: 's', v: 'Unsafe', l: { Target: 'javascript:alert(1)' } },
          C1: { t: 's', v: 'Malformed', l: { Target: 'https://' } },
          '!ref': 'A1:C1',
        },
      },
    });
    const snapshot = decodeWorkbookFile(imported.file);
    expect(snapshot.sheets[0].cells['0:0'].note).toBe('备注\nline two');
    expect(snapshot.sheets[0].cells['0:0'].link).toEqual({
      target: 'https://example.com',
      tooltip: 'Reference',
    });
    expect(snapshot.sheets[0].cells['0:1'].link).toBeUndefined();
    expect(snapshot.sheets[0].cells['0:2'].link).toBeUndefined();
    expect(imported.report.summary.dropped).toBeGreaterThanOrEqual(2);
    const exported = toSheetJS(imported.file);
    const bytes = XLSX.write(exported.workbook, { type: 'array', bookType: 'xlsx' });
    const restored = decodeWorkbookFile(fromSheetJS(XLSX.read(bytes, { type: 'array' })).file);
    expect(restored.sheets[0].cells['0:0'].note).toBe('备注\nline two');
    expect(restored.sheets[0].cells['0:0'].link?.target).toBe('https://example.com');
  });

  it('recalculates independently saved future text functions with raw OOXML names', async () => {
    const input = XLSX.read(
      readFileSync(
        new URL(
          '../benchmarks/formula-reader-2026-10-23-text-prefixed.recalculated.xlsx',
          import.meta.url,
        ),
      ),
      { type: 'buffer', xlfn: true },
    );
    const cases: [string, string][] = [
      ['_xlfn.CONCAT("x","y")', 'xy'],
      ['TEXTJOIN(0,TRUE,"a","b")', 'a0b'],
    ];
    const sheet = input.Sheets.Semantics;
    const addresses = cases.map(([expression, expected]) => {
      const cell = Object.keys(sheet).find(
        (key) => /^C\d+$/.test(key) && sheet[key].v === expression,
      );
      expect(cell, expression).toBeDefined();
      const address = 'B' + cell!.slice(1);
      expect(sheet[address].f).toMatch(/_xlfn\./);
      expect(sheet[address].v).toBe(expected);
      return address;
    });
    let book = await new Workbook(fromSheetJS(input).file).ready();
    try {
      await book
        .getSheets()[0]
        .range('D1')
        .setValues([['expire caches']]);
      for (const [i, address] of addresses.entries())
        expect(await book.getSheets()[0].range(address).getValues()).toEqual([[cases[i][1]]]);
      const id = book.id;
      await book.close();
      book = await new Workbook(undefined, {}, id).ready();
      for (const [i, address] of addresses.entries())
        expect(await book.getSheets()[0].range(address).getValues()).toEqual([[cases[i][1]]]);
    } finally {
      await book.close();
    }
  });
  it('keeps LibreOffice boolean formulas valid after imported caches expire and reopen', async () => {
    const input = XLSX.read(
      readFileSync(
        new URL('../benchmarks/formula-reader-2026-10-23.recalculated.xlsx', import.meta.url),
      ),
      { type: 'buffer' },
    );
    // Independently generated reader fixture, not an OpenSheet self-roundtrip.
    expect(input.Sheets.Semantics.A1.f).toBe('TRUE()');
    expect(input.Sheets.Semantics.A7.f).toBe('FALSE()');
    const file = fromSheetJS(input).file;
    let book = await new Workbook(file).ready();
    try {
      await book
        .getSheets()[0]
        .range('D1')
        .setValues([['invalidate imported caches']]);
      expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[true]]);
      expect(await book.getSheets()[0].range('A7').getValues()).toEqual([[false]]);
      const id = book.id;
      await book.close();
      book = await new Workbook(undefined, {}, id).ready();
      expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[true]]);
      expect(await book.getSheets()[0].range('A7').getValues()).toEqual([[false]]);
    } finally {
      await book.close();
    }
  });
  it('round-trips formulas without a cached value', async () => {
    const original = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', f: '1+2' } } },
    });
    const exported = toSheetJS(original.file);
    const restored = await new Workbook(fromSheetJS(exported.workbook).file).ready();
    expect(await restored.getSheets()[0].range('A1').getValues()).toEqual([[3]]);
    const bytes = XLSX.write(exported.workbook, { type: 'array', bookType: 'xlsx' });
    const binary = await new Workbook(
      fromSheetJS(XLSX.read(bytes, { type: 'array', sheetStubs: true })).file,
    ).ready();
    expect(await binary.getSheets()[0].range('A1').getValues()).toEqual([[3]]);
  });
  it('performs real XLSX write/read round trips with types, dates, merges, formats and sheet order', async () => {
    const b = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Other' }] }),
      s = b.getSheets()[0];
    await s.range('A1:D2').setValues([
      ['0001', 42, false, null],
      ['date', 45000, true, ''],
    ]);
    await s.range('B2').setStyle({ numberFormat: 'yyyy-mm-dd' });
    await s.range('A4:B4').merge();
    const output = toSheetJS(await b.toJSON());
    const bytes = XLSX.write(output.workbook, { type: 'array', bookType: 'xlsx' });
    const imported = fromSheetJS(
      XLSX.read(bytes, { type: 'array', cellNF: true, sheetStubs: true }),
    );
    const restored = await new Workbook(imported.file).ready();
    expect(restored.getSheets().map((s) => s.name)).toEqual(['Data', 'Other']);
    expect(await restored.getSheets()[0].range('A1:D2').getValues()).toEqual([
      ['0001', 42, false, null],
      ['date', 45000, true, ''],
    ]);
    expect((await restored.getCell(restored.getSheets()[0].id, 1, 1))?.numberFormat).toBe(
      'yyyy-mm-dd',
    );
    expect(decodeWorkbookFile(imported.file).sheets[0].merges).toHaveLength(1);
  });
  it('reads dense and sparse layouts identically', async () => {
    const bytes = XLSX.write(
      XLSX.utils.book_new(
        XLSX.utils.aoa_to_sheet([
          [1, 'text', false],
          ['', null, 2],
        ]),
        'Data',
      ),
      { type: 'array', bookType: 'xlsx' },
    );
    for (const dense of [true, false]) {
      const imported = fromSheetJS(XLSX.read(bytes, { type: 'array', dense }));
      expect(
        await (await new Workbook(imported.file).ready()).getSheets()[0].range('A1:C2').getValues(),
      ).toEqual([
        [1, 'text', false],
        ['', null, 2],
      ]);
    }
  });
  it('rejects malformed dense sheet data instead of silently dropping it', () => {
    for (const data of [undefined, null, false, 'bad', { 0: [{ t: 'n', v: 1 }], length: 1 }]) {
      expect(() =>
        fromSheetJS({ SheetNames: ['Data'], Sheets: { Data: { '!data': data } } } as any),
      ).toThrow(/Invalid dense worksheet data/);
    }
    for (const row of [false, 'bad', { 0: { t: 'n', v: 1 }, length: 1 }]) {
      expect(() =>
        fromSheetJS({ SheetNames: ['Data'], Sheets: { Data: { '!data': [row] } } } as any),
      ).toThrow(/Invalid dense worksheet row/);
    }
    expect(() =>
      fromSheetJS({
        SheetNames: ['Data'],
        Sheets: { Data: { '!data': [null, undefined, [{ t: 'n', v: 1 }]] } },
      } as any),
    ).not.toThrow();
  });
  it('ignores malicious !ref without allocating a rectangle', () => {
    const data = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { '!ref': 'A1:XFD1048576', A1: { t: 'n', v: 1 } } },
    });
    expect(decodeWorkbookFile(data.file).sheets[0].rowOrder).toHaveLength(100);
    expect(() =>
      fromSheetJS({ SheetNames: ['Data'], Sheets: { Data: { A100001: { t: 'n', v: 1 } } } }),
    ).toThrow();
  });
  it('accepts own special sheet names but rejects inherited worksheets', async () => {
    const inherited = Object.create({ Ghost: { A1: { t: 's', v: 'inherited' } } });
    expect(() => fromSheetJS({ SheetNames: ['Ghost'], Sheets: inherited })).toThrow(
      /Missing worksheet/,
    );
    expect(() => fromSheetJS({ SheetNames: ['__proto__'], Sheets: {} })).toThrow(
      /Missing worksheet/,
    );
    const sheets = Object.create(null);
    sheets.__proto__ = { A1: { t: 's', v: 'own value' } };
    const imported = fromSheetJS({ SheetNames: ['__proto__'], Sheets: sheets });
    const book = await new Workbook(imported.file).ready();
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([['own value']]);
    book.dispose();
    const sheet = Object.assign(Object.create({ '!data': [[{ t: 's', v: 'inherited' }]] }), {
      A1: { t: 's', v: 'own cell' },
    });
    const safe = await new Workbook(
      fromSheetJS({ SheetNames: ['Data'], Sheets: { Data: sheet } }).file,
    ).ready();
    expect(await safe.getSheets()[0].range('A1').getValues()).toEqual([['own cell']]);
    safe.dispose();
  });
  it('converts 1904 dates without treating serials as Unix time', async () => {
    const result = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 0, z: 'yyyy-mm-dd' } } },
      Workbook: { WBProps: { date1904: true } },
    });
    const b = await new Workbook(result.file).ready();
    expect(await b.getSheets()[0].range('A1').getDisplayValues()).toEqual([['1904-01-01']]);
  });
  it('surfaces loss and supports strict mode', () => {
    const input = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 's', v: 'styled', s: { font: { bold: true } } } } },
    };
    expect(fromSheetJS(input).report.summary.dropped).toBeGreaterThan(0);
    expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow();
    expect(() =>
      fromSheetJS(
        { SheetNames: ['Data'], Sheets: { Data: { A1: { t: 's', v: 'ok' } } } },
        { unsupported: 'strict' },
      ),
    ).not.toThrow();
  });
  it('enforces read-only for array formula imports', async () => {
    const { file: snapshot } = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 2, f: 'SUM(A2:A3)', F: 'A1:A1' } } },
    });
    const b = await new Workbook(snapshot).ready();
    b.setReadOnly(false);
    await expect(
      async () =>
        await b
          .getSheets()[0]
          .range('A1')
          .setValues([[3]]),
    ).rejects.toThrow(/read only/);
  });
  it('does not propagate stale formula caches', async () => {
    const { file: snapshot } = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 123, f: 'UNSUPPORTED(1)' } } },
    });
    const b = await new Workbook(snapshot).ready(),
      s = b.getSheets()[0];
    expect((await s.range('A1').getDisplayValues())[0][0]).toBe('123 †');
    expect(await b.search(s.id, '123 †')).toEqual({ count: 1, match: { row: 0, column: 0 } });
    await s.range('B1').setValues([[2]]);
    expect((await s.range('A1').getDisplayValues())[0][0]).toBe('#NAME?');
    expect((await b.search(s.id, '123 †')).count).toBe(0);
    expect(await b.search(s.id, '#NAME?')).toEqual({ count: 1, match: { row: 0, column: 0 } });
    expect(toSheetJS(await b.toJSON()).workbook.Sheets.Data.A1.v).toBeUndefined();
  });
  it('exports deep formulas without calculating or writing cached results', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 1 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[1]]);
      await sheet
        .range('A2:A130')
        .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 1}+1`]));
      const file = await book.toJSON();
      const calculate = vi.spyOn(FormulaCalculation.prototype, 'value');
      try {
        const result = toSheetJS(file);
        expect(calculate).not.toHaveBeenCalled();
        expect(result.workbook.Sheets.Data.A130).toMatchObject({ f: 'A129+1', t: 'n' });
        expect(result.workbook.Sheets.Data.A130.v).toBeUndefined();
        expect(result.report.issues).toContainEqual(
          expect.objectContaining({ code: 'RECALC_REQUIRED', count: 129 }),
        );
        const written = XLSX.read(
          XLSX.write(result.workbook, { type: 'buffer', bookType: 'xlsx' }),
          { type: 'buffer', cellFormula: true, sheetStubs: true },
        );
        expect(written.Sheets.Data.A130.f).toBe('A129+1');
        expect(written.Sheets.Data.A130.v).toBeUndefined();
      } finally {
        calculate.mockRestore();
      }
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
});
describe('formats', () => {
  it('parses CSV quotes, multiline fields, empty cells and CRLF', () => {
    expect(parseDelimited('"a,b",c\r\n"line\n2","a""b"\r\n')).toEqual([
      ['a,b', 'c'],
      ['line\n2', 'a"b'],
    ]);
    expect(() => parseDelimited('"unfinished')).toThrow();
  });
  it.each([',', '\t'])(
    'preserves trailing single-column empty rows using separator %s',
    (separator) => {
      for (const rows of [[['first'], [''], ['']], [[''], ['']], [['']]]) {
        const { text } = stringifyDelimited(rows, separator);
        expect(parseDelimited(text, separator)).toEqual(rows);
      }
    },
  );
  it.each([',', '\t'])(
    'round-trips generated rectangular delimited data with separator %s',
    (separator) => {
      let state = 0x4123abcd;
      const random = (n: number) => {
        state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
        return state % n;
      };
      const values = ['', 'x', 'x,y', 'x\ty', 'a"b', 'line\r\nnext', '=1+2'];
      for (let trial = 0; trial < 300; trial++) {
        const height = 1 + random(10);
        const width = 1 + random(3);
        const rows = Array.from({ length: height }, () =>
          Array.from({ length: width }, () => values[random(values.length)]),
        );
        if (trial % 2 === 0) rows[height - 1] = Array(width).fill('');
        expect(parseDelimited(stringifyDelimited(rows, separator, false).text, separator)).toEqual(
          rows,
        );
      }
    },
  );
  it.each([',', '\t'])('preserves a leading data BOM with separator %s', (separator) => {
    const rows = [['\uFEFFvalue'], ['\uFEFF'], ['tail']];
    const text = stringifyDelimited(rows, separator, false).text;
    expect(parseDelimited(text, separator)).toEqual(rows);
    expect(parseDelimited('\uFEFF' + text, separator)).toEqual(rows);
  });
  it('enforces row and column limits on the final field without a newline', () => {
    expect(parseDelimited(Array(1000).fill('x').join(','))[0]).toHaveLength(1000);
    expect(() => parseDelimited(Array(1001).fill('x').join(','))).toThrow(/exceeds limits/);
    expect(() => parseDelimited(`${Array(1001).fill('x').join(',')}\n`)).toThrow(/exceeds limits/);
    expect(parseDelimited(Array(100000).fill('x').join('\n'))).toHaveLength(100000);
    expect(() => parseDelimited(Array(100001).fill('x').join('\n'))).toThrow(/exceeds limits/);
    expect(() => parseDelimited(`${Array(100001).fill('x').join('\n')}\n`)).toThrow(
      /exceeds limits/,
    );
  });
  it('enforces the rectangular cell budget while parsing, including late wide rows', () => {
    const wide = Array(1000).fill('x').join(',');
    const exact = Array(200).fill(wide).join('\n');
    expect(parseDelimited(exact)).toHaveLength(200);
    expect(() => parseDelimited(`${exact}\nx\n"unfinished`)).toThrow('Too many cells');
    const narrow = Array(201).fill('x').join('\n');
    expect(() => parseDelimited(`${narrow}\n${wide}\n"unfinished`)).toThrow('Too many cells');
    expect(parseDelimited('a\nb,c')).toEqual([
      ['a', ''],
      ['b', 'c'],
    ]);
  });
  it('escapes HTML and LaTeX while keeping raw values intact', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1:B2').setValues([
      ['A&B', '<script>'],
      ['50%', 'a_b'],
    ]);
    const html = exportRange(await b.toJSON(), { sheetId: s.id, format: 'html' }).text;
    expect(html).toContain('&lt;script&gt;');
    expect(html).not.toContain('<script>');
    const tex = exportRange(await b.toJSON(), { sheetId: s.id, format: 'latex' });
    expect(tex.text).toContain('A\\&B');
    expect(tex.text).toContain('50\\%');
    expect(tex.text).toContain('a\\_b');
    expect(tex.requiredPackages).toEqual(['booktabs']);
  });
  it('generates merged HTML and LaTeX', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1').setValues([['Group']]);
    await s.range('A1:B2').merge();
    const html = exportRange(await b.toJSON(), {
      sheetId: s.id,
      range: 'A1:C3',
      format: 'html',
    }).text;
    expect(html).toContain('rowspan="2" colspan="2"');
    const tex = exportRange(await b.toJSON(), {
      sheetId: s.id,
      range: 'A1:C3',
      format: 'latex',
    });
    expect(tex.text).toContain('\\multicolumn{2}{c}{\\multirow{2}{*}{Group}}');
  });
  it('protects CSV formula injection without corrupting numeric negatives', () => {
    const result = stringifyDelimited([['=HYPERLINK("bad")', -5, '+1']]);
    expect(result.sanitized).toBe(2);
    expect(result.text).toContain(',-5,');
    expect(stringifyDelimited([['=1']], ',', false).text).toBe('=1');
  });
  it.each([',', '\t'])(
    'escapes dangerous text prefixes and preserves field boundaries with %j',
    (separator) => {
      const dangerous = [
        '=1',
        '+1',
        '-1',
        '@SUM(1)',
        '\ttext',
        '\rtext',
        '\ntext',
        '\0text',
        '＝1',
        '＋1',
        '－1',
        '＠SUM(1)',
        '  ＝1',
        ' \t=1',
      ];
      const breakout = `ordinary${separator}"${separator}=1\nnext`;
      const original = [...dangerous, breakout, -5, 'plain'];
      const safe = stringifyDelimited([original], separator);
      expect(safe.sanitized).toBe(dangerous.length);
      expect(parseDelimited(safe.text, separator)).toEqual([
        [...dangerous.map((value) => `'${value}`), breakout, '-5', 'plain'],
      ]);
      const raw = stringifyDelimited([original], separator, false);
      expect(raw.sanitized).toBe(0);
      expect(parseDelimited(raw.text, separator)).toEqual([
        [...dangerous, breakout, '-5', 'plain'],
      ]);
    },
  );
});

it('preserves empty notes and reports XLSX loss of empty link tooltips', async () => {
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['linked']]);
    await sheet.range('A1').setNote('');
    await sheet.range('B1').setNote('');
    await sheet.range('A1').setLink({ target: 'https://example.com/', tooltip: '' });
    const awaitedFile = await book.toJSON();
    const output = toSheetJS(awaitedFile);
    expect(output.workbook.Sheets[sheet.name].A1.c).toEqual([{ a: 'OpenSheet', t: '' }]);
    expect(output.workbook.Sheets[sheet.name].B1).toMatchObject({
      t: 'z',
      c: [{ a: 'OpenSheet', t: '' }],
    });
    expect(output.report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'EMPTY_LINK_TOOLTIP_DROPPED',
          action: 'dropped',
          address: 'A1',
        }),
      ]),
    );
    expect(() => toSheetJS(awaitedFile, { unsupported: 'strict' })).toThrowError(
      expect.objectContaining({ code: 'UNSUPPORTED_FEATURE' }),
    );
    const bytes = XLSX.write(output.workbook, { type: 'buffer', bookType: 'xlsx' });
    const imported = fromSheetJS(XLSX.read(bytes, { type: 'buffer' }));
    const restored = decodeWorkbookFile(imported.file).sheets[0].cells['0:0'];
    expect(restored.note).toBe('');
    expect(restored.link).toEqual({ target: 'https://example.com/' });
    expect(decodeWorkbookFile(imported.file).sheets[0].cells['0:1'].note).toBe('');
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('exports Unicode link targets as reported, equivalent ASCII URLs', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Links', rows: 4, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A4').setValues([['one'], ['two'], ['ascii'], ['long']]);
    const targets = [
      'https://例子.测试/é/🙂?x=1&y=2#片段',
      'mailto:用户@example.com?subject=你好',
      'https://example.com/a%20b?x=1&y=2#frag',
      'https://example.com/' + 'é'.repeat(1000),
    ];
    for (let i = 0; i < targets.length; i++)
      await sheet.range(`A${i + 1}`).setLink({ target: targets[i] });
    const file = await book.toJSON();
    const result = writeXlsx(file, XLSX);
    expect(result.report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'XLSX_LINK_URL_NORMALIZED', count: 2, address: 'A1' }),
        expect.objectContaining({ code: 'XLSX_LINK_URL_UNSAFE_LENGTH', count: 1, address: 'A4' }),
      ]),
    );
    const sheetjs = XLSX.read(result.bytes, { type: 'array' }).Sheets.Links;
    expect(sheetjs.A1.l.Target).toBe(new URL(targets[0]).href);
    expect(sheetjs.A2.l.Target).toBe(new URL(targets[1]).href);
    expect(sheetjs.A3.l.Target).toBe(targets[2]);
    expect(sheetjs.A4.l.Target).not.toBe(targets[3]);
    const imported = fromSheetJS(XLSX.read(result.bytes, { type: 'array' }));
    expect(imported.report.issues).toContainEqual(
      expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN', count: 1, address: 'A4' }),
    );
    const restored = new FileReader(imported.file);
    try {
      const sheetId = imported.file.sheets[0].id;
      expect(restored.getCell(sheetId, 0, 0)?.link?.target).toBe(new URL(targets[0]).href);
      expect(restored.getCell(sheetId, 1, 0)?.link?.target).toBe(new URL(targets[1]).href);
      expect(restored.getCell(sheetId, 2, 0)?.link?.target).toBe(targets[2]);
    } finally {
      restored.dispose();
    }
    expect(() => writeXlsx(file, XLSX, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    expect((await sheet.range('A1').getLinks())[0][0]?.target).toBe(targets[0]);
  } finally {
    await book.deleteStorage();
  }
});

it('reports nonempty link tooltip round-trip risk without confusing empty or absent tooltips', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 4, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A4').setValues([['one'], ['empty'], ['absent'], ['two']]);
    await sheet.range('A1').setLink({ target: 'https://example.com/', tooltip: 'ordinary tip' });
    await sheet.range('A2').setLink({ target: 'https://example.com/', tooltip: '' });
    await sheet.range('A3').setLink({ target: 'https://example.com/' });
    await sheet.range('A4').setLink({ target: 'https://example.com/', tooltip: 'tip🙂' });
    const file = await book.toJSON();
    const result = writeXlsx(file, XLSX);
    expect(result.report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'LINK_TOOLTIP_INTEROP',
          action: 'approximated',
          count: 2,
          address: 'A1',
        }),
        expect.objectContaining({
          code: 'EMPTY_LINK_TOOLTIP_DROPPED',
          action: 'dropped',
          count: 1,
          address: 'A2',
        }),
      ]),
    );
    const parsed = XLSX.read(result.bytes, { type: 'array' }).Sheets.Data;
    expect(parsed.A1.l.Tooltip).toBe('ordinary tip');
    expect(parsed.A4.l.Tooltip).toBe('tip🙂');
    expect(() => writeXlsx(file, XLSX, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    await sheet.range('A1').setLink({ target: 'https://example.com/' });
    await sheet.range('A2').setLink({ target: 'https://example.com/' });
    await sheet.range('A4').setLink({ target: 'https://example.com/' });
    const withoutTooltips = await book.toJSON();
    expect(() => writeXlsx(withoutTooltips, XLSX, { unsupported: 'strict' })).not.toThrow();
  } finally {
    await book.deleteStorage();
  }
});

it('reports unsupported number formats while retaining their original codes', async () => {
  const source = {
    SheetNames: ['Data'],
    Sheets: {
      Data: {
        A1: { t: 'n', v: 45293.5, z: 'd-mmm-yy h:mm' },
        A2: { t: 'n', v: 1234, z: '0.00E+00' },
        '!ref': 'A1:A2',
      },
    },
  };
  const imported = fromSheetJS(source);
  expect(imported.report.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        code: 'NUMBER_FORMAT_APPROXIMATED',
        action: 'approximated',
        count: 2,
        address: 'A1',
      }),
    ]),
  );
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrowError(
    expect.objectContaining({ code: 'UNSUPPORTED_FEATURE' }),
  );
  const exported = toSheetJS(imported.file);
  expect(exported.workbook.Sheets.Data.A1.z).toBe('d-mmm-yy h:mm');
  expect(exported.workbook.Sheets.Data.A2.z).toBe('0.00E+00');
});

it('recognizes escaped date-time separators and displays the time', async () => {
  const source = {
    SheetNames: ['Data'],
    Sheets: { Data: { A1: { t: 'n', v: 45293.5, z: 'yyyy\\-mm\\-dd\\ hh:mm' }, '!ref': 'A1' } },
  };
  const imported = fromSheetJS(source, { unsupported: 'strict' });
  const book = await new Workbook(imported.file).ready();
  try {
    expect(await book.getSheets()[0].range('A1').getDisplayValues()).toEqual([
      ['2024-01-02 12:00'],
    ]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('converts row pixels to XLSX points and reports font-dependent column widths', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Dimensions', rows: 120, columns: 26 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.setRowHeight(100, 42);
    await sheet.setRowHidden(100, true);
    await sheet.setColumnWidth(20, 140);
    await sheet.setColumnHidden(20, true);
    const file = await book.toJSON();
    const result = toSheetJS(file);
    expect(result.workbook.Sheets.Dimensions['!rows'][100]).toEqual({ hpt: 31.5, hidden: true });
    expect(result.report.issues).toContainEqual(
      expect.objectContaining({
        code: 'EXPORT_WIDTH_APPROXIMATED',
        action: 'approximated',
        count: 1,
      }),
    );
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const bytes = XLSX.write(result.workbook as XLSX.WorkBook, {
      bookType: 'xlsx',
      type: 'buffer',
      cellStyles: true,
    });
    const parsed = XLSX.read(bytes, { type: 'buffer', cellStyles: true });
    expect(parsed.Sheets.Dimensions['!rows']![100]).toMatchObject({ hpt: 31.5, hidden: true });
    const imported = fromSheetJS(parsed);
    expect(imported.report.issues).toContainEqual(
      expect.objectContaining({ code: 'WIDTH_APPROXIMATED', action: 'approximated' }),
    );
    const snapshot = decodeWorkbookFile(imported.file);
    const restored = snapshot.sheets[0];
    expect(restored.rows[restored.rowOrder[100]]).toEqual({ size: 42, hidden: true });
    expect(restored.columns[restored.columnOrder[20]]?.hidden).toBe(true);
    const rowOnly = structuredClone(file);
    delete rowOnly.sheets[0].columns.meta;
    expect(() => toSheetJS(rowOnly, { unsupported: 'strict' })).not.toThrow();
    const zero = structuredClone(rowOnly);
    zero.sheets[0].rows.meta![0] = { size: 0 };
    expect(toSheetJS(zero).report.issues).toContainEqual(
      expect.objectContaining({ code: 'ZERO_ROW_HEIGHT_NOT_EXPORTED', action: 'dropped' }),
    );
    expect(() => toSheetJS(zero, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  } finally {
    await book.deleteStorage();
  }
});

it('imports independently authored row heights in points as pixels', () => {
  const source = XLSX.read(readFileSync('benchmarks/axis-dimension-openpyxl-2026-11-12.xlsx'), {
    type: 'buffer',
    cellStyles: true,
  });
  const imported = fromSheetJS(source);
  const sheet = decodeWorkbookFile(imported.file).sheets[0];
  expect(sheet.rows[sheet.rowOrder[100]]).toEqual({ size: 42, hidden: true });
  expect(sheet.columns[sheet.columnOrder[20]]?.hidden).toBe(true);
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WIDTH_APPROXIMATED' }),
  );
});

it('imports a standard frozen pane and reports independently authored split panes', () => {
  const source = XLSX.read(readFileSync('benchmarks/xlsx-pane-openpyxl-2026-11-12.xlsx'), {
    type: 'buffer',
    bookFiles: true,
  });
  const imported = fromSheetJS(source);
  expect(imported.report.issues).toContainEqual(
    expect.objectContaining({
      code: 'WORKSHEET_PANE_UNSUPPORTED',
      action: 'dropped',
      count: 1,
    }),
  );
  expect(imported.file.sheets.map((sheet) => sheet.freeze)).toEqual([
    { rows: 3, columns: 2 },
    { rows: 0, columns: 0 },
  ]);
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('does not count pane-like text inside worksheet comments or CDATA', () => {
  const source = {
    SheetNames: ['Data'],
    Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
    files: {
      'xl/worksheets/sheet1.xml': {
        content: new TextEncoder().encode(
          '<worksheet><!-- <pane/> --><![CDATA[<pane/>]]><sheetViews/></worksheet>',
        ),
      },
    },
  };
  const result = fromSheetJS(source);
  expect(result.report.issues.some((issue) => issue.code === 'WORKSHEET_PANE_UNSUPPORTED')).toBe(
    false,
  );
});

it('imports independent row-only, column-only and combined frozen panes', () => {
  const source = XLSX.read(readFileSync('benchmarks/xlsx-freeze-openpyxl-2026-11-12.xlsx'), {
    type: 'buffer',
    bookFiles: true,
  });
  const result = fromSheetJS(source);
  expect(result.report.issues.some((issue) => issue.feature === '!margins')).toBe(true);
  expect(result.file.sheets.map((sheet) => sheet.freeze)).toEqual([
    { rows: 4, columns: 0 },
    { rows: 0, columns: 3 },
    { rows: 3, columns: 2 },
  ]);
  expect(result.report.issues.some((issue) => issue.code === 'WORKSHEET_PANE_UNSUPPORTED')).toBe(
    false,
  );
});

it('writes verifiable frozen panes while keeping direct SheetJS export conservative', async () => {
  const book = await createWorkbook({
    sheets: [
      { name: 'Rows', rows: 100, columns: 26 },
      { name: 'Columns', rows: 100, columns: 26 },
      { name: 'Both', rows: 100, columns: 26 },
    ],
  });
  try {
    const sheets = book.getSheets();
    await sheets[0].setFreeze(4, 0);
    await sheets[1].setFreeze(0, 3);
    await sheets[2].setFreeze(3, 2);
    await sheets[0].range('A1').setValues([[42]]);
    const file = await book.toJSON();
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const result = writeXlsx(file, XLSX, { unsupported: 'strict' });
    expect(result.report.issues.some((issue) => issue.code === 'VIEW_NOT_EXPORTED')).toBe(false);
    const parsed = XLSX.read(result.bytes, { type: 'array', bookFiles: true });
    const imported = fromSheetJS(parsed);
    expect(imported.file.sheets.map((sheet) => sheet.freeze)).toEqual([
      { rows: 4, columns: 0 },
      { rows: 0, columns: 3 },
      { rows: 3, columns: 2 },
    ]);
    expect(
      imported.report.issues.some((issue) => issue.code === 'WORKSHEET_PANE_UNSUPPORTED'),
    ).toBe(false);
    const oversized = { ...XLSX, write: () => new Uint8Array(20 * 1024 * 1024 + 1) };
    expect(
      writeXlsx(file, oversized).report.issues.some((issue) => issue.code === 'VIEW_NOT_EXPORTED'),
    ).toBe(true);
    expect(() => writeXlsx(file, oversized, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const largeThenCompressed = {
      ...XLSX,
      write: (workbook: XLSX.WorkBook, options: { compression?: boolean }) =>
        options.compression
          ? XLSX.write(workbook, { type: 'array', bookType: 'xlsx', compression: true })
          : new Uint8Array(20 * 1024 * 1024 + 1),
    };
    const fallback = writeXlsx(file, largeThenCompressed);
    expect(fallback.bytes.byteLength).toBeLessThan(20 * 1024 * 1024);
    expect(fallback.report).toEqual(toSheetJS(file).report);
    expect(fallback.report.issues.some((issue) => issue.code === 'VIEW_NOT_EXPORTED')).toBe(true);
    const fallbackWorkbook = XLSX.read(fallback.bytes, { type: 'array' });
    expect(fallbackWorkbook.SheetNames).toEqual(['Rows', 'Columns', 'Both']);
    expect(fallbackWorkbook.Sheets.Rows.A1.v).toBe(42);
    const withoutFreeze = structuredClone(file);
    for (const sheet of withoutFreeze.sheets) sheet.freeze = { rows: 0, columns: 0 };
    const strictCompressed = writeXlsx(withoutFreeze, largeThenCompressed, {
      unsupported: 'strict',
    });
    expect(strictCompressed.report.issues).toEqual([]);
    expect(XLSX.read(strictCompressed.bytes, { type: 'array' }).Sheets.Rows.A1.v).toBe(42);
    await sheets[0].setFilter({ column: 0, query: '42' });
    const withReport = structuredClone(await book.toJSON());
    withReport.extensions['opensheet.importReport'] = {
      adapterVersion: 'test',
      sourceFormat: 'SheetJS',
      summary: { exact: 0, approximated: 0, dropped: 200, blocked: 0 },
      issues: Array.from({ length: 200 }, (_, i) => ({
        code: `PRIOR_${i}`,
        feature: 'test',
        message: 'Prior loss',
        action: 'dropped',
        severity: 'warning',
      })),
      omittedIssueOccurrences: 0,
    };
    expect(writeXlsx(withReport, largeThenCompressed).report).toEqual(toSheetJS(withReport).report);
    const compressed = {
      ...XLSX,
      write: (workbook: XLSX.WorkBook) =>
        XLSX.write(workbook, { type: 'array', bookType: 'xlsx', compression: true }),
    };
    expect(
      writeXlsx(file, compressed).report.issues.some((issue) => issue.code === 'VIEW_NOT_EXPORTED'),
    ).toBe(true);
    await sheets[0].setFreeze(100, 0);
    const boundary = await book.toJSON();
    expect(
      writeXlsx(boundary, XLSX).report.issues.some((issue) => issue.code === 'VIEW_NOT_EXPORTED'),
    ).toBe(true);
    expect(() => writeXlsx(boundary, XLSX, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  } finally {
    await book.deleteStorage();
  }
});

it('restores exported freeze boundaries after an independent LibreOffice save', () => {
  const input = XLSX.read(
    readFileSync('benchmarks/xlsx-freeze-export-reader-2026-11-12.saved.xlsx'),
    {
      type: 'buffer',
      bookFiles: true,
      sheetStubs: true,
    },
  );
  const result = fromSheetJS(input);
  expect(result.file.sheets.map((sheet) => [sheet.name, sheet.freeze])).toEqual([
    ['Rows', { rows: 4, columns: 0 }],
    ['Columns', { rows: 0, columns: 3 }],
    ["Both '中' &", { rows: 3, columns: 2 }],
    ['Unfrozen', { rows: 0, columns: 0 }],
  ]);
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED' }),
  );
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_ACTIVE_TAB_DROPPED' }),
  );
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_SELECTION_DROPPED' }),
  );
  expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('reports unpaired surrogates in XLSX text and names while preserving valid pairs', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Bad\uD800', rows: 8, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([['\uD800'], ['\uDC00'], ['🙂']]);
    await sheet.range('A4').setNote('note\uD800');
    await sheet.range('A5').setLink({ target: 'https://example.com/', tooltip: 'tip\uDC00' });
    await sheet.range('A6').setFormulas([['"\uD800"']]);
    await sheet.range('A7').setStyle({ numberFormat: '0"\uD800"' });
    await sheet.range('A8').setValues([['\uFFFE']]);
    const file = await book.toJSON();
    const { bytes, report } = writeXlsx(file, XLSX);
    expect(report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'XLSX_UNPAIRED_SURROGATE_SHEET_NAME', count: 1 }),
        expect.objectContaining({ code: 'XLSX_UNPAIRED_SURROGATE_TEXT', count: 6, address: 'A1' }),
        expect.objectContaining({
          code: 'XLSX_CONTROL_TEXT_INTEROP',
          count: 1,
          address: 'A8',
          message: expect.stringContaining('replace Unicode noncharacters'),
        }),
      ]),
    );
    expect(() => writeXlsx(file, XLSX, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const parsed = XLSX.read(bytes, { type: 'array' });
    const output = parsed.Sheets[parsed.SheetNames[0]];
    expect(parsed.SheetNames[0]).not.toBe('Bad\uD800');
    expect(output.A1.v).toBe('�');
    expect(output.A2.v).toBe('�');
    expect(output.A3.v).toBe('🙂');
    expect(output.A4.c[0].t).toBe('note�');
    expect(output.A5.l.Tooltip).toBe('tip�');
    expect(output.A6.f).toBe('"�"');
    expect(output.A8.v).toBe('�');
    const validBook = await createWorkbook({ sheets: [{ name: 'Valid🙂', rows: 1, columns: 1 }] });
    try {
      await validBook
        .getSheets()[0]
        .range('A1')
        .setValues([['🙂']]);
      const validFile = await validBook.toJSON();
      expect(() => writeXlsx(validFile, XLSX, { unsupported: 'strict' })).not.toThrow();
      await validBook
        .getSheets()[0]
        .range('A1')
        .setValues([['\uD800']]);
      const unsupportedFile = await validBook.toJSON();
      try {
        writeXlsx(unsupportedFile, XLSX, { unsupported: 'strict' });
        throw new Error('Expected strict export to fail');
      } catch (error) {
        expect(error).toMatchObject({
          code: 'UNSUPPORTED_FEATURE',
          details: {
            issues: expect.arrayContaining([
              expect.objectContaining({ code: 'XLSX_UNPAIRED_SURROGATE_TEXT' }),
            ]),
          },
        });
      }
      expect(writeXlsx(validFile, XLSX).report.issues).not.toEqual(
        expect.arrayContaining([
          expect.objectContaining({ code: expect.stringContaining('UNPAIRED_SURROGATE') }),
        ]),
      );
    } finally {
      await validBook.deleteStorage();
    }
  } finally {
    await book.deleteStorage();
  }
});

it('reports literal XLSX escape-shaped text that the SheetJS reader changes', async () => {
  const book = await createWorkbook({ sheets: [{ name: '_x0041_', rows: 20, columns: 5 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([['_x0041_'], ['note'], ['link']]);
    await sheet.range('A2').setNote('note _x0041_');
    await sheet.range('A3').setLink({ target: 'https://example.com/_x0041_' });
    await sheet.range('A4').setFormulas([['"_x0041_"']]);
    await sheet.range('A5').setValues([['_xZZZZ_']]);
    await sheet.range('A6:A8').setValues([['before\0after'], ['before\rafter'], ['line\nfeed']]);
    await sheet.range('A9').setValues([['tooltip']]);
    await sheet.range('A9').setLink({ target: 'https://example.com/', tooltip: 'before\rafter' });
    await sheet.range('A10').setLink({ target: 'https://example.com/', tooltip: 'tip _x0041_' });
    const file = await book.toJSON();
    const result = writeXlsx(file, XLSX);
    expect(result.report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'XLSX_ESCAPE_SHEET_NAME_INTEROP', count: 1 }),
        expect.objectContaining({ code: 'XLSX_ESCAPE_TEXT_INTEROP', count: 4, address: 'A1' }),
        expect.objectContaining({ code: 'XLSX_CONTROL_TEXT_INTEROP', count: 3, address: 'A6' }),
      ]),
    );
    expect(() => writeXlsx(file, XLSX, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const parsed = XLSX.read(result.bytes, { type: 'array' });
    expect(parsed.SheetNames).toEqual(['A']);
    const output = parsed.Sheets.A;
    expect(output.A1.v).toBe('A');
    expect(output.A2.c[0].t).toBe('note A');
    expect(output.A3.l.Target).toBe('https://example.com/A');
    expect(output.A4.f).toBe('"A"');
    expect(output.A5.v).toBe('_xZZZZ_');
    expect(output.A9.l.Tooltip).toBe('before_x000d_after');
    expect(output.A10.l.Tooltip).toBe('tip _x0041_');
  } finally {
    await book.deleteStorage();
  }
});
