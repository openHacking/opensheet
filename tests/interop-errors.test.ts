import { expect, it } from 'vitest';
import { createWorkbook } from '@opensheetjs/core';
import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs';

it('rejects malformed boolean cells instead of coercing them', () => {
  for (const value of ['false', 0, 1, null]) {
    expect(() =>
      fromSheetJS({
        SheetNames: ['Data'],
        Sheets: { Data: { A1: { t: 'b', v: value }, '!ref': 'A1' } },
      }),
    ).toThrowError(expect.objectContaining({ code: 'INVALID_ARGUMENT' }));
  }
  const imported = fromSheetJS({
    SheetNames: ['Data'],
    Sheets: { Data: { A1: { t: 'b', v: false }, B1: { t: 'b', f: 'TRUE()' }, '!ref': 'A1:B1' } },
  });
  expect(imported.report.summary.blocked).toBe(0);
});

it('reports unmapped spreadsheet errors on import and export, and rejects them in strict mode', async () => {
  const source = {
    SheetNames: ['Data'],
    Sheets: { Data: { A1: { t: 'e', v: 255 }, '!ref': 'A1' } },
  };
  const imported = fromSheetJS(source);
  expect(imported.report.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        code: 'ERROR_CODE_APPROXIMATED',
        action: 'approximated',
        address: 'A1',
      }),
    ]),
  );
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrowError(
    expect.objectContaining({ code: 'UNSUPPORTED_FEATURE' }),
  );
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'error', code: '#CYCLE!' } }],
      },
    });
    const file = await book.toJSON();
    const exported = toSheetJS(file);
    expect(exported.workbook.Sheets[sheet.name].A1).toMatchObject({ t: 'e', v: 15 });
    expect(exported.report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'ERROR_CODE_APPROXIMATED',
          action: 'approximated',
          address: 'A1',
        }),
      ]),
    );
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrowError(
      expect.objectContaining({ code: 'UNSUPPORTED_FEATURE' }),
    );
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('preserves known error values and rejects malformed error codes', async () => {
  const codes = [0, 7, 15, 23, 29, 36, 42, 43];
  const cells = Object.fromEntries(codes.map((value, i) => [`A${i + 1}`, { t: 'e', v: value }]));
  const source = { SheetNames: ['Data'], Sheets: { Data: { ...cells, '!ref': 'A1:A8' } } };
  const imported = fromSheetJS(source, { unsupported: 'strict' });
  const exported = toSheetJS(imported.file, { unsupported: 'strict' });
  for (const [i, value] of codes.entries())
    expect(exported.workbook.Sheets.Data[`A${i + 1}`]).toMatchObject({ t: 'e', v: value });
  for (const value of ['__proto__', 0.5, null, Infinity]) {
    expect(() =>
      fromSheetJS({
        SheetNames: ['Data'],
        Sheets: { Data: { A1: { t: 'e', v: value }, '!ref': 'A1' } },
      }),
    ).toThrowError(expect.objectContaining({ code: 'INVALID_ARGUMENT' }));
  }
});
