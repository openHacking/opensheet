import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { xlsxEntityEscape } from './fixtures/xlsx-entity-escape.js';

it('reports ambiguous escape-shaped text from raw XLSX parts before conversion', () => {
  const bytes = readFileSync(
    new URL('../benchmarks/xlsx-escape-reader-2026-11-12.xlsx', import.meta.url),
  );
  const parsed = XLSX.read(bytes, { type: 'buffer', bookFiles: true });
  const result = fromSheetJS(parsed);
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_ESCAPE_IMPORT_AMBIGUOUS', action: 'approximated' }),
  );
  expect(() => fromSheetJS(parsed, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
  const withoutRaw = { ...parsed, files: undefined };
  expect(
    fromSheetJS(withoutRaw).report.issues.some(
      (issue) => issue.code === 'XLSX_ESCAPE_IMPORT_AMBIGUOUS',
    ),
  ).toBe(false);
});

it('detects a character-reference escape in a real XLSX parsed by SheetJS', () => {
  const parsed = XLSX.read(xlsxEntityEscape(), { type: 'buffer', bookFiles: true });
  expect(parsed.Sheets.Data.A1.v).toBe('A');
  expect(fromSheetJS(parsed).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_ESCAPE_IMPORT_AMBIGUOUS' }),
  );
});

it('detects character references and UTF-16 parts and ignores non-hex lookalikes', () => {
  const make = (text: string, encoding: 'utf8' | 'utf16le' | 'utf16be', bom = true) => {
    let data: Uint8Array;
    if (encoding === 'utf8') data = new TextEncoder().encode(text);
    else {
      const le = Buffer.from(text, 'utf16le');
      if (encoding === 'utf16be') le.swap16();
      data = bom
        ? Buffer.concat([Buffer.from(encoding === 'utf16le' ? [0xff, 0xfe] : [0xfe, 0xff]), le])
        : le;
    }
    return {
      SheetNames: ['Sheet'],
      Sheets: { Sheet: { A1: { t: 'n', v: 1 } } },
      files: { 'xl/sharedStrings.xml': { content: data } },
    };
  };
  for (const encoding of ['utf8', 'utf16le', 'utf16be'] as const)
    for (const bom of encoding === 'utf8' ? [false] : [true, false]) {
      for (const text of ['_X0041_', '&#95;x0041_', '_&#120;004&#49;&#x5f;'])
        expect(
          fromSheetJS(make(`<t>${text}</t>`, encoding, bom)).report.issues.some(
            (issue) => issue.code === 'XLSX_ESCAPE_IMPORT_AMBIGUOUS',
          ),
        ).toBe(true);
      expect(
        fromSheetJS(make('<t>_xZZZZ_</t>', encoding, bom)).report.issues.some(
          (issue) => issue.code === 'XLSX_ESCAPE_IMPORT_AMBIGUOUS',
        ),
      ).toBe(false);
    }
});

it('reports inspection as incomplete for a numeric reference exceeding the bound', () => {
  const content = new TextEncoder().encode('&#' + '0'.repeat(21) + '95;x0041_');
  const input = {
    SheetNames: ['Sheet'],
    Sheets: { Sheet: { A1: { t: 'n', v: 1 } } },
    files: { 'xl/sharedStrings.xml': { content } },
  };
  expect(fromSheetJS(input).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_ESCAPE_INSPECTION_INCOMPLETE' }),
  );
});

it('reports incomplete inspection when the candidate work budget is exhausted', () => {
  const content = new TextEncoder().encode('_'.repeat(1_000_001) + '_x0041_');
  const input = {
    SheetNames: ['Sheet'],
    Sheets: { Sheet: { A1: { t: 'n', v: 1 } } },
    files: { 'xl/sharedStrings.xml': { content } },
  };
  expect(fromSheetJS(input).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_ESCAPE_INSPECTION_INCOMPLETE', action: 'approximated' }),
  );
  expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});
