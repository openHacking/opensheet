import { expect, it } from 'vitest';
import { decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { FileReader } from '@opensheetjs/core';

it('imports ISO dates and explicit instants consistently across host timezones', () => {
  const previous = process.env.TZ;
  try {
    for (const timezone of ['UTC', 'America/Los_Angeles', 'Asia/Shanghai']) {
      process.env.TZ = timezone;
      for (const [value, expected] of [
        ['2024-01-02', 45293],
        ['2024-01-02T00:00:00Z', 45293],
        ['2024-01-02T08:00:00+08:00', 45293],
        ['2024-01-02T12:00', 45293.5],
        ['2024-01-02T17:30:00+05:30', 45293.5],
        ['1900-02-28', 59],
        ['1900-03-01', 61],
      ] as const) {
        const result = fromSheetJS({
          SheetNames: ['Data'],
          Sheets: { Data: { A1: { t: 'd', v: value }, '!ref': 'A1' } },
        });
        expect(
          decodeWorkbookFile(result.file).sheets[0].cells['0:0'].input,
          `${timezone}: ${value}`,
        ).toEqual({ type: 'number', value: expected });
      }
      const local = fromSheetJS({
        SheetNames: ['Data'],
        Sheets: { Data: { A1: { t: 'd', v: new Date(Date.UTC(2024, 0, 2)) }, '!ref': 'A1' } },
        Workbook: { WBProps: { date1904: true } },
      });
      expect(decodeWorkbookFile(local.file).sheets[0].cells['0:0'].input).toEqual({
        type: 'number',
        value: 43831,
      });
    }
  } finally {
    if (previous === undefined) delete process.env.TZ;
    else process.env.TZ = previous;
  }
});

it('rejects invalid and ambiguous date strings instead of normalizing them', () => {
  for (const value of [
    '2024-02-30',
    '01/02/2024',
    '2024-13-01',
    '2024-01-01T25:00:00Z',
    '2024-01-01T12:00+24:00',
    '',
  ]) {
    expect(() =>
      fromSheetJS({
        SheetNames: ['Data'],
        Sheets: { Data: { A1: { t: 'd', v: value }, '!ref': 'A1' } },
      }),
    ).toThrowError(expect.objectContaining({ code: 'INVALID_ARGUMENT' }));
  }
});

it('imports an uncached date-typed formula without inventing a cached date', () => {
  for (const [date1904, expected] of [
    [false, 45293],
    [true, 43831],
  ] as const) {
    const source = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'd', f: 'DATE(2024,1,2)' }, '!ref': 'A1' } },
      Workbook: { WBProps: { date1904 } },
    };
    const imported = fromSheetJS(source);
    const cell = decodeWorkbookFile(imported.file).sheets[0].cells['0:0'];
    expect(cell.input).toEqual({ type: 'formula', expression: 'DATE(2024,1,2)' });
    expect(cell.cached).toBeUndefined();
    expect(imported.report.issues.some((issue) => issue.code === 'DATE_NORMALIZED')).toBe(false);
    const reader = new FileReader(imported.file);
    try {
      expect(reader.getSheetById(imported.file.sheets[0].id)!.range('A1').getValues()).toEqual([
        [expected],
      ]);
    } finally {
      reader.dispose();
    }
  }
  expect(() =>
    fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'd' }, '!ref': 'A1' } },
    }),
  ).toThrowError(expect.objectContaining({ code: 'INVALID_ARGUMENT' }));
});

it('imports real XLSX dates consistently whether SheetJS returns serials or Date objects', async () => {
  const XLSX = await import('xlsx');
  const previous = process.env.TZ;
  try {
    for (const timezone of ['UTC', 'America/Los_Angeles', 'Asia/Shanghai']) {
      process.env.TZ = timezone;
      for (const system of ['1900', '1904'] as const) {
        const raw = {
          SheetNames: ['Data'],
          Sheets: {
            Data: {
              A1: { t: 'n', v: 45293, z: 'yyyy-mm-dd' },
              A2: { t: 'n', v: 45293.5, z: 'yyyy-mm-dd hh:mm' },
              A3: { t: 'n', v: 59, z: 'yyyy-mm-dd' },
              A4: { t: 'n', v: 60, z: 'yyyy-mm-dd' },
              A5: { t: 'n', v: 61, z: 'yyyy-mm-dd' },
              A6: { t: 'n', v: 45361.5, z: 'yyyy-mm-dd hh:mm' },
              A7: { t: 'n', v: 45599.5, z: 'yyyy-mm-dd hh:mm' },
              '!ref': 'A1:A7',
            },
          },
          Workbook: { WBProps: { date1904: system === '1904' } },
        };
        const bytes = XLSX.write(raw, { type: 'buffer', bookType: 'xlsx' });
        const numeric = fromSheetJS(
          XLSX.read(bytes, { type: 'buffer', cellNF: true, cellDates: false }),
        ).file;
        const dates = fromSheetJS(
          XLSX.read(bytes, { type: 'buffer', cellNF: true, cellDates: true }),
        ).file;
        const values = (file: typeof numeric) =>
          Object.values(decodeWorkbookFile(file).sheets[0].cells).map((cell) => cell.input);
        expect(values(dates), `${timezone} ${system}`).toEqual(values(numeric));
      }
    }
  } finally {
    if (previous === undefined) delete process.env.TZ;
    else process.env.TZ = previous;
  }
});

it('matches numeric and Date decoding for independently saved LibreOffice files', async () => {
  const { readFileSync } = await import('node:fs');
  const { createHash } = await import('node:crypto');
  const XLSX = await import('xlsx');
  const fixtures = [
    {
      path: '../benchmarks/interop-reader-2026-10-23-type-probe.authored.xlsx',
      report: '../benchmarks/interop-reader-2026-10-23-type-probe.json',
      artifact: 'authored',
      sheet: 0,
      keys: ['0:3'],
      expected: [45293],
    },
    {
      path: '../benchmarks/formula-reader-2026-10-23-dates-1904.recalculated.xlsx',
      report: '../benchmarks/formula-reader-2026-10-23-dates-1904.json',
      artifact: 'recalculated',
      sheet: 0,
      keys: ['0:1', '1:1', '2:1', '3:1'],
      expected: [0, 59, 43889, 2957003],
    },
    {
      path: '../benchmarks/interop-reader-2026-10-31-dates.reader.xlsx',
      report: '../benchmarks/interop-reader-2026-10-31-dates.json',
      artifact: 'reopened',
      sheet: 0,
      keys: ['0:3'],
      expected: [45293],
    },
  ];
  const previous = process.env.TZ;
  try {
    for (const fixture of fixtures) {
      const bytes = readFileSync(new URL(fixture.path, import.meta.url));
      const evidence = JSON.parse(readFileSync(new URL(fixture.report, import.meta.url), 'utf8'));
      const sha = evidence.recalculatedSha256 ?? evidence.artifacts[fixture.artifact].sha256;
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(sha);
      for (const timezone of ['UTC', 'America/Los_Angeles', 'Asia/Shanghai']) {
        process.env.TZ = timezone;
        const numeric = decodeWorkbookFile(
          fromSheetJS(XLSX.read(bytes, { type: 'buffer', cellNF: true, xlfn: true })).file,
        );
        const dates = decodeWorkbookFile(
          fromSheetJS(
            XLSX.read(bytes, { type: 'buffer', cellNF: true, cellDates: true, xlfn: true }),
          ).file,
        );
        for (const [index, key] of fixture.keys.entries()) {
          const a = numeric.sheets[fixture.sheet].cells[key],
            b = dates.sheets[fixture.sheet].cells[key];
          expect(b.input, `${fixture.path} ${timezone} ${key}`).toEqual(a.input);
          expect(b.cached, `${fixture.path} ${timezone} ${key}`).toEqual(a.cached);
          expect(a.input.type === 'formula' ? a.cached : a.input).toEqual({
            type: 'number',
            value: fixture.expected[index],
          });
        }
      }
    }
  } finally {
    if (previous === undefined) delete process.env.TZ;
    else process.env.TZ = previous;
  }
});
