import { describe, expect, it } from 'vitest';
import {
  createWorkbook,
  formatValue,
  isSupportedNumberFormat,
  openWorkbook,
} from '@opensheetjs/core';
import { copyText } from '../packages/renderer/src/clipboard.js';

describe('number format support profile', () => {
  it('formats finite percentage values even when multiplying by 100 would overflow', async () => {
    const value = 1e308;
    for (const format of ['0%', '0.00%', '#,##0%', '$#,##0.00%']) {
      const currency = format.startsWith('$'),
        grouping = format.includes(','),
        decimals = format.includes('.00') ? 2 : 0;
      for (const input of [value, -value, Number.MAX_VALUE, -Number.MAX_VALUE]) {
        let expected = new Intl.NumberFormat('en-US', {
          style: 'percent',
          useGrouping: grouping,
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals,
        }).format(input);
        if (currency)
          expected = expected.startsWith('-') ? '-$' + expected.slice(1) : '$' + expected;
        expect(formatValue(input, format)).toBe(expected);
        expect(expected).not.toContain('∞');
      }
    }
    let book = await createWorkbook();
    try {
      let range = book.getSheets()[0].range('A1');
      await range.setValues([[value]]);
      await range.setStyle({ numberFormat: '0%' });
      const expected = formatValue(value, '0%');
      expect(await range.getDisplayValues()).toEqual([[expected]]);
      expect(await copyText(range)).toBe(expected);
      await book.undo();
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      range = book.getSheets()[0].range('A1');
      expect(await range.getValues()).toEqual([[value]]);
      expect(await range.getDisplayValues()).toEqual([[expected]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
  it('recognizes only the declared display subset', () => {
    for (const code of [
      'General',
      '0',
      '0.00',
      '#,##0.00',
      '$0.00',
      '0.0%',
      'yyyy\\-mm\\-dd\\ hh:mm',
      'h:mm:ss',
      'm/d/yy',
      'm/d/yy h:mm',
    ])
      expect(isSupportedNumberFormat(code), code).toBe(true);
    for (const code of ['0.##', '0;[Red]-0', 'd-mmm-yy', '0.00E+00', 'yyyy-mm-dd hh:mm AM/PM'])
      expect(isSupportedNumberFormat(code), code).toBe(false);
  });
});

describe('date and time presentation', () => {
  it('renders supported date-time and time patterns with rounding across midnight', () => {
    expect(formatValue(45293.5, 'yyyy-mm-dd hh:mm')).toBe('2024-01-02 12:00');
    expect(formatValue(45293.5, 'yyyy-mm-ddThh:mm')).toBe('2024-01-02T12:00');
    expect(formatValue(45293.999999, 'yyyy-mm-dd hh:mm')).toBe('2024-01-03 00:00');
    expect(formatValue(59.999999, 'yyyy-mm-dd hh:mm')).toBe('1900-02-29 00:00');
    expect(formatValue(60.5, 'yyyy-mm-dd hh:mm:ss')).toBe('1900-02-29 12:00:00');
    expect(formatValue(0.5, 'yyyy-mm-dd hh:mm', '1904')).toBe('1904-01-01 12:00');
    expect(formatValue(1.5, 'hh:mm')).toBe('12:00');
    expect(formatValue(0.5, 'h:mm:ss')).toBe('12:00:00');
    expect(formatValue(-0.5, 'hh:mm')).toBe('#####');
  });
  it('keeps datetime display and clipboard text after undo and reopen', async () => {
    let book = await createWorkbook();
    try {
      let range = book.getSheets()[0].range('A1');
      await range.setValues([[45293.5]]);
      await range.setStyle({ numberFormat: 'yyyy-mm-dd hh:mm' });
      expect(await range.getDisplayValues()).toEqual([['2024-01-02 12:00']]);
      expect(await copyText(range)).toBe('2024-01-02 12:00');
      await book.undo();
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      range = book.getSheets()[0].range('A1');
      expect(await range.getValues()).toEqual([[45293.5]]);
      expect(await range.getDisplayValues()).toEqual([['2024-01-02 12:00']]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
});

it('pins independent LibreOffice date-time observations and their known differences', async () => {
  const { readFileSync } = await import('node:fs');
  const { createHash } = await import('node:crypto');
  const { parseDelimited } = await import('@opensheetjs/formats');
  const root = new URL('../benchmarks/', import.meta.url);
  const report = JSON.parse(readFileSync(new URL('datetime-format-2026-10-31.json', root), 'utf8'));
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  }
  for (const result of report.results as {
    system: '1900' | '1904';
    comparisons: {
      value: number;
      format: string;
      opensheet: string;
      reader: string;
      agrees: boolean;
    }[];
  }[]) {
    const csv = readFileSync(new URL(report.artifacts[`${result.system}csv`].file, root), 'utf8');
    expect(parseDelimited(csv).map((row) => row[0])).toEqual(
      result.comparisons.map((item) => item.reader),
    );
    expect(
      result.comparisons.map((item) => formatValue(item.value, item.format, result.system)),
    ).toEqual(result.comparisons.map((item) => item.opensheet));
  }
  expect(
    report.results.map((result: { comparisons: { agrees: boolean }[] }) =>
      result.comparisons.map((item) => item.agrees),
    ),
  ).toEqual([
    [true, false, true, false],
    [true, false, true, true],
  ]);
});

it('renders common slash dates and date-times including 1900 serial boundaries', () => {
  for (const [value, format, expected] of [
    [0, 'm/d/yy', '1/0/00'],
    [59, 'm/d/yy', '2/28/00'],
    [60, 'mm/dd/yyyy', '02/29/1900'],
    [61, 'm/d/yy', '3/1/00'],
    [45293.5, 'm/d/yy h:mm', '1/2/24 12:00'],
    [45293.999999, 'm/d/yy h:mm', '1/3/24 0:00'],
    [0, 'yyyy-mm-dd', '1900-01-00'],
  ] as const)
    expect(formatValue(value, format), `${value} ${format}`).toBe(expected);
  expect(formatValue(0, 'm/d/yy', '1904')).toBe('1/1/04');
});

it('documents SheetJS SSF date/time boundary differences', async () => {
  const XLSX = await import('xlsx');
  const differences: unknown[] = [];
  for (const system of ['1900', '1904'] as const) {
    for (const format of [
      'yyyy-mm-dd',
      'm/d/yy',
      'mm/dd/yyyy',
      'm/d/yy h:mm',
      'yyyy-mm-dd hh:mm:ss',
      'hh:mm',
    ]) {
      for (const value of [0, 0.5, 59, 59.999999, 60, 61, 45293.5, 45293.999999]) {
        const actual = formatValue(value, format, system);
        const reference = XLSX.SSF.format(format, value, { date1904: system === '1904' });
        if (actual !== reference) differences.push({ system, format, value, actual, reference });
      }
    }
  }
  expect(differences).toEqual([
    {
      system: '1904',
      format: 'm/d/yy h:mm',
      value: 59.999999,
      actual: '3/1/04 0:00',
      reference: '2/29/00 0:00',
    },
    {
      system: '1904',
      format: 'm/d/yy h:mm',
      value: 45293.999999,
      actual: '1/4/28 0:00',
      reference: '1/3/24 0:00',
    },
    {
      system: '1904',
      format: 'yyyy-mm-dd hh:mm:ss',
      value: 59.999999,
      actual: '1904-03-01 00:00:00',
      reference: '1900-02-29 00:00:00',
    },
    {
      system: '1904',
      format: 'yyyy-mm-dd hh:mm:ss',
      value: 45293.999999,
      actual: '2028-01-04 00:00:00',
      reference: '2024-01-03 00:00:00',
    },
  ]);
});

it('records LibreOffice slash-date display differences with hash-pinned artifacts', async () => {
  const { readFileSync } = await import('node:fs');
  const { createHash } = await import('node:crypto');
  const { parseDelimited } = await import('@opensheetjs/formats');
  const root = new URL('../benchmarks/', import.meta.url);
  const report = JSON.parse(
    readFileSync(new URL('datetime-format-2026-10-31-slash.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  for (const result of report.results as {
    system: '1900' | '1904';
    comparisons: { value: number; format: string; opensheet: string; reader: string }[];
  }[]) {
    const csv = readFileSync(new URL(report.artifacts[`${result.system}csv`].file, root), 'utf8');
    expect(parseDelimited(csv).map((row) => row[0])).toEqual(
      result.comparisons.map((item) => item.reader),
    );
    expect(
      result.comparisons.map((item) => formatValue(item.value, item.format, result.system)),
    ).toEqual(result.comparisons.map((item) => item.opensheet));
  }
  expect(
    report.results.map((result: { comparisons: { agrees: boolean }[] }) =>
      result.comparisons.map((item) => item.agrees),
    ),
  ).toEqual([
    [true, false, true, false, false, false, false, false],
    [true, false, true, true, false, true, false, false],
  ]);
});

it('keeps numeric formatting equivalent to the uncached Intl options', () => {
  const values = [-1234567.125, -0, 0, 0.000001, 1.005, 1234567.891, 1e20];
  for (let digits = 0; digits <= 12; digits++) {
    for (const grouping of [false, true]) {
      const pattern = (grouping ? '#,##0' : '0') + (digits ? '.' + '0'.repeat(digits) : '');
      for (const value of values) {
        const expected = value.toLocaleString('en-US', {
          minimumFractionDigits: digits,
          maximumFractionDigits: digits,
          useGrouping: grouping,
        });
        expect(formatValue(value, pattern), `${pattern} ${value}`).toBe(expected);
      }
    }
  }
});

it('places the sign before currency for negative values, including rounded negative zero', () => {
  for (const [value, format, expected] of [
    [-1234.56, '$0.00', '-$1234.56'],
    [-1234.56, '$#,##0.00', '-$1,234.56'],
    [-0, '$0.00', '-$0.00'],
    [-0.001, '$0.00', '-$0.00'],
    [1234.56, '$#,##0.00', '$1,234.56'],
    [-0.125, '$0.0%', '-$12.5%'],
  ] as const)
    expect(formatValue(value, format)).toBe(expected);
});

it('preserves negative currency presentation through clipboard, undo and reopen', async () => {
  let book = await createWorkbook();
  try {
    let range = book.getSheets()[0].range('A1');
    await range.setValues([[-1234.56]]);
    await range.setStyle({ numberFormat: '$#,##0.00' });
    expect(await range.getDisplayValues()).toEqual([['-$1,234.56']]);
    expect(await copyText(range)).toBe('-$1,234.56');
    await book.undo();
    await book.redo();
    const id = book.id;
    await book.close();
    book = await openWorkbook(id);
    range = book.getSheets()[0].range('A1');
    expect(await range.getDisplayValues()).toEqual([['-$1,234.56']]);
    expect(await copyText(range)).toBe('-$1,234.56');
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('pins independent numeric display observations and rounded-zero differences', async () => {
  const { readFileSync } = await import('node:fs');
  const { createHash } = await import('node:crypto');
  const { parseDelimited } = await import('@opensheetjs/formats');
  const root = new URL('../benchmarks/', import.meta.url);
  const report = JSON.parse(readFileSync(new URL('numeric-format-2026-10-31.json', root), 'utf8'));
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  for (const result of report.results as {
    system: '1900' | '1904';
    comparisons: {
      value: number;
      format: string;
      opensheet: string;
      reader: string;
      agrees: boolean;
    }[];
  }[]) {
    const csv = readFileSync(new URL(report.artifacts[`${result.system}csv`].file, root), 'utf8');
    expect(parseDelimited(csv).map((row) => row[0])).toEqual(
      result.comparisons.map((item) => item.reader),
    );
    expect(
      result.comparisons.map((item) => formatValue(item.value, item.format, result.system)),
    ).toEqual(result.comparisons.map((item) => item.opensheet));
    expect(result.comparisons.map((item) => item.opensheet === item.reader)).toEqual(
      result.comparisons.map((item) => item.agrees),
    );
    expect(result.comparisons.map((item) => item.agrees)).toEqual([
      true,
      true,
      false,
      true,
      true,
      true,
      true,
      true,
    ]);
  }
});
