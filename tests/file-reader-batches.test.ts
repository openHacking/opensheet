import { expect, it } from 'vitest';
import { createWorkbookFile, encodeFileCells, FileReader } from '@opensheetjs/core';

it('reads bounded row batches with values, formatting and cross-sheet formulas', () => {
  const file = createWorkbookFile({
    sheets: [
      { name: 'Data', rows: 5, columns: 3 },
      { name: 'Other', rows: 1, columns: 1 },
    ],
  });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [
          { rowId: 0, columnId: 0, input: { type: 'number', value: 2 }, numberFormat: '0.00' },
          { rowId: 0, columnId: 1, input: { type: 'formula', expression: 'Other!A1+A1' } },
          { rowId: 4, columnId: 2, input: { type: 'string', value: 'last' } },
        ],
        new Map(),
      ),
    },
  ];
  file.sheets[1].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [{ rowId: 0, columnId: 0, input: { type: 'number', value: 3 } }],
        new Map(),
      ),
    },
  ];
  const book = new FileReader(file);
  try {
    const range = book.getSheetById(file.sheets[0].id)!.range('A1:C5');
    expect(() => range.batches({ maxCells: 2 })).toThrow(
      expect.objectContaining({ code: 'INVALID_ARGUMENT' }),
    );
    const options = { maxCells: 6, display: false };
    const batches = range.batches(options);
    options.maxCells = 1;
    range.bounds.endRow = 1;
    const result = [...batches];
    expect(result.map((item) => item.bounds)).toEqual([
      { startRow: 0, endRow: 2, startColumn: 0, endColumn: 3 },
      { startRow: 2, endRow: 4, startColumn: 0, endColumn: 3 },
      { startRow: 4, endRow: 5, startColumn: 0, endColumn: 3 },
    ]);
    expect(result.flatMap((item) => item.values)).toEqual([
      [2, 5, null],
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, 'last'],
    ]);
    const formatted = [
      ...book.getSheetById(file.sheets[0].id)!.range('A1:B1').batches({ display: true }),
    ];
    expect(formatted[0].values).toEqual([['2.00', '5']]);
  } finally {
    book.dispose();
  }
});

it('validates batch limits before iteration and handles a large sparse range incrementally', () => {
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 50000, columns: 1 }] });
  const book = new FileReader(file);
  try {
    const sheet = book.getSheetById(file.sheets[0].id)!;
    const range = sheet.range('A1:A50000');
    expect(() => range.batches({ maxCells: 0 })).toThrow(
      expect.objectContaining({ code: 'INVALID_ARGUMENT' }),
    );
    expect(() => range.batches({ maxCells: 16385 })).toThrow(
      expect.objectContaining({ code: 'INVALID_ARGUMENT' }),
    );
    expect(() => range.batches({ display: 'yes' as never })).toThrow(
      expect.objectContaining({ code: 'INVALID_ARGUMENT' }),
    );
    let count = 0;
    for (const batch of range.batches({ maxCells: 1024 })) {
      expect(batch.values.length).toBeLessThanOrEqual(1024);
      expect(batch.values.every((row) => row.length === 1 && row[0] === null)).toBe(true);
      count += batch.values.length;
    }
    expect(count).toBe(50000);
  } finally {
    book.dispose();
  }
});

it('revalidates changed bounds before allocating a whole range matrix', () => {
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 3, columns: 2 }] });
  const book = new FileReader(file);
  try {
    const range = book.getSheetById(file.sheets[0].id)!.range('A1:B3');
    range.bounds.endRow = 2;
    expect(range.getValues()).toEqual([
      [null, null],
      [null, null],
    ]);
    expect(range.getDisplayValues()).toEqual([
      ['', ''],
      ['', ''],
    ]);
    range.bounds.endRow = 4;
    expect(() => range.getValues()).toThrow(expect.objectContaining({ code: 'INVALID_RANGE' }));
    expect(() => range.getDisplayValues()).toThrow(
      expect.objectContaining({ code: 'INVALID_RANGE' }),
    );
    range.bounds.endRow = 1_000_000_000;
    expect(() => range.getValues()).toThrow(expect.objectContaining({ code: 'LIMIT_EXCEEDED' }));
    range.bounds.endRow = 1.5;
    expect(() => range.getValues()).toThrow(expect.objectContaining({ code: 'INVALID_RANGE' }));
    let reads = 0;
    Object.defineProperty(range.bounds, 'endRow', {
      enumerable: true,
      get: () => (++reads === 1 ? 2 : 1_000_000_000),
    });
    expect(range.getValues()).toEqual([
      [null, null],
      [null, null],
    ]);
    expect(reads).toBe(1);
    reads = 0;
    expect([...range.batches()].flatMap((batch) => batch.values)).toEqual([
      [null, null],
      [null, null],
    ]);
    expect(reads).toBe(1);
  } finally {
    book.dispose();
  }
});

it('does not expose cached formula errors through direct or batched reads', () => {
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [
          { rowId: 0, columnId: 0, input: { type: 'formula', expression: '1/0' } },
          { rowId: 1, columnId: 0, input: { type: 'formula', expression: 'A1+1' } },
        ],
        new Map(),
      ),
    },
  ];
  const book = new FileReader(file);
  try {
    const id = file.sheets[0].id;
    const first = book.value(id, 0, 0) as { error: string };
    expect(first.error).toBe('#DIV/0!');
    first.error = 'modified';
    const range = book.getSheetById(id)!.range('A1:A2');
    expect(range.getValues()).toEqual([[{ error: '#DIV/0!' }], [{ error: '#DIV/0!' }]]);
    const batch = [...range.batches()][0];
    (batch.values[0][0] as { error: string }).error = 'modified again';
    (batch.values[1][0] as { error: string }).error = 'modified dependent';
    expect(range.getValues()).toEqual([[{ error: '#DIV/0!' }], [{ error: '#DIV/0!' }]]);
  } finally {
    book.dispose();
  }
});

it('isolates returned cells, nested metadata and styles from the page cache', () => {
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 2, columns: 1 }] });
  file.styles.push({ numberFormat: '0.00', bold: true });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [
          {
            rowId: 0,
            columnId: 0,
            input: { type: 'number', value: 7 },
            styleId: 's0',
            link: { target: 'https://example.com', tooltip: 'original' },
          },
          {
            rowId: 1,
            columnId: 0,
            input: { type: 'formula', expression: 'A1+1' },
            cached: { type: 'number', value: 8 },
          },
        ],
        new Map([['s0', 0]]),
      ),
    },
  ];
  const book = new FileReader(file);
  try {
    const id = file.sheets[0].id;
    const cell = book.getCell(id, 0, 0)!;
    if (cell.input.type !== 'number') throw new Error('expected number');
    cell.input.value = 99;
    cell.link!.target = 'https://changed.example';
    book.getStyle('s0').numberFormat = '0';
    const formula = [...book.cells(id)][1];
    if (formula.input.type !== 'formula' || formula.cached?.type !== 'number')
      throw new Error('expected formula and cached number');
    formula.input.expression = '99';
    formula.cached.value = 99;
    expect(book.getCell(id, 0, 0)).toMatchObject({
      input: { type: 'number', value: 7 },
      link: { target: 'https://example.com', tooltip: 'original' },
    });
    expect(book.getCell(id, 1, 0)).toMatchObject({
      input: { type: 'formula', expression: 'A1+1' },
      cached: { type: 'number', value: 8 },
    });
    expect(book.getStyle('s0')).toEqual({ numberFormat: '0.00', bold: true });
    expect(book.getSheetById(id)!.range('A1:A2').getValues()).toEqual([[7], [8]]);
    expect(book.display(id, 0, 0)).toBe('7.00');
  } finally {
    book.dispose();
  }
});
