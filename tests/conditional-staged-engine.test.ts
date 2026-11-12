import { expect, it, vi } from 'vitest';
import { Engine } from '../packages/core/src/engine.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { Database } from '../packages/core/src/storage.js';
import { key } from '../packages/core/src/types.js';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import {
  conditionalViewportWithFormulas,
  validateConditionalRules,
} from '../packages/core/src/conditional-format.js';

it('calculates conditional appearance against immutable staged and committed heads', async () => {
  const snapshot = createSnapshot({
    sheets: [
      { name: 'Data', rows: 2, columns: 1 },
      { name: 'Limits', rows: 1, columns: 1 },
    ],
  });
  const [data, limits] = snapshot.sheets;
  snapshot.styles.style1 = { color: '#111111', bold: true };
  data.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 5 } };
  data.cells[key(1, 0)] = {
    rowId: 1,
    columnId: 0,
    input: { type: 'formula', expression: 'A1*2' },
    styleId: 'style1',
  };
  limits.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 10 } };
  const name = `conditional-staged-${crypto.randomUUID()}`;
  let db = await Database.open(name);
  let engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  let internal = engine as any;
  const rules = [
    {
      id: 'over-limit',
      range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      condition: { kind: 'formula', expression: 'A1>Limits!$A$1' },
      style: { background: '#ff0000' },
      stopIfTrue: false,
    },
  ];
  const positions = [
    { row: 0, column: 0 },
    { row: 1, column: 0 },
  ];
  const oldStyles = [{}, { color: '#111111', bold: true }];
  const newStyles = [
    { background: '#ff0000' },
    { color: '#111111', bold: true, background: '#ff0000' },
  ];
  const run = () =>
    internal.conditionalStagedViewport(internal.head, internal.metadata, data.id, rules, positions);
  try {
    await engine.read(data.id, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
    expect(await run()).toEqual(oldStyles);
    const publish = internal.publish.bind(engine);
    const spy = vi.spyOn(internal, 'publish').mockImplementation(async (next, metadata, ranges) => {
      expect(await run()).toEqual(oldStyles);
      expect(
        await internal.conditionalStagedViewport(next, metadata, data.id, rules, positions),
      ).toEqual(newStyles);
      expect(engine.view().revision).toBe(0);
      expect((await engine.cell(data.id, 0, 0))?.input).toEqual({ type: 'number', value: 5 });
      return publish(next, metadata, ranges);
    });
    await engine.execute([
      {
        type: 'core.cells.set',
        payload: {
          sheetId: data.id,
          cells: [{ row: 0, column: 0, input: { type: 'number', value: 12 } }],
        },
      },
      {
        type: 'core.cells.set',
        payload: {
          sheetId: limits.id,
          cells: [{ row: 0, column: 0, input: { type: 'number', value: 8 } }],
        },
      },
    ]);
    spy.mockRestore();
    expect(await run()).toEqual(newStyles);
    await engine.history(false);
    expect(await run()).toEqual(oldStyles);
    await engine.history(true);
    expect(await run()).toEqual(newStyles);
    engine.close();
    db = await Database.open(name);
    engine = new Engine(db, snapshot.workbookId, 0);
    await engine.open();
    internal = engine as any;
    expect(await run()).toEqual(newStyles);
  } finally {
    engine.close();
  }
});

it('rejects invalid work and cancellation before returning a staged appearance', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 2, columns: 1 }] });
  const db = await Database.open(`conditional-limits-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any,
    sheetId = snapshot.sheets[0].id;
  const rule = {
    id: 'x',
    range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    condition: { kind: 'number', operator: 'gt', first: 0 },
    style: { bold: true },
    stopIfTrue: false,
  };
  try {
    await expect(
      internal.conditionalStagedViewport(
        internal.head,
        internal.metadata,
        sheetId,
        [rule],
        [{ row: 2, column: 0 }],
      ),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    const spy = vi.spyOn(internal, 'stagedCell');
    const rules = Array.from({ length: 101 }, (_, index) => ({ ...rule, id: `r${index}` }));
    await expect(
      internal.conditionalStagedViewport(
        internal.head,
        internal.metadata,
        sheetId,
        rules,
        Array.from({ length: 10000 }, () => ({ row: 0, column: 0 })),
      ),
    ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    expect(spy).not.toHaveBeenCalled();
    let release!: (value: undefined) => void;
    let started!: () => void;
    const ready = new Promise<void>((resolve) => {
      started = resolve;
    });
    spy.mockImplementation(() => {
      started();
      return new Promise((resolve) => {
        release = resolve;
      });
    });
    const controller = new AbortController();
    const pending = internal.conditionalStagedViewport(
      internal.head,
      internal.metadata,
      sheetId,
      [rule],
      [{ row: 0, column: 0 }],
      controller.signal,
    );
    await ready;
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    release(undefined);
    spy.mockRestore();
  } finally {
    engine.close();
  }
});

it('reuses the staged cell for calculation and stops before retaining excessive text', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 2, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells[key(0, 0)] = {
    rowId: 0,
    columnId: 0,
    input: { type: 'string', value: 'a'.repeat(32768) },
  };
  const db = await Database.open(`conditional-text-budget-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any;
  const rules = [
    {
      id: 'x',
      range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      condition: { kind: 'number', operator: 'gt', first: 0 },
      style: { bold: true },
      stopIfTrue: false,
    },
  ];
  const spy = vi.spyOn(internal, 'stagedCell');
  try {
    await expect(
      internal.conditionalStagedViewport(
        internal.head,
        internal.metadata,
        sheet.id,
        rules,
        Array.from({ length: 400 }, () => ({ row: 0, column: 0 })),
      ),
    ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    expect(spy).toHaveBeenCalledTimes(306);
  } finally {
    spy.mockRestore();
    engine.close();
  }
});

it('reuses only viewport scalar inputs for formula predicates with a zero cache reservation', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 200, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  for (let row = 0; row < 200; row++)
    sheet.cells[key(row, 0)] = { rowId: row, columnId: 0, input: { type: 'number', value: row } };
  const db = await Database.open(`conditional-scalar-reuse-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any;
  const rules = [
    {
      id: 'positive',
      range: { startRow: 0, endRow: 200, startColumn: 0, endColumn: 1 },
      condition: { kind: 'formula', expression: 'A1>0' },
      style: { bold: true },
      stopIfTrue: false,
    },
  ];
  const positions = Array.from({ length: 200 }, (_, row) => ({ row, column: 0 }));
  const spy = vi.spyOn(internal, 'stagedCell');
  const pages = vi.spyOn(internal, 'block');
  try {
    const styles = await internal.conditionalStagedViewport(
      internal.head,
      internal.metadata,
      sheet.id,
      rules,
      positions,
    );
    expect(styles).toHaveLength(200);
    expect(styles.every((style: { bold?: boolean }, row: number) => !!style.bold === row > 0)).toBe(
      true,
    );
    expect(spy).toHaveBeenCalledTimes(200);
    expect(pages).toHaveBeenCalledTimes(4);
  } finally {
    spy.mockRestore();
    pages.mockRestore();
    engine.close();
  }
});

it('does not retain an oversized page during viewport reads', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 3, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  for (let row = 0; row < 3; row++)
    sheet.cells[key(row, 0)] = { rowId: row, columnId: 0, input: { type: 'number', value: row } };
  const db = await Database.open(`conditional-page-cap-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any,
    original = internal.block.bind(engine);
  const spy = vi.spyOn(internal, 'block').mockImplementation(async (...args) => {
    const page = await original(...args);
    Object.defineProperty(page, 'bytes', { value: 1024 * 1024 + 1 });
    return page;
  });
  const rules = [
    {
      id: 'positive',
      range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
      condition: { kind: 'number', operator: 'gte', first: 0 },
      style: { bold: true },
      stopIfTrue: false,
    },
  ];
  try {
    const styles = await internal.conditionalStagedViewport(
      internal.head,
      internal.metadata,
      sheet.id,
      rules,
      Array.from({ length: 3 }, (_, row) => ({ row, column: 0 })),
    );
    expect(styles).toEqual([{ bold: true }, { bold: true }, { bold: true }]);
    expect(spy).toHaveBeenCalledTimes(3);
  } finally {
    spy.mockRestore();
    engine.close();
  }
});

it('does not retain an actually oversized decoded page', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 40, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  for (let row = 0; row < 40; row++)
    sheet.cells[key(row, 0)] = {
      rowId: row,
      columnId: 0,
      input: { type: 'string', value: `${row}:` + 'x'.repeat(32760) },
    };
  const db = await Database.open(`conditional-real-page-cap-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any,
    original = internal.block.bind(engine),
    sizes: number[] = [];
  const spy = vi.spyOn(internal, 'block').mockImplementation(async (...args) => {
    const page = await original(...args);
    sizes.push(page.bytes);
    return page;
  });
  const rules = [
    {
      id: 'text',
      range: { startRow: 0, endRow: 40, startColumn: 0, endColumn: 1 },
      condition: { kind: 'text', operator: 'contains', value: 'x', caseSensitive: true },
      style: { bold: true },
      stopIfTrue: false,
    },
  ];
  try {
    const result = await internal.conditionalStagedViewport(
      internal.head,
      internal.metadata,
      sheet.id,
      rules,
      [0, 1, 2].map((row) => ({ row, column: 0 })),
    );
    expect(result).toEqual([{ bold: true }, { bold: true }, { bold: true }]);
    expect(spy).toHaveBeenCalledTimes(3);
    expect(sizes.every((size) => size > 1024 * 1024)).toBe(true);
  } finally {
    spy.mockRestore();
    engine.close();
  }
});

it('keeps viewport order and cell identities when reads cross reordered pages', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 130, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  [sheet.rowOrder[0], sheet.rowOrder[65]] = [sheet.rowOrder[65], sheet.rowOrder[0]];
  sheet.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 10 } };
  sheet.cells[key(65, 0)] = { rowId: 65, columnId: 0, input: { type: 'number', value: 20 } };
  const db = await Database.open(`conditional-reordered-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any,
    pages = vi.spyOn(internal, 'block');
  const rules = [
    {
      id: 'high',
      range: { startRow: 0, endRow: 130, startColumn: 0, endColumn: 1 },
      condition: { kind: 'number', operator: 'gte', first: 15 },
      style: { bold: true },
      stopIfTrue: false,
    },
  ];
  try {
    const positions = [0, 1, 65, 0].map((row) => ({ row, column: 0 }));
    expect(
      await internal.conditionalStagedViewport(
        internal.head,
        internal.metadata,
        sheet.id,
        rules,
        positions,
      ),
    ).toEqual([{ bold: true }, {}, {}, { bold: true }]);
    expect(pages).toHaveBeenCalledTimes(3);
  } finally {
    pages.mockRestore();
    engine.close();
  }
});

it('preserves blank, scalar error and formula-cycle semantics while reusing inputs', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 5, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 5 } };
  sheet.cells[key(1, 0)] = {
    rowId: 1,
    columnId: 0,
    input: { type: 'formula', expression: 'A1*2' },
  };
  sheet.cells[key(3, 0)] = { rowId: 3, columnId: 0, input: { type: 'error', code: '#DIV/0!' } };
  sheet.cells[key(4, 0)] = { rowId: 4, columnId: 0, input: { type: 'formula', expression: 'A5' } };
  const db = await Database.open(`conditional-errors-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any;
  const rules = [
    {
      id: 'error',
      range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 },
      condition: { kind: 'formula', expression: 'IFERROR(A1,42)=42' },
      style: { bold: true },
      stopIfTrue: false,
    },
  ];
  try {
    const styles = await internal.conditionalStagedViewport(
      internal.head,
      internal.metadata,
      sheet.id,
      rules,
      Array.from({ length: 5 }, (_, row) => ({ row, column: 0 })),
    );
    expect(styles).toEqual([{}, {}, {}, { bold: true }, {}]);
  } finally {
    engine.close();
  }
});

it('agrees with an uncached reader for mixed scalar and formula references', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 9, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  const inputs = [
    { type: 'number', value: 5 },
    { type: 'string', value: 'text' },
    null,
    { type: 'error', code: '#DIV/0!' },
    { type: 'formula', expression: 'A1*2' },
    { type: 'formula', expression: 'A6' },
    { type: 'boolean', value: true },
    { type: 'formula', expression: 'IFERROR(A4,7)' },
    { type: 'formula', expression: 'A5+A1' },
  ] as const;
  inputs.forEach((input, row) => {
    if (input) sheet.cells[key(row, 0)] = { rowId: row, columnId: 0, input };
  });
  const db = await Database.open(`conditional-differential-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any;
  const range = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 1 };
  const rules = validateConditionalRules(
    [
      {
        id: 'error',
        range,
        condition: { kind: 'formula', expression: 'IFERROR(A1,42)=42' },
        style: { color: '#ff0000' },
        stopIfTrue: false,
      },
      {
        id: 'large',
        range,
        condition: { kind: 'formula', expression: 'A1>3' },
        style: { bold: true },
        stopIfTrue: false,
      },
    ],
    9,
    1,
  );
  const positions = Array.from({ length: 9 }, (_, row) => ({ row, column: 0 }));
  try {
    const calculator = new AsyncCalculation(
      (id, row, column) => internal.stagedCell(id, row, column, internal.head, internal.metadata),
      () => undefined,
      0,
    );
    const reader = calculator.validationSession().forSheet(sheet.id);
    const cells = [];
    for (const position of positions)
      cells.push({
        ...position,
        value: await reader.read(position.row, position.column),
        base: {},
      });
    const expected = await conditionalViewportWithFormulas(rules, cells, reader.evaluateFormula);
    const actual = await internal.conditionalStagedViewport(
      internal.head,
      internal.metadata,
      sheet.id,
      rules,
      positions,
    );
    expect(actual).toEqual(expected);
    expect(actual[3]).toEqual({ color: '#ff0000' });
    expect(actual[5]).toEqual({});
  } finally {
    engine.close();
  }
});
