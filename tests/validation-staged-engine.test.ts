import { expect, it, vi } from 'vitest';
import { Engine } from '../packages/core/src/engine.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { Database } from '../packages/core/src/storage.js';
import { key } from '../packages/core/src/types.js';

it('validates staged and committed views independently before publication', async () => {
  const snapshot = createSnapshot({
    sheets: [
      { name: 'Data', rows: 2, columns: 1 },
      { name: 'Limits', rows: 1, columns: 1 },
    ],
  });
  const [data, limits] = snapshot.sheets;
  data.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 5 } };
  limits.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 10 } };
  const db = await Database.open(`validation-staged-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const internal = engine as any;
  const before = engine.view();
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
  const rules = [
    {
      id: 'minimum',
      range,
      rule: { kind: 'number', operator: 'gte', first: 10, allowBlank: false, policy: 'reject' },
    },
    {
      id: 'limit',
      range,
      rule: { kind: 'formula', expression: 'A1<Limits!$A$1', allowBlank: false, policy: 'warn' },
    },
  ];
  const positions = [{ row: 0, column: 0 }];
  const publish = internal.publish.bind(engine);
  let checked = false;
  const spy = vi.spyOn(internal, 'publish').mockImplementation(async (next, metadata, ranges) => {
    const old = await internal.validateStagedCells(
      internal.head,
      internal.metadata,
      data.id,
      rules,
      positions,
    );
    const staged = await internal.validateStagedCells(next, metadata, data.id, rules, positions);
    expect(old).toMatchObject({ rejected: 1, warnings: 0, details: [{ ruleId: 'minimum' }] });
    expect(staged).toMatchObject({ rejected: 0, warnings: 1, details: [{ ruleId: 'limit' }] });
    await expect(
      internal.validateStagedCells(next, metadata, data.id, rules, [{ row: 2, column: 0 }]),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    const controller = new AbortController();
    controller.abort();
    await expect(
      internal.validateStagedCells(next, metadata, data.id, rules, positions, controller.signal),
    ).rejects.toMatchObject({ name: 'AbortError' });
    expect(engine.view().revision).toBe(before.revision);
    expect((await engine.cell(data.id, 0, 0))?.input).toEqual({ type: 'number', value: 5 });
    checked = true;
    return publish(next, metadata, ranges);
  });
  try {
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
    expect(checked).toBe(true);
    expect(engine.view().revision).toBe(before.revision + 1);
    expect((await engine.cell(data.id, 0, 0))?.input).toEqual({ type: 'number', value: 12 });
  } finally {
    spy.mockRestore();
    engine.close();
  }
});

it('cleans staged writes when a caller rejects a validation report before publication', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 2, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 5 } };
  const name = `validation-reject-${crypto.randomUUID()}`;
  let db = await Database.open(name);
  let engine = new Engine(db, snapshot.workbookId, 0);
  await engine.open(snapshot);
  const before = engine.view();
  const internal = engine as any;
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
  const rules = [
    {
      id: 'maximum',
      range,
      rule: { kind: 'number', operator: 'lte', first: 10, allowBlank: false, policy: 'reject' },
    },
  ];
  const spy = vi.spyOn(internal, 'publish').mockImplementation(async (next, metadata) => {
    const report = await internal.validateStagedCells(next, metadata, sheet.id, rules, [
      { row: 0, column: 0 },
    ]);
    expect(report.rejected).toBe(1);
    throw new Error('Injected validation rejection');
  });
  try {
    await expect(
      engine.execute([
        {
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [{ row: 0, column: 0, input: { type: 'number', value: 12 } }],
          },
        },
      ]),
    ).rejects.toThrow('Injected validation rejection');
    spy.mockRestore();
    expect(engine.view()).toMatchObject({
      revision: before.revision,
      bytes: before.bytes,
      canUndo: before.canUndo,
    });
    expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({ type: 'number', value: 5 });
    engine.close();
    db = await Database.open(name);
    engine = new Engine(db, snapshot.workbookId, 0);
    await engine.open();
    expect(engine.view()).toMatchObject({
      revision: before.revision,
      bytes: before.bytes,
      canUndo: before.canUndo,
    });
    expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({ type: 'number', value: 5 });
  } finally {
    spy.mockRestore();
    engine.close();
  }
});
