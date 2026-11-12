import { expect, it } from 'vitest';
import { createWorkbook } from '@opensheetjs/core';

it('changes only the requested identities, preserves sizes, and rejects out-of-bounds changes atomically', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.insertRows(2, 1);
    await sheet.insertColumns(2, 1);
    await sheet.setRowHeight(2, 42);
    await sheet.setColumnWidth(2, 140);
    await sheet.setRowsHidden(1, 3, true);
    await sheet.setColumnsHidden(1, 3, true);
    let data = book.sheetData(sheet.id);
    expect(data.rowOrder.slice(0, 5).map((id) => !!data.rows[id]?.hidden)).toEqual([
      false,
      true,
      true,
      true,
      false,
    ]);
    expect(data.columnOrder.slice(0, 5).map((id) => !!data.columns[id]?.hidden)).toEqual([
      false,
      true,
      true,
      true,
      false,
    ]);
    const revision = book.revision;
    await expect(sheet.setRowsHidden(8, 2, false)).rejects.toMatchObject({ code: 'INVALID_RANGE' });
    expect(book.revision).toBe(revision);
    await sheet.setRowsHidden(0, sheet.rowCount, false);
    await sheet.setColumnsHidden(0, sheet.columnCount, false);
    data = book.sheetData(sheet.id);
    expect(data.rows[data.rowOrder[2]]).toMatchObject({ size: 42, hidden: false });
    expect(data.columns[data.columnOrder[2]]).toMatchObject({ size: 140, hidden: false });
    expect(data.rows[data.rowOrder[0]]).toBeUndefined();
    expect(data.columns[data.columnOrder[0]]).toBeUndefined();
    await book.undo();
    expect(book.sheetData(sheet.id).columns[data.columnOrder[2]].hidden).toBe(true);
    await book.redo();
    expect(book.sheetData(sheet.id).columns[data.columnOrder[2]].hidden).toBe(false);
  } finally {
    await book.deleteStorage();
  }
});

it('resizes row and column ranges atomically while preserving visibility and axis identity', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.insertRows(2);
    await sheet.insertColumns(2);
    await sheet.setRowHidden(2, true);
    await sheet.setColumnHidden(2, true);
    await sheet.setRowsHeight(1, 3, 52.5);
    await sheet.setColumnsWidth(1, 3, 174);
    const data = book.sheetData(sheet.id);
    expect(data.rowOrder.slice(1, 4).map((id) => data.rows[id].size)).toEqual([52.5, 52.5, 52.5]);
    expect(data.columnOrder.slice(1, 4).map((id) => data.columns[id].size)).toEqual([
      174, 174, 174,
    ]);
    expect(data.rows[data.rowOrder[2]].hidden).toBe(true);
    expect(data.columns[data.columnOrder[2]].hidden).toBe(true);
    const revision = book.revision;
    await expect(sheet.setRowsHeight(8, 2, 30)).rejects.toMatchObject({ code: 'INVALID_RANGE' });
    await expect(sheet.setColumnsWidth(0, 1, 2001)).rejects.toMatchObject({
      code: 'INVALID_ARGUMENT',
    });
    expect(book.revision).toBe(revision);
    await book.undo();
    expect(book.sheetData(sheet.id).columns[data.columnOrder[2]]).toEqual({ hidden: true });
    await book.redo();
    expect(book.sheetData(sheet.id).columns[data.columnOrder[2]]).toEqual({
      hidden: true,
      size: 174,
    });
  } finally {
    await book.deleteStorage();
  }
});

it('keeps revision and redo for standalone visibility no-ops, while validating their targets', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.setRowsHidden(2, 1, true);
    await book.undo();
    const revision = book.revision;
    await book.prefetch(sheet.id, sheet.range('A1').bounds);
    expect(book.isLoaded(sheet.id, 0, 0)).toBe(true);
    expect(await sheet.setRowsHidden(0, 8, false)).toBeUndefined();
    expect(await sheet.setColumnsHidden(0, 8, false)).toBeUndefined();
    expect(book.isLoaded(sheet.id, 0, 0)).toBe(true);
    expect(book.revision).toBe(revision);
    const envelope = {
      protocolVersion: 1 as const,
      commandId: 'visibility-noop',
      workbookId: book.id,
      baseRevision: revision,
      type: 'core.axis.visibility',
      payload: { sheetId: sheet.id, axis: 'row', index: 0, count: 8, hidden: false },
    };
    expect(await book.execute(envelope)).toBeUndefined();
    expect(await book.execute(envelope)).toBeUndefined();
    await expect(
      book.execute({ ...envelope, payload: { ...envelope.payload, hidden: true } }),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    await expect(sheet.setRowsHidden(8, 1, false)).rejects.toMatchObject({ code: 'INVALID_RANGE' });
    await expect(
      book.execute({
        type: 'core.axis.visibility',
        payload: { sheetId: 'missing', axis: 'row', index: 0, count: 1, hidden: false },
      }),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    await book.redo();
    expect(await book.execute(envelope)).toBeUndefined();
    const data = book.sheetData(sheet.id);
    expect(data.rows[data.rowOrder[2]].hidden).toBe(true);
    const hiddenRevision = book.revision;
    expect(await sheet.setRowsHidden(2, 1, true)).toBeUndefined();
    expect(book.revision).toBe(hiddenRevision);
    await book.undo();
    expect(!!book.sheetData(sheet.id).rows[data.rowOrder[2]]?.hidden).toBe(false);
  } finally {
    await book.deleteStorage();
  }
});

it('resolves all-axis visibility after a structural command in the same transaction', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.setRowsHidden(7, 1, true);
    await sheet.setColumnsHidden(7, 1, true);
    await book.transaction({}, async (tx) => {
      const scoped = tx.getSheets()[0];
      await scoped.insertRows(0);
      await scoped.insertColumns(0);
      await scoped.setAllRowsHidden(false);
      await scoped.setAllColumnsHidden(false);
    });
    let data = book.sheetData(sheet.id);
    expect(data.rows[data.rowOrder[8]].hidden).toBe(false);
    expect(data.columns[data.columnOrder[8]].hidden).toBe(false);
    await book.undo();
    data = book.sheetData(sheet.id);
    expect(data.rowOrder).toHaveLength(8);
    expect(data.columnOrder).toHaveLength(8);
    expect(data.rows[data.rowOrder[7]].hidden).toBe(true);
    expect(data.columns[data.columnOrder[7]].hidden).toBe(true);
    await book.redo();
    data = book.sheetData(sheet.id);
    expect(data.rows[data.rowOrder[8]].hidden).toBe(false);
    expect(data.columns[data.columnOrder[8]].hidden).toBe(false);
  } finally {
    await book.deleteStorage();
  }
});
