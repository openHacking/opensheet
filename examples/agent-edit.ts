import { address, type CellInput, type Workbook } from '@opensheetjs/core';

/** A deliberately small, reviewable operation: change two cells and observe a third. */
export async function prepareTwoCellEdit(
  book: Workbook,
  sheetId: string,
  row: number,
  column: number,
  values: readonly [number, number],
) {
  if (!values.every(Number.isFinite)) throw new Error('Values must be finite numbers');
  const revision = book.revision;
  const rowAxis = book.getAxes(sheetId, 'row', row, 1);
  const columnAxis = book.getAxes(sheetId, 'column', column, 3);
  const range = { startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 3 };
  const before = await book.readCells(sheetId, range, { fields: ['input'] });
  if (before.revision !== revision || book.revision !== revision)
    throw new Error('Workbook changed while preparing the edit; inspect again');
  const cells = values.map((value, offset) => ({
    row,
    column: column + offset,
    input: { type: 'number' as const, value },
    expected: { rowId: rowAxis.ids[0], columnId: columnAxis.ids[offset] },
  }));
  const command = {
    protocolVersion: 1 as const,
    workbookId: book.id,
    baseRevision: revision,
    commandId: crypto.randomUUID(),
    type: 'core.cells.set',
    payload: { sheetId, cells },
  };
  return {
    command,
    range,
    guardInput: structuredClone(before.cells[2].input as CellInput),
    summary: cells.map(
      (cell, i) =>
        `${address(row, column + i)}: ${JSON.stringify(before.cells[i].input)} → ${JSON.stringify(cell.input)}`,
    ),
  };
}

export async function applyTwoCellEdit(
  book: Workbook,
  plan: Awaited<ReturnType<typeof prepareTwoCellEdit>>,
) {
  const captured = structuredClone(plan);
  const commit = await book.execute(captured.command);
  const revision = commit?.revision ?? captured.command.baseRevision;
  let after: Awaited<ReturnType<Workbook['readCells']>>;
  try {
    after = await book.readCells(captured.command.payload.sheetId, captured.range, {
      fields: ['input'],
    });
  } catch (cause) {
    throw new Error('Edit may have committed, but verification could not complete', { cause });
  }
  if (after.revision !== revision || book.revision !== revision)
    throw new Error('Edit may have committed, but the workbook changed before verification');
  const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
  if (
    !captured.command.payload.cells.every((cell, i) => same(cell.input, after.cells[i].input)) ||
    !same(captured.guardInput, after.cells[2].input)
  )
    throw new Error('Edit committed, but verification failed');
  return { revision, cells: after.cells };
}
