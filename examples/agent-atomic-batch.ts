import { address, OpenSheetError, type CellInput, type Workbook } from '@opensheetjs/core';

/** Two explicit numeric edits on distinct targets, each with an observed right-hand neighbor. */
export async function prepareAtomicBatch(
  book: Workbook,
  edits: readonly [
    { sheetId: string; row: number; column: number; value: number },
    { sheetId: string; row: number; column: number; value: number },
  ],
) {
  const requested = structuredClone(edits);
  if (requested.some((edit) => !Number.isFinite(edit.value)))
    throw new Error('Values must be finite');
  const key = (edit: (typeof requested)[number], column: number) =>
    `${edit.sheetId}:${edit.row}:${column}`;
  const targets = new Set(requested.map((edit) => key(edit, edit.column)));
  if (
    targets.size !== requested.length ||
    requested.some((edit) => targets.has(key(edit, edit.column + 1)))
  )
    throw new Error('Targets and observed neighbors must be distinct');
  const revision = book.revision;
  const changes = [];
  for (const edit of requested) {
    const rows = book.getAxes(edit.sheetId, 'row', edit.row, 1);
    const columns = book.getAxes(edit.sheetId, 'column', edit.column, 2);
    const range = {
      startRow: edit.row,
      endRow: edit.row + 1,
      startColumn: edit.column,
      endColumn: edit.column + 2,
    };
    const before = await book.readCells(edit.sheetId, range, { fields: ['input'] });
    if (before.revision !== revision || book.revision !== revision)
      throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed while preparing the batch');
    changes.push({
      sheetId: edit.sheetId,
      range,
      cell: {
        row: edit.row,
        column: edit.column,
        input: { type: 'number' as const, value: edit.value },
        expected: { rowId: rows.ids[0], columnId: columns.ids[0] },
      },
      neighbor: structuredClone(before.cells[1].input as CellInput),
      summary: `${edit.sheetId}!${address(edit.row, edit.column)}: ${JSON.stringify(before.cells[0].input)} → ${JSON.stringify({ type: 'number', value: edit.value })}`,
    });
  }
  return { workbookId: book.id, revision, changes };
}

export async function applyAtomicBatch(
  book: Workbook,
  plan: Awaited<ReturnType<typeof prepareAtomicBatch>>,
) {
  const captured = structuredClone(plan);
  if (captured.workbookId !== book.id || captured.revision !== book.revision)
    throw new OpenSheetError(
      'REVISION_CONFLICT',
      'Workbook changed; inspect and prepare the batch again',
    );
  // No await occurs between the revision check and transaction start. The transaction
  // also refuses pending writes and excludes writes through the parent until it ends.
  let revision = captured.revision;
  const stop = book.onCommit((commit) => {
    revision = commit.revision;
  });
  try {
    await book.transaction({ label: 'Reviewed agent batch' }, async (tx) => {
      for (const change of captured.changes)
        await tx.execute({
          type: 'core.cells.set',
          payload: { sheetId: change.sheetId, cells: [change.cell] },
        });
    });
  } finally {
    stop();
  }
  const observed = [];
  try {
    for (const change of captured.changes) {
      const result = await book.readCells(change.sheetId, change.range, { fields: ['input'] });
      if (result.revision !== revision || book.revision !== revision)
        throw new Error('Workbook changed before verification');
      if (
        JSON.stringify(result.cells[0].input) !== JSON.stringify(change.cell.input) ||
        JSON.stringify(result.cells[1].input) !== JSON.stringify(change.neighbor)
      )
        throw new Error('The observed cells differ from the reviewed plan');
      observed.push(result.cells);
    }
  } catch (cause) {
    throw new Error('Batch may have committed, but verification failed; inspect before retrying', {
      cause,
    });
  }
  return { revision, observed };
}
