import { address, type Workbook } from '@opensheetjs/core';
import { streamExport, type DelimitedExportReport } from '@opensheetjs/formats';

/** Prepare one explicitly specified repair, keeping an adjacent input as a witness. */
export async function prepareFormulaRepair(
  book: Workbook,
  sheetId: string,
  row: number,
  column: number,
  replacement: string,
) {
  const revision = book.revision;
  const rows = book.getAxes(sheetId, 'row', row, 1);
  const columns = book.getAxes(sheetId, 'column', column, 2);
  const range = { startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 2 };
  const before = await book.readCells(sheetId, range, { fields: ['input', 'value'] });
  if (before.revision !== revision || book.revision !== revision)
    throw new Error('Workbook changed during inspection; prepare again');
  const original = before.cells[0].input;
  if (!original || original.type !== 'formula') throw new Error('Target is not a formula');
  const command = {
    protocolVersion: 1 as const,
    workbookId: book.id,
    baseRevision: revision,
    commandId: crypto.randomUUID(),
    type: 'core.cells.set',
    payload: {
      sheetId,
      cells: [
        {
          row,
          column,
          input: { type: 'formula' as const, expression: replacement },
          expected: { rowId: rows.ids[0], columnId: columns.ids[0], input: original },
        },
      ],
    },
  };
  return {
    command,
    range,
    witness: structuredClone(before.cells[1].input),
    summary: `${address(row, column)}: ${JSON.stringify(original.expression)} → ${JSON.stringify(replacement)}`,
  };
}

export async function applyFormulaRepair(
  book: Workbook,
  plan: Awaited<ReturnType<typeof prepareFormulaRepair>>,
) {
  const captured = structuredClone(plan);
  const commit = await book.execute(captured.command);
  const revision = commit?.revision ?? captured.command.baseRevision;
  let after: Awaited<ReturnType<Workbook['readCells']>>;
  try {
    after = await book.readCells(captured.command.payload.sheetId, captured.range, {
      fields: ['input', 'value'],
    });
  } catch (cause) {
    throw new Error('Repair may have committed, but verification could not complete', { cause });
  }
  const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
  if (after.revision !== revision || book.revision !== revision)
    throw new Error('Repair may have committed, but the workbook changed before verification');
  if (
    !equal(after.cells[0].input, captured.command.payload.cells[0].input) ||
    !equal(after.cells[1].input, captured.witness)
  )
    throw new Error('Repair committed, but verification failed');
  return { revision, cells: after.cells };
}

/** Stage a small CSV in memory and return it only if export completes at one revision. */
export async function exportCheckedCSV(book: Workbook, sheetId: string, range: string) {
  const revision = book.revision;
  let report: DelimitedExportReport | undefined;
  let text = '';
  for await (const chunk of streamExport(
    book,
    { sheetId, range, format: 'csv' },
    {
      onReport: (r) => {
        report = r;
      },
    },
  )) {
    if (text.length + chunk.length > 64 * 1024)
      throw new Error('Example export exceeds 64 Ki UTF-16 units');
    text += chunk;
  }
  if (!report || report.revision !== revision || book.revision !== revision)
    throw new Error('Workbook changed during export; discard the output');
  return { text, report };
}
