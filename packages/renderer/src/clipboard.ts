import {
  OpenSheetError,
  LIMITS,
  key,
  parseInput,
  offsetFormula,
  type CellInput,
  type Command,
  type Range,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
import { parseClipboardHTML } from './html-clipboard.js';
import { encodeClipboardText } from './clipboard-text.js';

export type PasteOptions = {
  mode?: 'all' | 'values' | 'formulas' | 'formats';
  baseRevision?: number;
};
function pasteInput(input: CellInput, options: PasteOptions): CellInput {
  return options.mode === 'values' && input.type === 'formula'
    ? { type: 'string', value: '=' + input.expression }
    : input;
}
function pasteTextInput(text: string, options: PasteOptions, parsed?: CellInput): CellInput {
  return options.mode === 'values' && text.startsWith("'")
    ? { type: 'string', value: text }
    : pasteInput(parsed ?? parseInput(text), options);
}
function executePaste(book: Workbook, command: Command, options: PasteOptions) {
  if (
    options.mode !== undefined &&
    !['all', 'values', 'formulas', 'formats'].includes(options.mode)
  )
    throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown paste mode');
  return book.execute(
    options.baseRevision === undefined
      ? command
      : {
          ...command,
          protocolVersion: 1,
          commandId: crypto.randomUUID(),
          workbookId: book.id,
          baseRevision: options.baseRevision,
        },
  );
}
export function pasteClipboard(
  book: Workbook,
  selected: Selection,
  data: { text: string; html?: string },
  options: PasteOptions = {},
) {
  return data.html
    ? pasteHTML(book, selected, data.html, data.text, options)
    : pasteText(book, selected, data.text, options);
}

// Tile only when the complete selected rectangle is an exact multiple of the source.
function pasteDimensions(selected: Selection, sourceRows: number, sourceColumns: number) {
  const selectedRows = selected.endRow - selected.startRow;
  const selectedColumns = selected.endColumn - selected.startColumn;
  const tile =
    selectedRows >= sourceRows &&
    selectedColumns >= sourceColumns &&
    selectedRows % sourceRows === 0 &&
    selectedColumns % sourceColumns === 0;
  const rows = tile ? selectedRows : sourceRows;
  const columns = tile ? selectedColumns : sourceColumns;
  if (
    rows * columns > LIMITS.rangeCells ||
    selected.startRow + rows > LIMITS.rows ||
    selected.startColumn + columns > LIMITS.columns
  )
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste dimensions are too large');
  return { rows, columns };
}

// A small source can otherwise expand into gigabytes when a long value or note is tiled.
function guardPasteExpansion(
  dimensions: { rows: number; columns: number },
  sourceRows: number,
  sourceColumns: number,
  sourceUnits: number,
) {
  const copies = (dimensions.rows * dimensions.columns) / (sourceRows * sourceColumns);
  if (sourceUnits * copies > 32 * 1024 * 1024)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Expanded clipboard content is too large');
}

export async function pasteHTML(
  book: Workbook,
  selected: Selection,
  html: string,
  fallbackText = '',
  options: PasteOptions = {},
): Promise<Selection> {
  const target = { ...selected };
  const rows = parseClipboardHTML(html);
  if (!rows) return pasteText(book, target, fallbackText, options);
  const dimensions = pasteDimensions(target, rows.length, rows[0].length);
  let units = 0;
  for (const row of rows) for (const cell of row) units += JSON.stringify(cell).length + 32;
  guardPasteExpansion(dimensions, rows.length, rows[0].length, units);
  if (dimensions.rows !== rows.length || dimensions.columns !== rows[0].length) {
    const cells = rows.flatMap((row) =>
      row.map((cell) => {
        const input =
          options.mode === 'values' && cell.value
            ? cell.value
            : cell.externalText !== undefined
              ? pasteTextInput(cell.externalText, options, cell.input)
              : pasteInput(cell.input, options);
        if (options.mode === 'values' && input.type === 'error' && input.code === '#LIMIT!')
          throw new OpenSheetError('LIMIT_EXCEEDED', 'Cannot paste an incomplete calculation');
        return {
          ...(options.mode !== 'formats' ? { input } : {}),
          ...(options.mode !== 'formats' && input.type === 'formula' && cell.source
            ? { source: cell.source }
            : {}),
          ...((options.mode === undefined || options.mode === 'all') && cell.source
            ? { note: cell.note, link: cell.link }
            : {}),
          ...(options.mode === undefined || options.mode === 'all' || options.mode === 'formats'
            ? { style: cell.style }
            : {}),
        };
      }),
    );
    const range = {
      startRow: target.startRow,
      endRow: target.startRow + dimensions.rows,
      startColumn: target.startColumn,
      endColumn: target.startColumn + dimensions.columns,
    };
    await executePaste(
      book,
      {
        type: 'core.cells.pasteTile',
        payload: {
          sheetId: target.sheetId,
          range,
          rows: rows.length,
          columns: rows[0].length,
          cells,
        },
      },
      options,
    );
    return { ...target, ...range };
  }
  const cells = Array.from({ length: dimensions.rows }, (_, i) =>
    Array.from({ length: dimensions.columns }, (_, j) => {
      const cell = rows[i % rows.length][j % rows[0].length];
      let input =
        options.mode === 'values' && cell.value
          ? cell.value
          : cell.externalText !== undefined
            ? pasteTextInput(cell.externalText, options, cell.input)
            : pasteInput(cell.input, options);
      if (options.mode === 'values' && input.type === 'error' && input.code === '#LIMIT!')
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Cannot paste an incomplete calculation');
      if (options.mode !== 'formats' && input.type === 'formula' && cell.source)
        input = {
          type: 'formula',
          expression: offsetFormula(
            input.expression,
            target.startRow + i - cell.source.row,
            target.startColumn + j - cell.source.column,
          ),
        };
      return {
        row: target.startRow + i,
        column: target.startColumn + j,
        ...(options.mode !== 'formats' ? { input } : {}),
        ...((options.mode === undefined || options.mode === 'all') && cell.source
          ? { note: cell.note, link: cell.link }
          : {}),
        ...(options.mode === undefined || options.mode === 'all' || options.mode === 'formats'
          ? { style: cell.style }
          : {}),
      };
    }),
  ).flat();
  await executePaste(
    book,
    { type: 'core.cells.paste', payload: { sheetId: target.sheetId, cells } },
    options,
  );
  return {
    ...target,
    endRow: target.startRow + dimensions.rows,
    endColumn: target.startColumn + dimensions.columns,
  };
}
export async function copyText(range: Range) {
  const { workbook, id } = range.sheet;
  const bounds = { ...range.bounds };
  const revision = workbook.revision;
  const data = await workbook.readRange(id, bounds);
  if (data.revision !== revision || workbook.revision !== revision)
    throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed during copy');
  const sheet = workbook.sheetData(id);
  const rows: string[][] = [];
  for (let row = bounds.startRow; row < bounds.endRow; row++) {
    const values: string[] = [];
    for (let column = bounds.startColumn; column < bounds.endColumn; column++) {
      const cellKey = key(sheet.rowOrder[row], sheet.columnOrder[column]);
      const value = data.calculated[cellKey];
      if (value && typeof value === 'object' && value.error === '#LIMIT!')
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Cannot copy an incomplete calculation');
      values.push(data.display[cellKey] ?? '');
    }
    rows.push(values);
  }
  return encodeClipboardText(rows);
}

/** Clear only the captured source after the clipboard accepts its contents. */
export async function cutText(
  book: Workbook,
  range: Range,
  write: (text: string) => Promise<void>,
  canClear?: () => boolean,
) {
  const baseRevision = book.revision;
  const bounds = { ...range.bounds };
  const sheetId = range.sheet.id;
  const text = await copyText(range);
  await write(text);
  if (canClear && !canClear()) throw new OpenSheetError('READ_ONLY', 'Editor is read only');
  await book.execute({
    protocolVersion: 1,
    commandId: crypto.randomUUID(),
    workbookId: book.id,
    baseRevision,
    type: 'core.cells.clear',
    payload: { sheetId, range: bounds, all: false },
  });
}

export function parseClipboardText(text: string): string[][] {
  if (text.length > 2 * 1024 * 1024)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste is too large');
  const rows: string[][] = [];
  let row: string[] = [],
    cell = '',
    quoted = false,
    closedQuote = false,
    fieldStarted = false,
    width = 0;
  function appendCell() {
    row.push(cell);
    cell = '';
    fieldStarted = false;
    closedQuote = false;
    width = Math.max(width, row.length);
    if (width > LIMITS.columns || (rows.length + 1) * width > LIMITS.rangeCells)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste dimensions are too large');
  }
  function appendRow() {
    rows.push(row);
    row = [];
    if (rows.length > LIMITS.rows)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste has too many rows');
  }
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (closedQuote && c !== '\t' && c !== '\n' && c !== '\r')
      throw new OpenSheetError('INVALID_ARGUMENT', 'Unexpected text after closing quote');
    if (c === '"' && (quoted || (!cell && !fieldStarted))) {
      if (quoted && text[i + 1] === '"') {
        cell += '"';
        i++;
      } else {
        quoted = !quoted;
        if (!quoted) closedQuote = true;
        fieldStarted = true;
      }
    } else if (!quoted && (c === '\t' || c === '\n' || c === '\r')) {
      appendCell();
      if (c !== '\t') {
        appendRow();
        if (c === '\r' && text[i + 1] === '\n') i++;
      }
    } else cell += c;
    if (cell.length > LIMITS.text)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Pasted cell text is too long');
  }
  if (quoted) throw new OpenSheetError('INVALID_ARGUMENT', 'Unclosed quote');
  if (cell || row.length || fieldStarted || !rows.length) {
    appendCell();
    appendRow();
  }
  return rows;
}

export async function pasteText(
  book: Workbook,
  selected: Selection,
  text: string,
  options: PasteOptions = {},
): Promise<Selection> {
  const target = { ...selected };
  if (options.mode === 'formats')
    throw new OpenSheetError(
      'UNSUPPORTED_FEATURE',
      'Formats require an HTML table on the clipboard',
    );
  const rows = parseClipboardText(text);
  const width = rows.reduce((width, row) => Math.max(width, row.length), 0);
  const dimensions = pasteDimensions(target, rows.length, width);
  let units = rows.length * width * 32;
  for (const row of rows) for (const cell of row) units += JSON.stringify(cell).length;
  guardPasteExpansion(dimensions, rows.length, width, units);
  if (dimensions.rows !== rows.length || dimensions.columns !== width) {
    const cells = rows.flatMap((row) =>
      Array.from({ length: width }, (_, column) => ({
        input: pasteTextInput(row[column] ?? '', options),
      })),
    );
    const range = {
      startRow: target.startRow,
      endRow: target.startRow + dimensions.rows,
      startColumn: target.startColumn,
      endColumn: target.startColumn + dimensions.columns,
    };
    await executePaste(
      book,
      {
        type: 'core.cells.pasteTile',
        payload: { sheetId: target.sheetId, range, rows: rows.length, columns: width, cells },
      },
      options,
    );
    return { ...target, ...range };
  }
  const cells = Array.from({ length: dimensions.rows }, (_, i) =>
    Array.from({ length: dimensions.columns }, (_, j) => ({
      row: target.startRow + i,
      column: target.startColumn + j,
      input: pasteTextInput(rows[i % rows.length][j % width] ?? '', options),
    })),
  ).flat();
  await executePaste(
    book,
    { type: 'core.cells.paste', payload: { sheetId: target.sheetId, cells } },
    options,
  );
  return {
    ...target,
    endRow: target.startRow + dimensions.rows,
    endColumn: target.startColumn + dimensions.columns,
  };
}
