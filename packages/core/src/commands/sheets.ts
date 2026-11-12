import { filterColumns } from '../filter.js';
import { compareSortKeys, sortColumns, sortKey } from '../sort-values.js';
import { FormulaCalculation } from '../calculation.js';
import { intersects } from '../address.js';
import { checkBounds, validateSnapshot } from '../model.js';
import {
  assert,
  key,
  LIMITS,
  type CellRecord,
  type Command,
  type WorkbookSnapshot,
} from '../types.js';
import { offsetCellFormula, sheetFor, structuralReferences, transformAll } from './helpers.js';
export function sheetsCommand(s: WorkbookSnapshot, command: Command): void {
  const p = command.payload as any; // Validated before dispatch.
  if (command.type === 'core.sheet.add') {
    assert(s.sheets.length < LIMITS.sheets, 'LIMIT_EXCEEDED', 'Sheet limit reached');
    assert(
      !s.sheets.some(
        (sh) => sh.id === p.sheet.id || sh.name.toLowerCase() === p.sheet.name.toLowerCase(),
      ),
      'INVALID_ARGUMENT',
      'Duplicate sheet',
    );
    validateSnapshot({
      ...s,
      sheets: [...s.sheets, p.sheet],
      sheetOrder: [...s.sheetOrder, p.sheet.id],
    });
    s.sheets.push(p.sheet);
    s.sheetOrder.push(p.sheet.id);
    return;
  }
  const sh = sheetFor(s, p.sheetId);
  switch (command.type) {
    case 'core.sheet.remove':
      assert(
        s.sheets.length > 1 && s.sheets.some((x) => x.id !== sh.id && !x.hidden),
        'INVALID_ARGUMENT',
        'Cannot remove last visible sheet',
      );
      transformAll(
        s,
        structuralReferences(command, sh.name),
        undefined,
        undefined,
        (sheet) => sheet.id === sh.id,
      );
      s.sheets = s.sheets.filter((x) => x.id !== sh.id);
      s.sheetOrder = s.sheetOrder.filter((id) => id !== sh.id);
      break;
    case 'core.sheet.rename':
      assert(
        !s.sheets.some((x) => x.id !== sh.id && x.name.toLowerCase() === p.name.toLowerCase()),
        'INVALID_ARGUMENT',
        'Duplicate sheet name',
      );
      transformAll(s, structuralReferences(command, sh.name));
      sh.name = p.name;
      break;
    case 'core.sheet.freeze':
      assert(
        p.rows <= sh.rowOrder.length && p.columns <= sh.columnOrder.length,
        'INVALID_RANGE',
        'Invalid freeze',
      );
      sh.freeze = { rows: p.rows, columns: p.columns };
      break;
    case 'core.sheet.filter':
      assert(
        !p.filter ||
          (filterColumns(p.filter).every((column) => column < sh.columnOrder.length) &&
            ('column' in p.filter || p.filter.headerRows <= sh.rowOrder.length)),
        'INVALID_RANGE',
        'Invalid filter column',
      );
      sh.filter = p.filter;
      break;
    case 'core.sheet.sort': {
      checkBounds(sh, p.range);
      const keys = sortColumns(p, p.range);
      assert(
        !sh.merges.some((m) => intersects(m, p.range)),
        'INVALID_RANGE',
        'Unmerge before sorting',
      );
      const rows = Array.from(
        { length: p.range.endRow - p.range.startRow },
        (_, i) => p.range.startRow + i,
      );
      const calculator = new FormulaCalculation(
        (id) => sheetFor(s, id),
        (name) => s.sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase())?.id,
        s.dateSystem,
      );
      const values = new Map(
        rows.map((row) => [
          row,
          keys.map((key) => sortKey(calculator.value(sh.id, row, key.column))),
        ]),
      );
      calculator.clear();
      rows.sort((a, b) => compareSortKeys(values.get(a)!, values.get(b)!, keys) || a - b);
      const copied = rows.map((r) =>
        Array.from({ length: p.range.endColumn - p.range.startColumn }, (_, i) => {
          const x = sh.cells[key(sh.rowOrder[r], sh.columnOrder[p.range.startColumn + i])];
          return x ? (JSON.parse(JSON.stringify(x)) as CellRecord) : undefined;
        }),
      );
      copied.forEach((cells, i) =>
        cells.forEach((x) => {
          if (x?.input.type === 'formula')
            x.input.expression = offsetCellFormula(
              x.input.expression,
              p.range.startRow + i - rows[i],
              0,
            );
        }),
      );
      copied.forEach((cells, i) =>
        cells.forEach((x, j) => {
          const r = p.range.startRow + i,
            c = p.range.startColumn + j,
            k = key(sh.rowOrder[r], sh.columnOrder[c]);
          if (!x) {
            delete sh.cells[k];
            return;
          }
          x.rowId = sh.rowOrder[r];
          x.columnId = sh.columnOrder[c];
          delete x.cached;
          sh.cells[k] = x;
        }),
      );
      break;
    }
    default:
      assert(false, 'UNSUPPORTED_FEATURE', 'Unknown command');
  }
}
