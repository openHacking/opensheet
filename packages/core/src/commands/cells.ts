import { numericSeries } from '../numeric-series.js';
import { createTileOffsetter, tileCell, validatePasteTile, type PasteTile } from '../paste-tile.js';
import { calendarSeries } from '../calendar-series.js';
import { FormulaCalculation } from '../calculation.js';
import { selectiveCopyCell } from '../selective-copy.js';
import { validateCommand } from './schemas.js';
import { intersects } from '../address.js';
import { checkBounds } from '../model.js';
import {
  assert,
  LIMITS,
  key,
  type CellInput,
  type CellRecord,
  type Command,
  type Rect,
  type WorkbookSnapshot,
} from '../types.js';
import {
  applyPastedCell,
  cell,
  guardMerges,
  iterate,
  offsetCellFormula,
  sheetFor,
  styleId,
} from './helpers.js';
function sameInput(a: CellInput | null, b: CellInput | null): boolean {
  if (a === null || b === null) return a === b;
  if (a.type !== b.type) return false;
  switch (a.type) {
    case 'blank':
      return true;
    case 'formula':
      return b.type === 'formula' && a.expression === b.expression;
    case 'error':
      return b.type === 'error' && a.code === b.code;
    case 'number':
      return b.type === 'number' && Object.is(a.value, b.value);
    default:
      return 'value' in b && a.value === b.value;
  }
}
export function cellsCommand(s: WorkbookSnapshot, command: Command): void {
  const p = command.payload as any; // Validated before dispatch.
  const sh = sheetFor(s, p.sheetId);
  switch (command.type) {
    case 'core.cells.copySpecial': {
      const source = p.range as Rect;
      const target = p.target as Rect;
      checkBounds(sh, source);
      checkBounds(sh, target);
      const rows = source.endRow - source.startRow;
      const columns = source.endColumn - source.startColumn;
      assert(
        rows === target.endRow - target.startRow &&
          columns === target.endColumn - target.startColumn,
        'INVALID_RANGE',
        'Copy ranges must have equal dimensions',
      );
      assert(rows * columns <= LIMITS.rangeCells, 'LIMIT_EXCEEDED', 'Copy exceeds the cell limit');
      assert(
        !sh.merges.some((merge) => intersects(merge, source) || intersects(merge, target)),
        'INVALID_RANGE',
        'Unmerge before copying',
      );
      const calculator = new FormulaCalculation(
        (id) => sheetFor(s, id),
        (name) => s.sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase())?.id,
        s.dateSystem,
      );
      const items = [];
      let bytes = 0;
      for (let row = source.startRow; row < source.endRow; row++)
        for (let column = source.startColumn; column < source.endColumn; column++) {
          const original = sh.cells[key(sh.rowOrder[row], sh.columnOrder[column])];
          const style = original?.styleId ? s.styles[original.styleId] : {};
          const item = selectiveCopyCell(
            {
              row,
              column,
              ...(p.mode === 'values' ? { value: calculator.value(sh.id, row, column) } : {}),
              ...(p.mode === 'formulas' ? { input: original?.input } : {}),
              ...(p.mode === 'formats'
                ? { style, numberFormat: original?.numberFormat ?? style?.numberFormat }
                : {}),
            },
            p.mode,
            target.startRow - source.startRow,
            target.startColumn - source.startColumn,
          );
          bytes += JSON.stringify(item).length * 2;
          assert(bytes <= 32 * 1024 * 1024, 'LIMIT_EXCEEDED', 'Copy working set is too large');
          items.push(item);
        }
      const paste = validateCommand({
        type: 'core.cells.paste',
        payload: { sheetId: sh.id, cells: items },
      });
      for (const item of paste.payload.cells as typeof items)
        applyPastedCell(s, cell(sh, item.row, item.column), item);
      break;
    }
    case 'core.cells.pasteTile': {
      const tile = p as PasteTile;
      let units = validatePasteTile(sh, tile);
      const charge = (extra: number) => {
        units += extra;
        assert(
          units <= 32 * 1024 * 1024,
          'LIMIT_EXCEEDED',
          'Expanded clipboard content is too large',
        );
      };
      const offset = createTileOffsetter();
      // Relocation can grow a formula or exceed the total clipboard budget.
      // Check every destination before creating or changing any cell.
      if (tile.cells.some((item) => item.source && item.input?.type === 'formula'))
        iterate(sh, tile.range, (row, column) => {
          tileCell(tile, row, column, charge, offset);
        });
      iterate(sh, tile.range, (row, column) => {
        const item = tileCell(tile, row, column, () => {}, offset);
        applyPastedCell(s, cell(sh, row, column), item);
      });
      break;
    }
    case 'core.cells.paste':
      for (const item of p.cells) {
        checkBounds(sh, {
          startRow: item.row,
          endRow: item.row + 1,
          startColumn: item.column,
          endColumn: item.column + 1,
        });
        assert(
          !sh.merges.some(
            (m) =>
              item.row >= m.startRow &&
              item.row < m.endRow &&
              item.column >= m.startColumn &&
              item.column < m.endColumn,
          ),
          'INVALID_RANGE',
          'Unmerge before pasting cell inputs or formats',
        );
      }
      for (const item of p.cells) applyPastedCell(s, cell(sh, item.row, item.column), item);
      break;
    case 'core.cells.replace': {
      const growth = p.replacement.length - p.query.length;
      if (growth > 0)
        for (const cell of Object.values(sh.cells))
          if (cell.input.type === 'string' && cell.input.value.includes(p.query)) {
            const original = cell.input.value;
            let length = original.length;
            let position = 0;
            while ((position = original.indexOf(p.query, position)) !== -1) {
              length += growth;
              assert(length <= LIMITS.text, 'LIMIT_EXCEEDED', 'Replacement exceeds the text limit');
              position += p.query.length;
            }
          }
      for (const cell of Object.values(sh.cells))
        if (cell.input.type === 'string' && cell.input.value.includes(p.query)) {
          const original = cell.input.value;
          const input = {
            type: 'string',
            value: original.replaceAll(p.query, () => p.replacement),
          };
          cell.input = input as typeof cell.input;
        }
      break;
    }
    case 'core.cells.fillSeries': {
      checkBounds(sh, p.range);
      assert(
        !sh.merges.some((m) => intersects(m, p.range)),
        'INVALID_RANGE',
        'Unmerge before filling a series',
      );
      const count =
        p.axis === 'row'
          ? p.range.endRow - p.range.startRow
          : p.range.endColumn - p.range.startColumn;
      const values = p.calendar
        ? calendarSeries(p.start, p.step, count, p.calendar, s.dateSystem)
        : numericSeries(p.start, p.step, count);
      iterate(sh, p.range, (r, c) => {
        const target = cell(sh, r, c);
        target.input = {
          type: 'number',
          value: values[p.axis === 'row' ? r - p.range.startRow : c - p.range.startColumn],
        };
        delete target.cached;
      });
      break;
    }
    case 'core.cells.copy':
    case 'core.cells.fillDown': {
      const source = p.range as Rect;
      const target = (command.type === 'core.cells.fillDown' ? p.range : p.target) as Rect;
      checkBounds(sh, source);
      checkBounds(sh, target);
      assert(
        !sh.merges.some((m) => intersects(m, source) || intersects(m, target)),
        'INVALID_RANGE',
        'Unmerge before copying or filling',
      );
      const fill = command.type === 'core.cells.fillDown';
      assert(
        fill ||
          (source.endRow - source.startRow === target.endRow - target.startRow &&
            source.endColumn - source.startColumn === target.endColumn - target.startColumn),
        'INVALID_RANGE',
        'Copy ranges must have equal dimensions',
      );
      const original = new Map<string, CellRecord>();
      let hasFormula = false;
      const snapshotRange = fill ? { ...source, endRow: source.startRow + 1 } : source;
      iterate(sh, snapshotRange, (r, c) => {
        const value = sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
        if (value) {
          original.set(`${r}:${c}`, structuredClone(value));
          hasFormula ||= value.input.type === 'formula';
        }
      });
      if (hasFormula)
        iterate(sh, target, (r, c) => {
          if (fill && r === source.startRow) return;
          const sr = fill ? source.startRow : source.startRow + r - target.startRow;
          const sc = source.startColumn + c - target.startColumn;
          const stored = original.get(`${sr}:${sc}`);
          if (stored?.input.type === 'formula')
            offsetCellFormula(stored.input.expression, r - sr, c - sc);
        });
      iterate(sh, target, (r, c) => {
        if (fill && r === source.startRow) return;
        const sr = fill ? source.startRow : source.startRow + r - target.startRow;
        const sc = source.startColumn + c - target.startColumn;
        const stored = original.get(`${sr}:${sc}`);
        const k = key(sh.rowOrder[r], sh.columnOrder[c]);
        if (!stored) {
          delete sh.cells[k];
          return;
        }
        const value = structuredClone(stored) as CellRecord;
        value.rowId = sh.rowOrder[r];
        value.columnId = sh.columnOrder[c];
        if (value.input.type === 'formula')
          value.input.expression = offsetCellFormula(value.input.expression, r - sr, c - sc);
        delete value.cached;
        sh.cells[k] = value;
      });
      break;
    }
    case 'core.cells.set':
      for (const x of p.cells) {
        checkBounds(sh, {
          startRow: x.row,
          startColumn: x.column,
          endRow: x.row + 1,
          endColumn: x.column + 1,
        });
        assert(
          !x.expected ||
            (sh.rowOrder[x.row] === x.expected.rowId &&
              sh.columnOrder[x.column] === x.expected.columnId),
          'REVISION_CONFLICT',
          'Cell position changed; select it again before editing',
        );
        if (x.expected?.input !== undefined) {
          const current =
            sh.cells[key(sh.rowOrder[x.row], sh.columnOrder[x.column])]?.input ?? null;
          assert(
            sameInput(current, x.expected.input),
            'REVISION_CONFLICT',
            'Cell content changed; review the draft before saving',
          );
        }
        const m = sh.merges.find(
          (m) =>
            x.row >= m.startRow &&
            x.row < m.endRow &&
            x.column >= m.startColumn &&
            x.column < m.endColumn,
        );
        assert(
          !m || (x.row === m.startRow && x.column === m.startColumn),
          'INVALID_RANGE',
          'Cannot write a covered merged cell',
        );
      }
      // Preconditions describe the state at the beginning of this command,
      // including when more than one entry names the same coordinate.
      for (const x of p.cells) {
        const target = cell(sh, x.row, x.column);
        target.input = x.input;
        delete target.cached;
      }
      break;
    case 'core.cells.annotation':
      checkBounds(sh, p.range);
      for (const merge of sh.merges) {
        const startRow = Math.max(p.range.startRow, merge.startRow);
        const endRow = Math.min(p.range.endRow, merge.endRow);
        const startColumn = Math.max(p.range.startColumn, merge.startColumn);
        const endColumn = Math.min(p.range.endColumn, merge.endColumn);
        assert(
          startRow >= endRow ||
            startColumn >= endColumn ||
            (startRow === merge.startRow &&
              startColumn === merge.startColumn &&
              endRow === startRow + 1 &&
              endColumn === startColumn + 1),
          'INVALID_RANGE',
          'Cannot annotate a covered merged cell',
        );
      }
      iterate(sh, p.range, (r, c) => {
        if (!sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])] && p.note == null && p.link == null)
          return;
        const target = cell(sh, r, c);
        if (p.note !== undefined) {
          if (p.note === null) delete target.note;
          else target.note = p.note;
        }
        if (p.link !== undefined) {
          if (p.link === null) delete target.link;
          else target.link = { ...p.link };
        }
      });
      break;
    case 'core.cells.style':
      guardMerges(sh, p.range);
      iterate(sh, p.range, (r, c) => {
        const x = cell(sh, r, c);
        x.styleId = styleId(s, { ...(x.styleId ? s.styles[x.styleId] : {}), ...p.style });
        // Style formats live in the shared style table, not in every cell.
        if (p.style.numberFormat !== undefined) delete x.numberFormat;
      });
      break;
    case 'core.cells.clear':
      guardMerges(sh, p.range);
      iterate(sh, p.range, (r, c) => {
        const k = key(sh.rowOrder[r], sh.columnOrder[c]);
        if (p.all) delete sh.cells[k];
        else if (sh.cells[k]) {
          sh.cells[k].input = { type: 'blank' };
          delete sh.cells[k].cached;
        }
      });
      break;
    case 'core.cells.merge': {
      checkBounds(sh, p.range);
      assert(
        !sh.merges.some((m) => intersects(m, p.range)),
        'INVALID_RANGE',
        'Overlapping merged cells',
      );
      const covered: CellRecord[] = [];
      iterate(sh, p.range, (r, c) => {
        if (r === p.range.startRow && c === p.range.startColumn) return;
        const x = sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
        assert(
          !x || x.input.type === 'blank' || p.discardCoveredValues,
          'INVALID_ARGUMENT',
          'Merge would discard values',
        );
        if (x) covered.push(x);
      });
      for (const x of covered) {
        x.input = { type: 'blank' };
        delete x.cached;
      }
      sh.merges.push(p.range);
      break;
    }
    case 'core.cells.unmerge':
      checkBounds(sh, p.range);
      sh.merges = sh.merges.filter((m) => !intersects(m, p.range));
      break;
    default:
      assert(false, 'UNSUPPORTED_FEATURE', 'Unknown command');
  }
}
