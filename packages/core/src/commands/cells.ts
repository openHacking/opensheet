import { offsetFormula } from '@opensheetjs/formula';
import { intersects } from '../address.js';
import { checkBounds } from '../model.js';
import {
  assert,
  inputSchema,
  key,
  type CellRecord,
  type Command,
  type Rect,
  type WorkbookSnapshot,
} from '../types.js';
import { cell, guardMerges, iterate, sheetFor, styleId } from './helpers.js';
export function cellsCommand(s: WorkbookSnapshot, command: Command): void {
  const p = command.payload as any; // Validated before dispatch.
  const sh = sheetFor(s, p.sheetId);
  switch (command.type) {
    case 'core.cells.replace':
      for (const cell of Object.values(sh.cells))
        if (cell.input.type === 'string' && cell.input.value.includes(p.query)) {
          const input = {
            type: 'string',
            value: cell.input.value.replaceAll(p.query, p.replacement),
          };
          assert(
            inputSchema.safeParse(input).success,
            'LIMIT_EXCEEDED',
            'Replacement exceeds the text limit',
          );
          cell.input = input as typeof cell.input;
        }
      break;
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
      iterate(sh, source, (r, c) => {
        const value = sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
        if (value) original.set(`${r}:${c}`, JSON.parse(JSON.stringify(value)));
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
        const value = JSON.parse(JSON.stringify(stored)) as CellRecord;
        value.rowId = sh.rowOrder[r];
        value.columnId = sh.columnOrder[c];
        if (value.input.type === 'formula')
          value.input.expression = offsetFormula(value.input.expression, r - sr, c - sc);
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
        const target = cell(sh, x.row, x.column);
        target.input = x.input;
        delete target.cached;
      }
      break;
    case 'core.cells.style':
      guardMerges(sh, p.range);
      iterate(sh, p.range, (r, c) => {
        const x = cell(sh, r, c);
        x.styleId = styleId(s, { ...(x.styleId ? s.styles[x.styleId] : {}), ...p.style });
        if (p.style.numberFormat !== undefined) x.numberFormat = p.style.numberFormat;
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
    case 'core.cells.merge':
      checkBounds(sh, p.range);
      assert(
        !sh.merges.some((m) => intersects(m, p.range)),
        'INVALID_RANGE',
        'Overlapping merged cells',
      );
      iterate(sh, p.range, (r, c) => {
        if (r === p.range.startRow && c === p.range.startColumn) return;
        const x = sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
        assert(
          !x || x.input.type === 'blank' || p.discardCoveredValues,
          'INVALID_ARGUMENT',
          'Merge would discard values',
        );
        if (x) x.input = { type: 'blank' };
      });
      sh.merges.push(p.range);
      break;
    case 'core.cells.unmerge':
      checkBounds(sh, p.range);
      sh.merges = sh.merges.filter((m) => !intersects(m, p.range));
      break;
    default:
      assert(false, 'UNSUPPORTED_FEATURE', 'Unknown command');
  }
}
