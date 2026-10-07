import { assert, LIMITS, type Command, type WorkbookSnapshot } from '../types.js';
import { sheetFor, transformAll } from './helpers.js';
export function axisCommand(s: WorkbookSnapshot, command: Command): void {
  const p = command.payload as any; // Validated before dispatch.
  const sh = sheetFor(s, p.sheetId);
  switch (command.type) {
    case 'core.axis.insert':
    case 'core.axis.delete': {
      const row = p.axis === 'row',
        order = row ? sh.rowOrder : sh.columnOrder,
        insert = command.type === 'core.axis.insert',
        count = insert ? p.ids.length : p.count;
      assert(
        p.index <= (insert ? order.length : order.length - count) && count > 0,
        'INVALID_RANGE',
        'Invalid structural range',
      );
      assert(
        insert
          ? order.length + count <= (row ? LIMITS.rows : LIMITS.columns)
          : order.length - count >= 1,
        'LIMIT_EXCEEDED',
        'Invalid sheet size',
      );
      if (insert)
        assert(
          new Set([...order, ...p.ids]).size === order.length + count,
          'INVALID_ARGUMENT',
          'Duplicate axis IDs',
        );
      const start = row ? 'startRow' : 'startColumn',
        end = row ? 'endRow' : 'endColumn';
      for (const m of sh.merges)
        assert(
          !(p.index < m[end] && p.index + (!insert ? count : 0) > m[start]),
          'INVALID_RANGE',
          'Unmerge cells before editing this axis',
        );
      transformAll(
        s,
        (r) => {
          if (r.sheet !== sh.name) return r;
          const at = row ? r.row : r.column;
          if (!insert && at >= p.index && at < p.index + count) return null;
          const changed =
            at >= p.index ? at + (insert ? count : at >= p.index + count ? -count : 0) : at;
          return { ...r, [row ? 'row' : 'column']: changed };
        },
        sh.name,
      );
      const deleted = insert ? [] : order.slice(p.index, p.index + count);
      order.splice(p.index, insert ? 0 : count, ...(insert ? p.ids : []));
      for (const k of Object.keys(sh.cells))
        if (deleted.includes(row ? sh.cells[k].rowId : sh.cells[k].columnId)) delete sh.cells[k];
      for (const id of deleted) delete (row ? sh.rows : sh.columns)[id];
      sh.merges.forEach((m) => {
        if (m[start] >= p.index) {
          m[start] += insert ? count : -count;
          m[end] += insert ? count : -count;
        }
      });
      const f = row ? 'rows' : 'columns';
      sh.freeze[f] = Math.min(
        order.length,
        sh.freeze[f] > p.index
          ? Math.max(p.index, sh.freeze[f] + (insert ? count : -count))
          : sh.freeze[f],
      );
      if (!row && sh.filter) {
        if (!insert && sh.filter.column >= p.index && sh.filter.column < p.index + count)
          sh.filter = null;
        else if (sh.filter.column >= p.index) sh.filter.column += insert ? count : -count;
      }
      break;
    }
    case 'core.axis.meta': {
      const order = p.axis === 'row' ? sh.rowOrder : sh.columnOrder;
      assert(p.index < order.length, 'INVALID_RANGE', 'Unknown row/column');
      const meta = p.axis === 'row' ? sh.rows : sh.columns;
      meta[order[p.index]] = {
        ...meta[order[p.index]],
        ...(p.size !== undefined ? { size: p.size } : {}),
        ...(p.hidden !== undefined ? { hidden: p.hidden } : {}),
      };
      break;
    }
    default:
      assert(false, 'UNSUPPORTED_FEATURE', 'Unknown command');
  }
}
