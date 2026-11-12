import { assert, LIMITS, type Command, type WorkbookSnapshot } from '../types.js';
import {
  sheetFor,
  structuralRangeReferences,
  structuralReferences,
  transformAll,
} from './helpers.js';
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
      if (insert) {
        const next = row ? 'nextRowId' : 'nextColumnId';
        assert(
          p.ids.every((id: number, i: number) => id === sh[next] + i) &&
            sh[next] + count <= 0x100000000,
          'INVALID_ARGUMENT',
          'Axis IDs must be allocated monotonically',
        );
      }
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
      const deleted = insert ? [] : order.slice(p.index, p.index + count);
      const removed = new Set(deleted);
      transformAll(
        s,
        structuralReferences(command, sh.name),
        sh.name,
        structuralRangeReferences(command, sh.name),
        insert
          ? undefined
          : (sheet, cell) => sheet.id === sh.id && removed.has(row ? cell.rowId : cell.columnId),
      );
      if (insert) sh[row ? 'nextRowId' : 'nextColumnId'] += count;
      if (insert) {
        // Avoid exceeding engine-specific argument limits for large inserts.
        for (let i = 0; i < p.ids.length; i += 8192)
          order.splice(p.index + i, 0, ...p.ids.slice(i, i + 8192));
      } else {
        order.splice(p.index, count);
        for (const k of Object.keys(sh.cells))
          if (removed.has(row ? sh.cells[k].rowId : sh.cells[k].columnId)) delete sh.cells[k];
      }
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
      if (sh.filter) {
        const filter = sh.filter;
        if (!row) {
          const removed = (column: number) =>
            !insert && column >= p.index && column < p.index + count;
          const shifted = (column: number) =>
            column >= p.index ? column + (insert ? count : -count) : column;
          if ('column' in filter) {
            if (removed(filter.column)) sh.filter = null;
            else filter.column = shifted(filter.column);
          } else {
            filter.criteria = filter.criteria.filter((criterion) => !removed(criterion.column));
            for (const criterion of filter.criteria) criterion.column = shifted(criterion.column);
            if (!filter.criteria.length) sh.filter = null;
          }
        } else if ('headerRows' in filter && p.index < filter.headerRows) {
          filter.headerRows = insert
            ? filter.headerRows + count
            : Math.max(p.index, filter.headerRows - count);
        }
      }
      break;
    }
    case 'core.axis.resize': {
      const order = p.axis === 'row' ? sh.rowOrder : sh.columnOrder;
      assert(p.index + p.count <= order.length, 'INVALID_RANGE', 'Unknown row/column');
      const meta = p.axis === 'row' ? sh.rows : sh.columns;
      for (let i = p.index; i < p.index + p.count; i++) {
        const id = order[i];
        if (meta[id]?.size === p.size) continue;
        meta[id] = { ...meta[id], size: p.size };
      }
      break;
    }
    case 'core.axis.visibility':
    case 'core.axis.visibilityAll': {
      const order = p.axis === 'row' ? sh.rowOrder : sh.columnOrder;
      const start = command.type === 'core.axis.visibilityAll' ? 0 : p.index;
      const end = command.type === 'core.axis.visibilityAll' ? order.length : p.index + p.count;
      assert(end <= order.length, 'INVALID_RANGE', 'Unknown row/column');
      const meta = p.axis === 'row' ? sh.rows : sh.columns;
      for (let i = start; i < end; i++) {
        const id = order[i];
        if (!p.hidden && !meta[id]?.hidden) continue;
        if (p.hidden && meta[id]?.hidden) continue;
        meta[id] = { ...meta[id], hidden: p.hidden };
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
