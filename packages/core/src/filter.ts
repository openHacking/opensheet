import type { CellValue } from './types.js';
import type { FilterCriterion, WorksheetFilter } from './filter-schema.js';

export function filterColumns(filter: NonNullable<WorksheetFilter>): number[] {
  return 'column' in filter
    ? [filter.column]
    : filter.criteria.map((criterion) => criterion.column);
}

/** Typed predicates never coerce numeric text, booleans or formula errors into numbers. */
export function matchesFilter(
  value: CellValue,
  criterion: FilterCriterion,
  foldedQuery?: string,
): boolean {
  const error = value !== null && typeof value === 'object';
  const blank = value === null || value === '';
  switch (criterion.operator) {
    case 'isBlank':
      return blank;
    case 'isNotBlank':
      return !blank;
    case 'isError':
      return error;
    case 'isNotError':
      return !error;
    case 'eq':
    case 'neq': {
      if (error || value === null) return false;
      const equal =
        typeof value === typeof criterion.value &&
        (typeof value === 'string' && typeof criterion.value === 'string'
          ? value.toLowerCase() === (foldedQuery ?? criterion.value.toLowerCase())
          : value === criterion.value);
      return criterion.operator === 'eq' ? equal : !equal;
    }
    case 'gt':
      return typeof value === 'number' && value > criterion.value;
    case 'gte':
      return typeof value === 'number' && value >= criterion.value;
    case 'lt':
      return typeof value === 'number' && value < criterion.value;
    case 'lte':
      return typeof value === 'number' && value <= criterion.value;
    default: {
      if (typeof value !== 'string') return false;
      const text = value.toLowerCase(),
        query = foldedQuery ?? criterion.value.toLowerCase();
      switch (criterion.operator) {
        case 'contains':
          return text.includes(query);
        case 'notContains':
          return !text.includes(query);
        case 'startsWith':
          return text.startsWith(query);
        case 'endsWith':
          return text.endsWith(query);
      }
    }
  }
}
