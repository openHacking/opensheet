import {
  formulaReferenceSpanAt,
  formulaReferenceSpans,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';

const colors = ['#2563eb', '#d97706', '#9333ea', '#be123c'];
export const formulaReferenceColor = (index: number) => colors[index % colors.length];

/** Cycle selected A1 reference endpoints through relative and absolute modes. */
export function cycleFormulaReference(source: string, start: number, end: number) {
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > source.length
  )
    return undefined;
  const span = formulaReferenceSpanAt(source, start, end);
  if (!span) return undefined;
  const text = source.slice(span.start, span.end);
  const tail = text.slice(text.lastIndexOf('!') + 1);
  const tailStart = span.end - tail.length;
  const references = [...tail.matchAll(/\$?[A-Za-z]{1,3}\$?[1-9]\d*/g)].map((match) => ({
    start: tailStart + match.index,
    end: tailStart + match.index + match[0].length,
    text: match[0],
  }));
  const selected = references.filter((ref) =>
    start === end ? start >= ref.start && start <= ref.end : start < ref.end && end > ref.start,
  );
  if (!selected.length) return undefined;
  const first = selected[0].start,
    last = selected[selected.length - 1].end;
  let replacement = '',
    cursor = first;
  for (const ref of selected) {
    replacement += source.slice(cursor, ref.start);
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(ref.text)!;
    const columnAbsolute = !!match[1],
      rowAbsolute = !!match[3];
    // A1 -> $A$1 -> A$1 -> $A1 -> A1.
    const nextColumn = !columnAbsolute;
    const nextRow = columnAbsolute === rowAbsolute;
    replacement += `${nextColumn ? '$' : ''}${match[2]}${nextRow ? '$' : ''}${match[4]}`;
    cursor = ref.end;
  }
  if (source.length - (last - first) + replacement.length > 32767) return undefined;
  return { start: first, end: last, text: replacement };
}

/** Whether an incomplete formula can accept a reference at this UTF-16 offset. */
export function canInsertFormulaReference(source: string, start: number): boolean {
  if (
    !source.startsWith('=') ||
    source.length > 32767 ||
    !Number.isInteger(start) ||
    start < 1 ||
    start > source.length
  )
    return false;
  let quote = '';
  for (let i = 1; i < start; i++) {
    const ch = source[i]!;
    if (quote) {
      if (ch === quote) {
        if (source[i + 1] === quote && i + 1 < start) i++;
        else quote = '';
      }
    } else if (ch === '"' || ch === "'") quote = ch;
  }
  return !quote && /[=,+\-*/^&<>(:]$/.test(source.slice(0, start).trimEnd());
}

/** Resolve bounded draft hints against a workbook without reading cell contents. */
export function resolveFormulaReferences(
  book: Workbook,
  originSheetId: string,
  source: string,
): Selection[] {
  return resolveFormulaReferenceSpans(book, originSheetId, source).map(
    ({ selection }) => selection,
  );
}

export function resolveFormulaReferenceSpans(
  book: Workbook,
  originSheetId: string,
  source: string,
) {
  const hints = formulaReferenceSpans(source);
  if (!hints.length) return [];
  const sheets = new Map(book.getSheets().map((sheet) => [sheet.name.toLowerCase(), sheet]));
  const origin = book.getSheetById(originSheetId);
  const result: { start: number; end: number; selection: Selection }[] = [];
  for (const hint of hints) {
    const sheet = hint.sheet === undefined ? origin : sheets.get(hint.sheet.toLowerCase());
    if (!sheet || hint.endRow > sheet.rowCount || hint.endColumn > sheet.columnCount) continue;
    result.push({
      start: hint.start,
      end: hint.end,
      selection: {
        sheetId: sheet.id,
        startRow: hint.startRow,
        endRow: hint.endRow,
        startColumn: hint.startColumn,
        endColumn: hint.endColumn,
      },
    });
  }
  return result;
}
