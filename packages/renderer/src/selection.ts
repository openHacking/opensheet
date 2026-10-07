import type { Rect, Selection } from '@opensheetjs/core';
export class GridSelection {
  private mergeSource?: readonly Rect[];
  private mergeRows = new Map<number, Rect[]>();
  constructor(
    public value: Selection,
    private readMerges: (sheetId: string) => readonly Rect[],
  ) {}
  resetMerges() {
    this.mergeSource = undefined;
    this.mergeRows.clear();
  }
  mergesAtRow(row: number, sheetId = this.value.sheetId): Rect[] {
    const merges = this.readMerges(sheetId);
    if (this.mergeSource !== merges) {
      this.mergeSource = merges;
      this.mergeRows.clear();
    }
    let matching = this.mergeRows.get(row);
    if (!matching) {
      matching = merges.filter((m) => row >= m.startRow && row < m.endRow);
      if (this.mergeRows.size >= 256) this.mergeRows.clear();
      this.mergeRows.set(row, matching);
    }
    return matching;
  }
  mergeAt(row: number, column: number, sheetId = this.value.sheetId): Rect | undefined {
    return this.mergesAtRow(row, sheetId).find(
      (m) => column >= m.startColumn && column < m.endColumn,
    );
  }
  expand(selection: Selection): Selection {
    const result = { ...selection };
    if (result.endRow === result.startRow + 1 && result.endColumn === result.startColumn + 1) {
      const merge = this.mergeAt(result.startRow, result.startColumn, result.sheetId);
      return merge ? { ...result, ...merge } : result;
    }
    const merges = this.readMerges(selection.sheetId);
    let changed: boolean;
    do {
      changed = false;
      for (const merge of merges) {
        if (
          merge.endRow <= result.startRow ||
          merge.startRow >= result.endRow ||
          merge.endColumn <= result.startColumn ||
          merge.startColumn >= result.endColumn
        )
          continue;
        const startRow = Math.min(result.startRow, merge.startRow);
        const endRow = Math.max(result.endRow, merge.endRow);
        const startColumn = Math.min(result.startColumn, merge.startColumn);
        const endColumn = Math.max(result.endColumn, merge.endColumn);
        if (
          startRow !== result.startRow ||
          endRow !== result.endRow ||
          startColumn !== result.startColumn ||
          endColumn !== result.endColumn
        ) {
          Object.assign(result, { startRow, endRow, startColumn, endColumn });
          changed = true;
        }
      }
    } while (changed);
    return result;
  }
}
