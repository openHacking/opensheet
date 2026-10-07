import type { Selection, Workbook } from '@opensheetjs/core';
export class StatusBar {
  readonly element = document.createElement('div');
  constructor() {
    this.element.className = 'os-status';
    this.element.textContent = 'Ready';
  }
  async update(book: Workbook, s: Selection | null) {
    if (!s) return;
    const count = (s.endRow - s.startRow) * (s.endColumn - s.startColumn);
    let sum = 0,
      n = 0;
    if (count <= 10000) {
      for (const row of await book.getSheetById(s.sheetId)!.range(s).getValues())
        for (const v of row)
          if (typeof v === 'number') {
            sum += v;
            n++;
          }
    }
    this.element.textContent = `${count > 1 ? `${count} cells${n ? ` · Sum ${sum.toLocaleString()}` : ''}` : 'Ready'}  ·  Local only`;
  }
}
