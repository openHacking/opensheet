import type { Patch } from 'immer';
import { assert } from './types.js';
export type HistoryEntry = { forward: Patch[]; inverse: Patch[]; bytes: number };
const BUDGET = 32 * 1024 * 1024;
export class UndoHistory {
  private past: HistoryEntry[] = [];
  private future: HistoryEntry[] = [];
  private bytes = 0;
  get canUndo() {
    return this.past.length > 0;
  }
  get canRedo() {
    return this.future.length > 0;
  }
  prepare(forward: Patch[], inverse: Patch[]): HistoryEntry {
    const bytes = JSON.stringify([forward, inverse]).length * 2;
    assert(
      bytes <= BUDGET,
      'LIMIT_EXCEEDED',
      'Change exceeds the undo budget; import as a new workbook',
    );
    return { forward, inverse, bytes };
  }
  push(entry: HistoryEntry) {
    this.past.push(entry);
    this.bytes += entry.bytes;
    this.future = [];
    while (this.past.length > 100 || this.bytes > BUDGET) this.bytes -= this.past.shift()!.bytes;
  }
  take(redo: boolean) {
    return (redo ? this.future : this.past).pop();
  }
  accept(entry: HistoryEntry, redo: boolean) {
    (redo ? this.past : this.future).push(entry);
    this.bytes += redo ? entry.bytes : -entry.bytes;
  }
  clear() {
    this.past = [];
    this.future = [];
    this.bytes = 0;
  }
}
