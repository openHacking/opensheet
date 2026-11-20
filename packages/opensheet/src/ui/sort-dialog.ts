import {
  parseRange,
  rangeAddress,
  OpenSheetError,
  columnName,
  type Rect,
  type SortColumn,
  type Workbook,
} from '@opensheetjs/core';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './workbook-dialog.js';

export function showSortDialog(
  book: Workbook,
  selection: Rect & { sheetId: string },
  { modal, notice, isCurrent, isTargetCurrent = isCurrent }: WorkbookDialogFeedback,
  options: { column?: number; direction?: 'asc' | 'desc'; suggestRange?: boolean } = {},
) {
  if (book.isReadOnly) {
    notice('Switch off read only to sort.');
    return;
  }
  let range = { ...selection };
  const baseRevision = book.revision;
  const form = document.createElement('form');
  form.className = 'sort-form';
  const description = document.createElement('p');
  description.textContent = `Sort selected rows ${range.startRow + 1}–${range.endRow}, columns ${columnName(range.startColumn)}–${columnName(range.endColumn - 1)}. Levels are applied from top to bottom. Equal rows keep their original order.`;
  const headerLabel = document.createElement('label');
  const header = document.createElement('input');
  header.type = 'checkbox';
  header.tabIndex = 0;
  header.checked = options.suggestRange ?? false;
  headerLabel.append(header, ' First selected row contains headers');
  const levels = document.createElement('div');
  const add = document.createElement('button');
  add.tabIndex = 0;
  add.type = 'button';
  add.textContent = 'Add sort level';
  const error = document.createElement('p');
  error.setAttribute('role', 'alert');
  const apply = document.createElement('button');
  apply.tabIndex = 0;
  apply.type = 'submit';
  apply.textContent = 'Apply sort';
  const keys: SortColumn[] = [
    { column: options.column ?? range.startColumn, direction: options.direction ?? 'asc' },
  ];
  const rangeInput = document.createElement('input');
  rangeInput.setAttribute('aria-label', 'Sort range');
  rangeInput.value = rangeAddress(range);
  const useRange = document.createElement('button');
  useRange.type = 'button';
  useRange.textContent = 'Update range';
  const changeRange = () => {
    range = {
      sheetId: selection.sheetId,
      ...book.getSheetById(selection.sheetId)!.range(parseRange(rangeInput.value)).bounds,
    };
    description.textContent = `Sort ${rangeAddress(range)}. Levels are applied from top to bottom.`;
  };
  useRange.onclick = () => {
    try {
      changeRange();
      const surviving = keys.filter(
        ({ column }) => column >= range.startColumn && column < range.endColumn,
      );
      keys.splice(
        0,
        keys.length,
        ...(surviving.length
          ? surviving
          : [{ column: range.startColumn, direction: options.direction ?? 'asc' }]),
      );
      error.textContent = '';
      render();
    } catch (reason) {
      error.textContent = reason instanceof Error ? reason.message : String(reason);
    }
  };
  function syncRows() {
    [...levels.children].forEach((element, index) => {
      const row = element as HTMLElement;
      row.querySelector('span')!.textContent = `Level ${index + 1}`;
      const selects = row.querySelectorAll('select');
      selects[0].setAttribute('aria-label', `Level ${index + 1} column`);
      selects[1].setAttribute('aria-label', `Level ${index + 1} order`);
      const remove = row.querySelector('button')!;
      remove.setAttribute('aria-label', `Remove level ${index + 1}`);
      remove.disabled = keys.length === 1;
    });
    add.disabled = keys.length >= Math.min(64, range.endColumn - range.startColumn);
  }
  function render(append = false) {
    const start = append ? levels.childElementCount : 0;
    if (!append) levels.replaceChildren();
    keys.slice(start).forEach((key, offset) => {
      const index = start + offset;
      const row = document.createElement('div');
      row.className = 'sort-level';
      const label = document.createElement('span');
      label.textContent = `Level ${index + 1}`;
      const column = document.createElement('select');
      column.tabIndex = 0;
      column.setAttribute('aria-label', `Level ${index + 1} column`);
      for (let c = range.startColumn; c < range.endColumn; c++)
        column.add(new Option(`Column ${columnName(c)}`, String(c)));
      column.value = String(key.column);
      column.onchange = () => {
        key.column = Number(column.value);
      };
      const direction = document.createElement('select');
      direction.tabIndex = 0;
      direction.setAttribute('aria-label', `Level ${index + 1} order`);
      direction.add(new Option('Ascending', 'asc'));
      direction.add(new Option('Descending', 'desc'));
      direction.value = key.direction;
      direction.onchange = () => {
        key.direction = direction.value as SortColumn['direction'];
      };
      const remove = document.createElement('button');
      remove.tabIndex = 0;
      remove.type = 'button';
      remove.textContent = 'Remove';
      remove.setAttribute('aria-label', `Remove level ${index + 1}`);
      remove.disabled = keys.length === 1;
      remove.onclick = () => {
        const position = keys.indexOf(key);
        if (position < 0) return;
        keys.splice(position, 1);
        row.remove();
        syncRows();
        add.focus();
      };
      row.append(label, column, direction, remove);
      levels.append(row);
    });
    syncRows();
  }
  add.onclick = () => {
    let column = range.startColumn;
    while (keys.some((key) => key.column === column)) column++;
    keys.push({ column, direction: 'asc' });
    render(true);
    levels.querySelectorAll('select')[2 * (keys.length - 1)].focus();
  };
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, error, isTargetCurrent)) return;
    if (options.suggestRange) {
      try {
        changeRange();
        if (keys.some(({ column }) => column < range.startColumn || column >= range.endColumn))
          throw new Error(
            'Choose a sort range that contains each sort column, or update the range.',
          );
      } catch (reason) {
        error.textContent = reason instanceof Error ? reason.message : String(reason);
        return;
      }
    }
    if (new Set(keys.map((key) => key.column)).size !== keys.length) {
      error.textContent = 'Choose a different column for each level.';
      return;
    }
    const startRow = range.startRow + Number(header.checked);
    if (startRow >= range.endRow) {
      error.textContent = 'Select at least one data row below the header.';
      return;
    }
    const controls = [...form.querySelectorAll('button, input, select')] as (
      | HTMLButtonElement
      | HTMLInputElement
      | HTMLSelectElement
    )[];
    controls.forEach((control) => {
      control.disabled = true;
    });
    error.textContent = 'Sorting…';
    try {
      await book.execute({
        protocolVersion: 1,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision,
        type: 'core.sheet.sort',
        payload: {
          sheetId: range.sheetId,
          range: {
            startRow,
            endRow: range.endRow,
            startColumn: range.startColumn,
            endColumn: range.endColumn,
          },
          keys,
        },
      });
      if (!currentWorkbookDialog(form, error, isCurrent)) return;
      form.closest('dialog')?.close();
      notice('Selection sorted');
    } catch (reason) {
      if (!currentWorkbookDialog(form, error, isCurrent)) return;
      if (reason instanceof OpenSheetError && reason.code === 'REVISION_CONFLICT') {
        error.textContent = 'Workbook changed. Close this dialog and select the rows again.';
        return;
      }
      error.textContent = reason instanceof Error ? reason.message : String(reason);
      controls.forEach((control) => {
        control.disabled = false;
      });
      render();
    }
  };
  render();
  form.append(description);
  if (options.suggestRange) form.append(rangeInput, useRange);
  form.append(headerLabel, levels, add, error, apply);
  modal('Custom sort', form);
}
