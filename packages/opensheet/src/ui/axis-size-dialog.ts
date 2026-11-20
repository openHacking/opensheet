import { OpenSheetError, type Selection, type Workbook } from '@opensheetjs/core';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './workbook-dialog.js';

export function showAxisSizeDialog(
  book: Workbook,
  selection: Selection,
  axis: 'row' | 'column',
  { modal, notice, isCurrent, isTargetCurrent = isCurrent }: WorkbookDialogFeedback,
) {
  const rows = axis === 'row';
  const index = rows ? selection.startRow : selection.startColumn;
  const count = rows ? selection.endRow - index : selection.endColumn - index;
  const data = book.sheetData(selection.sheetId);
  const id = (rows ? data.rowOrder : data.columnOrder)[index];
  const initial = (rows ? data.rows : data.columns)[id]?.size ?? (rows ? 30 : 128);
  const baseRevision = book.revision;
  const form = document.createElement('form');
  form.className = 'sort-form';
  const label = document.createElement('label');
  label.textContent = rows ? 'Row height (pixels)' : 'Column width (pixels)';
  const input = document.createElement('input');
  input.type = 'number';
  input.min = '20';
  input.max = '2000';
  input.step = 'any';
  input.required = true;
  input.value = String(initial);
  label.append(input);
  const status = document.createElement('p');
  status.setAttribute('role', 'alert');
  const save = document.createElement('button');
  save.type = 'submit';
  save.textContent = 'Apply';
  const disable = (disabled: boolean) => {
    input.disabled = disabled;
    save.disabled = disabled;
  };
  disable(book.isReadOnly);
  if (book.isReadOnly) status.textContent = 'Read only. Resizing is disabled.';
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, status, isTargetCurrent)) return;
    const size = input.valueAsNumber;
    if (!Number.isFinite(size) || size < 20 || size > 2000) return;
    disable(true);
    status.textContent = 'Applying…';
    try {
      await book.execute({
        protocolVersion: 1,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision,
        type: 'core.axis.resize',
        payload: { sheetId: selection.sheetId, axis, index, count, size },
      });
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      form.closest('dialog')?.close();
      notice(rows ? 'Row height updated' : 'Column width updated');
    } catch (reason) {
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      if (reason instanceof OpenSheetError && reason.code === 'REVISION_CONFLICT') {
        status.textContent = 'Workbook changed. Close this dialog and select the range again.';
        return;
      }
      status.textContent = reason instanceof Error ? reason.message : String(reason);
      try {
        disable(book.isReadOnly);
      } catch {
        /* Replaced workbook stays disabled. */
      }
    }
  };
  form.append(label, status, save);
  modal(rows ? 'Set row height' : 'Set column width', form);
  if (!book.isReadOnly) input.focus();
}
