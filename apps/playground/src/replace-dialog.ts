import { LIMITS, OpenSheetError, type Workbook } from '@opensheetjs/core';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './feedback.js';
import { $ } from './dom.js';

export function showReplaceDialog(
  book: Workbook,
  sheetId: string,
  initialQuery: string,
  { modal, notice, isCurrent }: WorkbookDialogFeedback,
) {
  const baseRevision = book.revision;
  const form = document.createElement('form');
  form.className = 'sort-form';
  const info = document.createElement('p');
  info.textContent = 'Replace case-sensitive literal text in this sheet.';
  const field = (text: string, input: HTMLInputElement) => {
    const label = document.createElement('label');
    label.textContent = text;
    label.append(input);
    return label;
  };
  const query = document.createElement('input');
  query.value = initialQuery;
  query.required = true;
  query.maxLength = LIMITS.text;
  const replacement = document.createElement('input');
  replacement.maxLength = LIMITS.text;
  const status = document.createElement('p');
  status.setAttribute('role', 'alert');
  const save = document.createElement('button');
  save.type = 'submit';
  save.tabIndex = 0;
  save.textContent = 'Replace all';
  const controls = [query, replacement, save];
  const disable = (disabled: boolean) =>
    controls.forEach((control) => {
      control.disabled = disabled;
    });
  if (book.isReadOnly) {
    disable(true);
    status.textContent = 'Read only. Replacement is disabled.';
  }
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    if (!form.reportValidity()) return;
    disable(true);
    status.textContent = 'Replacing…';
    try {
      const commit = await book.execute({
        protocolVersion: 1,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision,
        type: 'core.cells.replace',
        payload: { sheetId, query: query.value, replacement: replacement.value },
      });
      if (currentWorkbookDialog(form, status, isCurrent)) {
        ($('dialog') as HTMLDialogElement).close();
        notice(commit ? 'Text replaced' : 'No text changed');
      }
    } catch (error) {
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      if (error instanceof OpenSheetError && error.code === 'REVISION_CONFLICT') {
        status.textContent = 'Workbook changed. Close this dialog and try again.';
        return;
      }
      status.textContent = error instanceof Error ? error.message : String(error);
      try {
        disable(book.isReadOnly);
      } catch {
        /* A replaced workbook stays disabled. */
      }
    }
  };
  form.append(info, field('Find text', query), field('Replace with', replacement), status, save);
  modal('Replace text', form);
  if (!book.isReadOnly) (query.value ? replacement : query).focus();
}
