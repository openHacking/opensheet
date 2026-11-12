import {
  address,
  isSafeLinkTarget,
  OpenSheetError,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './feedback.js';
import { $ } from './dom.js';

export async function showAnnotationDialog(
  book: Workbook,
  selection: Selection,
  { modal, notice, isCurrent }: WorkbookDialogFeedback,
) {
  const { sheetId, startRow, startColumn } = selection;
  const range = { startRow, endRow: startRow + 1, startColumn, endColumn: startColumn + 1 };
  const baseRevision = book.revision;
  const form = document.createElement('form');
  form.className = 'sort-form annotation-form';
  const info = document.createElement('p');
  info.textContent = `Applies to ${address(startRow, startColumn)}, the first cell of the selection. Cell values are preserved.`;
  const note = document.createElement('textarea');
  note.rows = 6;
  note.maxLength = 32768;
  let originalNote: string | undefined;
  let noteEdited = false;
  note.oninput = () => {
    noteEdited = true;
  };
  const url = document.createElement('input');
  url.type = 'url';
  url.maxLength = 4096;
  url.placeholder = 'https://example.com or mailto:name@example.com';
  const tooltip = document.createElement('input');
  tooltip.maxLength = 1000;
  let originalTooltip: string | undefined;
  let tooltipEdited = false;
  function label(text: string, control: HTMLElement) {
    const label = document.createElement('label');
    label.textContent = text;
    label.append(control);
    return label;
  }
  const link = document.createElement('a');
  link.tabIndex = 0;
  link.textContent = 'Open link';
  link.target = '_blank';
  link.rel = 'noopener noreferrer';
  link.hidden = true;
  const status = document.createElement('p');
  status.setAttribute('role', 'alert');
  status.textContent = 'Loading…';
  const save = document.createElement('button');
  save.tabIndex = 0;
  save.type = 'submit';
  save.textContent = 'Save note and link';
  const remove = document.createElement('button');
  remove.tabIndex = 0;
  remove.type = 'button';
  remove.textContent = 'Remove note and link';
  const controls = [note, url, tooltip, save, remove];
  controls.forEach((control) => {
    control.disabled = true;
  });
  function updateLink() {
    const valid = isSafeLinkTarget(url.value);
    url.setCustomValidity(
      url.value && !valid ? 'Use an absolute HTTP, HTTPS or mailto URL without spaces.' : '',
    );
    link.hidden = !valid;
    if (valid) link.href = url.value;
    else link.removeAttribute('href');
    link.title = tooltip.value;
  }
  url.oninput = updateLink;
  tooltip.oninput = () => {
    tooltipEdited = true;
    updateLink();
  };
  remove.onclick = () => {
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    note.value = '';
    noteEdited = true;
    url.value = '';
    tooltip.value = '';
    tooltipEdited = true;
    updateLink();
    form.requestSubmit();
  };
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    controls.forEach((control) => {
      control.disabled = true;
    });
    status.textContent = 'Saving…';
    try {
      await book.execute({
        protocolVersion: 1,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision,
        type: 'core.cells.annotation',
        payload: {
          sheetId,
          range,
          note: note.value || (!noteEdited && originalNote === '' ? '' : null),
          link: url.value
            ? {
                target: url.value,
                ...(tooltip.value || (!tooltipEdited && originalTooltip === '')
                  ? { tooltip: tooltip.value }
                  : {}),
              }
            : null,
        },
      });
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      ($('dialog') as HTMLDialogElement).close();
      notice('Note and link saved');
    } catch (reason) {
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      if (reason instanceof OpenSheetError && reason.code === 'REVISION_CONFLICT') {
        status.textContent = 'Workbook changed. Close this dialog and select the cell again.';
        return;
      }
      status.textContent = reason instanceof Error ? reason.message : String(reason);
      try {
        if (!book.isReadOnly)
          controls.forEach((control) => {
            control.disabled = false;
          });
      } catch {
        /* A replaced workbook keeps this form disabled. */
      }
    }
  };
  form.append(
    info,
    label('Cell note', note),
    label('Link URL', url),
    label('Link description', tooltip),
    link,
    status,
    save,
    remove,
  );
  modal(`Note and link — ${address(startRow, startColumn)}`, form);
  try {
    const data = await book.readCells(sheetId, range, { fields: ['note', 'link'] });
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    if (data.revision !== baseRevision) {
      status.textContent = 'Workbook changed. Close this dialog and select the cell again.';
      return;
    }
    originalNote = data.cells[0].note;
    note.value = originalNote ?? '';
    url.value = data.cells[0].link?.target ?? '';
    originalTooltip = data.cells[0].link?.tooltip;
    tooltip.value = originalTooltip ?? '';
    updateLink();
    status.textContent = book.isReadOnly
      ? 'Read only. You can view the note and open the link.'
      : '';
    note.readOnly = url.readOnly = tooltip.readOnly = book.isReadOnly;
    note.disabled = url.disabled = tooltip.disabled = false;
    save.disabled = remove.disabled = book.isReadOnly;
    note.focus();
  } catch (reason) {
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    status.textContent = reason instanceof Error ? reason.message : String(reason);
  }
}
