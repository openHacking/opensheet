import { address, OpenSheetError, type Selection, type Workbook } from '@opensheetjs/core';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './workbook-dialog.js';

export async function showSeriesDialog(
  book: Workbook,
  selection: Selection,
  { modal, notice, isCurrent, isTargetCurrent = isCurrent }: WorkbookDialogFeedback,
) {
  const { sheetId, startRow, endRow, startColumn, endColumn } = selection;
  const range = { startRow, endRow, startColumn, endColumn };
  const baseRevision = book.revision;
  const form = document.createElement('form');
  form.className = 'sort-form';
  const info = document.createElement('p');
  info.textContent = `Replace values in ${address(startRow, startColumn)}:${address(endRow - 1, endColumn - 1)} with a numeric series. Formatting, notes and links are preserved.`;
  function field(text: string, control: HTMLElement) {
    control.setAttribute('aria-label', text);
    const label = document.createElement('label');
    label.textContent = text;
    label.append(control);
    return label;
  }
  const kind = document.createElement('select');
  kind.add(new Option('Numbers', 'number'));
  kind.add(new Option('Dates (days)', 'day'));
  kind.add(new Option('Dates (months)', 'month'));
  kind.add(new Option('Dates (years)', 'year'));
  const date = document.createElement('input');
  date.type = 'date';
  date.required = true;
  date.value = book.metadata.dateSystem === '1904' ? '1904-01-01' : '1900-01-01';
  const start = document.createElement('input');
  const step = document.createElement('input');
  for (const input of [start, step]) {
    input.type = 'number';
    input.step = 'any';
    input.required = true;
  }
  start.value = '0';
  step.value = '1';
  const dateField = field('Start date', date);
  const startField = field('Start value', start);
  const updateKind = () => {
    const dates = kind.value !== 'number';
    dateField.hidden = !dates;
    date.disabled = !dates;
    startField.hidden = dates;
    start.disabled = dates;
    step.step = dates ? '1' : 'any';
    info.textContent = dates
      ? `Fill ${address(startRow, startColumn)}:${address(endRow - 1, endColumn - 1)} with dates by ${kind.value === 'day' ? 'serial day' : kind.value + ', clamping shorter months'}. Date formatting is applied; other formatting, notes and links are preserved.${book.metadata.dateSystem === '1900' ? ' The 1900 system includes the compatibility date 1900-02-29.' : ''}`
      : `Replace values in ${address(startRow, startColumn)}:${address(endRow - 1, endColumn - 1)} with a numeric series. Formatting, notes and links are preserved.`;
  };
  kind.onchange = updateKind;
  const axis = document.createElement('select');
  axis.tabIndex = 0;
  axis.add(new Option('Down rows (repeat in each column)', 'row'));
  axis.add(new Option('Across columns (repeat in each row)', 'column'));
  const status = document.createElement('p');
  status.setAttribute('role', 'alert');
  status.textContent = 'Loading…';
  const save = document.createElement('button');
  save.tabIndex = 0;
  save.type = 'submit';
  save.textContent = 'Fill series';
  const controls = [kind, start, date, step, axis, save];
  const disable = (value: boolean) =>
    controls.forEach((control) => {
      control.disabled = value;
    });
  disable(true);
  dateField.hidden = true;
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, status, isTargetCurrent)) return;
    const dates = kind.value !== 'number';
    let initial = start.valueAsNumber;
    if (
      !Number.isFinite(step.valueAsNumber) ||
      (dates && !Number.isSafeInteger(step.valueAsNumber))
    )
      return;
    if (dates) {
      const dateSystem = book.metadata.dateSystem;
      const origin = dateSystem === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
      const days = (date.valueAsNumber - origin) / 86400000;
      initial = days + (dateSystem === '1900' && days >= 60 ? 1 : 0);
      const length = axis.value === 'row' ? endRow - startRow : endColumn - startColumn;
      const last = kind.value === 'day' ? initial + step.valueAsNumber * (length - 1) : initial;
      const maximum = dateSystem === '1904' ? 2957003 : 2958465;
      if (
        !Number.isSafeInteger(initial) ||
        !Number.isSafeInteger(last) ||
        Math.min(initial, last) < (dateSystem === '1900' ? 1 : 0) ||
        Math.max(initial, last) > maximum
      ) {
        status.textContent =
          'Dates must stay within the workbook date system (through 9999-12-31).';
        return;
      }
    }
    if (!Number.isFinite(initial)) return;
    disable(true);
    status.textContent = 'Filling…';
    try {
      const command = {
        type: 'core.cells.fillSeries' as const,
        payload: {
          sheetId,
          range,
          start: initial,
          step: step.valueAsNumber,
          axis: axis.value as 'row' | 'column',
          ...(kind.value === 'month' || kind.value === 'year'
            ? { calendar: kind.value as 'month' | 'year' }
            : {}),
        },
      };
      if (dates) {
        if (book.revision !== baseRevision)
          throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed');
        await book.transaction({ label: 'fill date series' }, async (tx) => {
          await tx.execute(command);
          await tx.execute({
            type: 'core.cells.style',
            payload: { sheetId, range, style: { numberFormat: 'yyyy-mm-dd' } },
          });
        });
      } else
        await book.execute({
          protocolVersion: 1,
          commandId: crypto.randomUUID(),
          workbookId: book.id,
          baseRevision,
          ...command,
        });
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      form.closest('dialog')?.close();
      notice(dates ? 'Date series filled' : 'Numeric series filled');
    } catch (reason) {
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      if (reason instanceof OpenSheetError && reason.code === 'REVISION_CONFLICT') {
        status.textContent = 'Workbook changed. Close this dialog and select the range again.';
        return;
      }
      status.textContent = reason instanceof Error ? reason.message : String(reason);
      try {
        disable(book.isReadOnly);
        if (!book.isReadOnly) updateKind();
      } catch {
        /* Replaced workbook stays disabled. */
      }
    }
  };
  form.append(
    info,
    field('Series type', kind),
    startField,
    dateField,
    field('Step', step),
    field('Direction', axis),
    status,
    save,
  );
  modal('Fill series', form);
  try {
    const data = await book.readCells(
      sheetId,
      { startRow, endRow: startRow + 1, startColumn, endColumn: startColumn + 1 },
      { fields: ['value'] },
    );
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    if (data.revision !== baseRevision) {
      status.textContent = 'Workbook changed. Close this dialog and select the range again.';
      return;
    }
    const value = data.cells[0].value;
    if (typeof value === 'number') start.value = String(value);
    disable(book.isReadOnly);
    if (!book.isReadOnly) updateKind();
    status.textContent = book.isReadOnly ? 'Read only. Filling is disabled.' : '';
    if (!book.isReadOnly) start.focus();
  } catch (reason) {
    if (!currentWorkbookDialog(form, status, isCurrent)) return;
    status.textContent = reason instanceof Error ? reason.message : String(reason);
  }
}
