import { columnName, OpenSheetError, type FilterCriterion, type Workbook } from '@opensheetjs/core';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './workbook-dialog.js';

const operators = {
  eq: 'Equals',
  neq: 'Does not equal',
  gt: 'Greater than',
  gte: 'At least',
  lt: 'Less than',
  lte: 'At most',
  contains: 'Contains text',
  notContains: 'Does not contain text',
  startsWith: 'Starts with',
  endsWith: 'Ends with',
  isBlank: 'Is blank',
  isNotBlank: 'Is not blank',
  isError: 'Is error',
  isNotError: 'Is not error',
};
export function showFilterDialog(
  book: Workbook,
  sheetId: string,
  column: number,
  { modal, notice, isCurrent, isTargetCurrent = isCurrent }: WorkbookDialogFeedback,
) {
  if (book.isReadOnly) {
    notice('Switch off read only to filter.');
    return;
  }
  const sheet = book.sheetData(sheetId),
    baseRevision = book.revision;
  const existing = sheet.filter && 'criteria' in sheet.filter ? sheet.filter : null;
  const form = document.createElement('form');
  form.className = 'sort-form';
  const description = document.createElement('p');
  description.textContent =
    'Filter calculated values across the whole sheet. Text matching ignores case; numeric comparisons require numbers. Blank means an empty cell or empty text. Hidden rows are not deleted.';
  const mode = document.createElement('select');
  mode.tabIndex = 0;
  mode.setAttribute('aria-label', 'Match conditions');
  mode.add(new Option('All conditions', 'all'));
  mode.add(new Option('Any condition', 'any'));
  mode.value = existing?.mode ?? 'all';
  const headerLabel = document.createElement('label');
  headerLabel.textContent = 'Header rows (always visible to the filter) ';
  const headers = document.createElement('input');
  headers.type = 'number';
  headers.min = '0';
  headers.max = String(sheet.rowOrder.length);
  headers.step = '1';
  headers.required = true;
  headers.value = String(existing?.headerRows ?? 1);
  headerLabel.append(headers);
  const rows = document.createElement('div');
  const add = document.createElement('button');
  add.tabIndex = 0;
  add.type = 'button';
  add.textContent = 'Add condition';
  const error = document.createElement('p');
  error.setAttribute('role', 'alert');
  const apply = document.createElement('button');
  apply.tabIndex = 0;
  apply.type = 'submit';
  apply.textContent = 'Apply filter';
  const clear = document.createElement('button');
  clear.tabIndex = 0;
  clear.type = 'button';
  clear.textContent = 'Clear all filters';
  const items = (existing?.criteria ?? [{ column, operator: 'isNotBlank' } as FilterCriterion]).map(
    (c) => ({
      column: c.column,
      operator: c.operator,
      type: 'value' in c ? typeof c.value : 'string',
      value: 'value' in c ? String(c.value) : '',
    }),
  );
  function syncRows() {
    [...rows.children].forEach((element, index) => {
      const row = element as HTMLElement;
      const selects = row.querySelectorAll('select');
      selects[0].setAttribute('aria-label', `Condition ${index + 1} column`);
      selects[1].setAttribute('aria-label', `Condition ${index + 1} operator`);
      selects[2].setAttribute('aria-label', `Condition ${index + 1} value type`);
      selects[3].setAttribute('aria-label', `Condition ${index + 1} boolean`);
      row.querySelector('input')!.setAttribute('aria-label', `Condition ${index + 1} value`);
      const remove = row.querySelector('button')!;
      remove.setAttribute('aria-label', `Remove condition ${index + 1}`);
      remove.disabled = items.length === 1;
    });
    add.disabled = items.length >= 64;
  }
  function render(append = false) {
    const start = append ? rows.childElementCount : 0;
    if (!append) rows.replaceChildren();
    items.slice(start).forEach((item, offset) => {
      const index = start + offset;
      const row = document.createElement('div');
      row.className = 'sort-level';
      const columns = document.createElement('select');
      columns.tabIndex = 0;
      columns.setAttribute('aria-label', `Condition ${index + 1} column`);
      for (let c = 0; c < sheet.columnOrder.length; c++)
        columns.add(new Option(`Column ${columnName(c)}`, String(c)));
      columns.value = String(item.column);
      columns.onchange = () => {
        item.column = Number(columns.value);
      };
      const operator = document.createElement('select');
      operator.tabIndex = 0;
      operator.setAttribute('aria-label', `Condition ${index + 1} operator`);
      for (const [value, label] of Object.entries(operators))
        operator.add(new Option(label, value));
      operator.value = item.operator;
      const type = document.createElement('select');
      type.tabIndex = 0;
      type.setAttribute('aria-label', `Condition ${index + 1} value type`);
      for (const [value, label] of [
        ['string', 'Text'],
        ['number', 'Number'],
        ['boolean', 'Boolean'],
      ])
        type.add(new Option(label, value));
      type.value = item.type;
      const value = document.createElement('input');
      value.setAttribute('aria-label', `Condition ${index + 1} value`);
      value.value = item.value;
      const boolean = document.createElement('select');
      boolean.tabIndex = 0;
      boolean.setAttribute('aria-label', `Condition ${index + 1} boolean`);
      boolean.add(new Option('TRUE', 'true'));
      boolean.add(new Option('FALSE', 'false'));
      boolean.value = item.value === 'false' ? 'false' : 'true';
      function update() {
        const unary = item.operator.startsWith('is');
        const equality = item.operator === 'eq' || item.operator === 'neq';
        const numeric = ['gt', 'gte', 'lt', 'lte'].includes(item.operator);
        if (!equality) item.type = numeric ? 'number' : 'string';
        type.value = item.type;
        type.hidden = !equality;
        value.hidden = unary || item.type === 'boolean';
        boolean.hidden = unary || item.type !== 'boolean';
        value.type = item.type === 'number' ? 'number' : 'text';
        value.step = 'any';
        value.value = item.value;
        value.required = !value.hidden && item.type === 'number';
        item.value = item.type === 'boolean' ? boolean.value : value.value;
      }
      operator.onchange = () => {
        item.operator = operator.value as FilterCriterion['operator'];
        update();
      };
      type.onchange = () => {
        item.type = type.value as 'string' | 'number' | 'boolean';
        update();
      };
      value.oninput = () => {
        item.value = value.value;
      };
      boolean.onchange = () => {
        item.value = boolean.value;
      };
      const remove = document.createElement('button');
      remove.tabIndex = 0;
      remove.type = 'button';
      remove.textContent = 'Remove';
      remove.setAttribute('aria-label', `Remove condition ${index + 1}`);
      remove.disabled = items.length === 1;
      remove.onclick = () => {
        const position = items.indexOf(item);
        if (position < 0) return;
        items.splice(position, 1);
        row.remove();
        syncRows();
        add.focus();
      };
      update();
      row.append(columns, operator, type, value, boolean, remove);
      rows.append(row);
    });
    syncRows();
  }
  add.onclick = () => {
    items.push({ column, operator: 'isNotBlank', type: 'string', value: '' });
    render(true);
  };
  async function submit(reset: boolean) {
    if (!currentWorkbookDialog(form, error, isTargetCurrent)) return;
    const controls = [...form.querySelectorAll('button, input, select')] as (
      | HTMLButtonElement
      | HTMLInputElement
      | HTMLSelectElement
    )[];
    const criteria = items.map((item) => ({
      column: item.column,
      operator: item.operator,
      ...(item.operator.startsWith('is')
        ? {}
        : {
            value:
              item.type === 'number'
                ? Number(item.value)
                : item.type === 'boolean'
                  ? item.value === 'true'
                  : item.value,
          }),
    }));
    if (
      !reset &&
      criteria.some((c) => 'value' in c && typeof c.value === 'number' && !Number.isFinite(c.value))
    ) {
      error.textContent = 'Enter a finite number.';
      return;
    }
    controls.forEach((control) => {
      control.disabled = true;
    });
    error.textContent = 'Filtering…';
    try {
      await book.execute({
        protocolVersion: 1,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision,
        type: 'core.sheet.filter',
        payload: {
          sheetId,
          filter: reset ? null : { mode: mode.value, headerRows: Number(headers.value), criteria },
        },
      });
      if (!currentWorkbookDialog(form, error, isCurrent)) return;
      form.closest('dialog')?.close();
      notice(reset ? 'Filters cleared' : 'Filter applied');
    } catch (reason) {
      if (!currentWorkbookDialog(form, error, isCurrent)) return;
      if (reason instanceof OpenSheetError && reason.code === 'REVISION_CONFLICT') {
        error.textContent = 'Workbook changed. Close this dialog and open the filter again.';
        return;
      }
      error.textContent = reason instanceof Error ? reason.message : String(reason);
      controls.forEach((control) => {
        control.disabled = false;
      });
      render();
    }
  }
  form.onsubmit = (event) => {
    event.preventDefault();
    void submit(false);
  };
  clear.onclick = () => {
    void submit(true);
  };
  render();
  form.append(description, mode, headerLabel, rows, add, error, apply, clear);
  modal('Custom filter', form);
}
