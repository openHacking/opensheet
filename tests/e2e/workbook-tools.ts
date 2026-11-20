import { expect, type Page } from '@playwright/test';

export async function toolbarTool(page: Page, name: RegExp | string) {
  const toolbar = page.getByRole('toolbar', { name: 'Spreadsheet tools' });
  const tool = toolbar.getByRole('button', { name, exact: typeof name === 'string' });
  if (!(await tool.isVisible())) await toolbar.getByRole('button', { name: 'More tools' }).click();
  return tool;
}

const axisActions: Record<string, [string, RegExp | string]> = {
  'insert-row': ['Row actions', /^Insert \d+ rows? above$/],
  'insert-column': ['Column actions', /^Insert \d+ columns? left$/],
  'delete-row': ['Row actions', /^Delete \d+ rows?$/],
  'delete-column': ['Column actions', /^Delete \d+ columns?$/],
  'hide-row': ['Row actions', /^Hide \d+ rows?$/],
  'hide-column': ['Column actions', /^Hide \d+ columns?$/],
  'show-rows': ['Row actions', 'Show all rows'],
  'show-columns': ['Column actions', 'Show all columns'],
  'show-selected-rows': ['Row actions', 'Show selected rows'],
  'show-selected-columns': ['Column actions', 'Show selected columns'],
  'resize-row': ['Row actions', 'Set selected row height…'],
  'resize-column': ['Column actions', 'Set selected column width…'],
};

export async function workbookTool(
  page: Page,
  action: string,
  disabled = false,
  waitForCommit = true,
) {
  const menu = (name: RegExp | string) =>
    page.getByRole('menuitem', { name, exact: typeof name === 'string' });
  const select = async (name: RegExp | string) => {
    if (disabled) {
      await expect(menu(name)).toBeDisabled();
      await page.keyboard.press('Escape');
    } else {
      await menu(name).click();
      // The new editor dispatches through its normal asynchronous action runner.
      if (waitForCommit && !['resize-row', 'resize-column', 'fill-series'].includes(action))
        await expect(page.locator('.opensheet')).not.toHaveAttribute('aria-busy', 'true');
    }
  };
  const axis = axisActions[action];
  if (axis) {
    await (await toolbarTool(page, axis[0])).click();
    await select(axis[1]);
    return;
  }
  const freeze: Record<string, string> = {
    'freeze-selection': 'Freeze before selection',
    'freeze-first-row': 'Freeze first row',
    'freeze-first-column': 'Freeze first column',
    unfreeze: 'Unfreeze rows and columns',
  };
  if (freeze[action]) {
    const tool = await toolbarTool(page, 'Freeze');
    if (disabled) await expect(tool).toBeDisabled();
    else {
      await tool.click();
      await select(freeze[action]);
    }
    return;
  }
  switch (action) {
    case 'replace':
      await page.getByRole('button', { name: 'Replace literal text' }).click();
      break;
    case 'filter':
      await page.getByRole('button', { name: 'Filter column by search' }).click();
      break;
    case 'filter-custom':
    case 'clear-filter':
      const filter = await toolbarTool(page, /^Filter( \(active\))?$/);
      if (disabled) await expect(filter).toBeDisabled();
      else {
        await filter.click();
        await select(action === 'filter-custom' ? /^(Create|Edit) filter/ : 'Clear filter');
      }
      break;
    case 'sort-custom':
    case 'sort-asc':
    case 'sort-desc':
      const sort = await toolbarTool(page, 'Sort');
      if (disabled) {
        await expect(sort).toBeDisabled();
        break;
      }
      await sort.click();
      await select(
        action === 'sort-custom'
          ? 'Custom sort…'
          : action === 'sort-asc'
            ? 'Sort ascending'
            : 'Sort descending',
      );
      break;
    case 'annotation':
      await (await toolbarTool(page, 'Note & link')).click();
      break;
    case 'fill-series':
    case 'paste-special':
      await (await toolbarTool(page, 'Cell actions')).click();
      if (action === 'fill-series') await select('Fill series…');
      else if (disabled) await select('Paste special');
      else {
        await menu('Paste special').click();
        await select('Paste from keyboard…');
      }
      break;
    default:
      throw new Error(`Unknown workbook tool ${action}`);
  }
}
