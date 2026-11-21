import { expect, test } from '@playwright/test';

for (const failure of ['plugins', 'grid', 'book', 'toolbar'] as const) {
  test(`editor disposal continues after ${failure} cleanup throws`, async ({ page, context }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (failure) => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[421]]);
      let notifications = 0;
      app.on('lifecycle:disposed', () => {
        notifications++;
      });
      const target =
        failure === 'plugins'
          ? app.plugins
          : failure === 'grid'
            ? app.getGrid()
            : failure === 'book'
              ? book
              : (app as any).toolbarView;
      if (!target) throw new Error('Missing cleanup target');
      const dispose = target.dispose.bind(target);
      let injected = false;
      target.dispose = () => {
        const result = dispose();
        if (!injected) {
          injected = true;
          throw new Error(`Injected ${failure} cleanup failure`);
        }
        return result;
      };
      let error = '';
      try {
        app.dispose();
      } catch (reason) {
        error = String(reason);
      }
      let second = '';
      try {
        app.dispose();
      } catch (reason) {
        second = String(reason);
      }
      let workbookStatus = '';
      try {
        await book.getSheets()[0].range('A1').getValues();
      } catch (reason) {
        workbookStatus = (reason as { code?: string }).code ?? '';
      }
      let editorStatus = '';
      try {
        app.getWorkbook();
      } catch (reason) {
        editorStatus = (reason as { code?: string }).code ?? '';
      }
      return {
        id: book.id,
        error,
        second,
        notifications,
        workbookStatus,
        editorStatus,
        grids: document.querySelectorAll('.os-grid').length,
        editors: document.querySelectorAll('.opensheet').length,
      };
    }, failure);
    expect(result).toEqual({
      id: expect.any(String),
      error: `Error: Injected ${failure} cleanup failure`,
      second: '',
      notifications: 1,
      workbookStatus: 'DISPOSED',
      editorStatus: 'DISPOSED',
      grids: 0,
      editors: 0,
    });
    const reopened = await context.newPage();
    await reopened.goto('/#budget');
    await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
    expect(
      await reopened.evaluate(async () => ({
        id: window.opensheet.getWorkbook().id,
        readOnly: window.opensheet.getWorkbook().isReadOnly,
        value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      })),
    ).toEqual({ id: result.id, readOnly: false, value: [[421]] });
    await reopened.close();
  });
}

test('multiple cleanup failures preserve the first error even when logging fails', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      grid = app.getGrid()!;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[731]]);
    const pluginDispose = app.plugins.dispose.bind(app.plugins),
      gridDispose = grid.dispose.bind(grid);
    app.plugins.dispose = () => {
      pluginDispose();
      throw undefined;
    };
    let injected = false;
    grid.dispose = () => {
      gridDispose();
      if (!injected) {
        injected = true;
        throw new Error('Second cleanup failure');
      }
    };
    const originalLog = console.error;
    let logCalls = 0,
      caught = false,
      reason: unknown = 'not thrown';
    console.error = () => {
      logCalls++;
      throw new Error('Logging failed');
    };
    try {
      app.dispose();
    } catch (error) {
      caught = true;
      reason = error;
    } finally {
      console.error = originalLog;
    }
    let editorStatus = '';
    try {
      app.getWorkbook();
    } catch (error) {
      editorStatus = (error as { code?: string }).code ?? '';
    }
    return {
      id: book.id,
      caught,
      reasonWasUndefined: reason === undefined,
      logCalls,
      editorStatus,
      editors: document.querySelectorAll('.opensheet').length,
    };
  });
  expect(result).toEqual({
    id: expect.any(String),
    caught: true,
    reasonWasUndefined: true,
    logCalls: 1,
    editorStatus: 'DISPOSED',
    editors: 0,
  });
  const reopened = await context.newPage();
  await reopened.goto('/#budget');
  await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await reopened.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      readOnly: window.opensheet.getWorkbook().isReadOnly,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    })),
  ).toEqual({ id: result.id, readOnly: false, value: [[731]] });
  await reopened.close();
});
