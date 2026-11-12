import { expect, test } from '@playwright/test';

test('a late failed cell write retains its draft without replacing a newer draft', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    window.opensheet.selection.set({
      sheetId: sheet.id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
    const worker = (book as any).worker as Worker,
      originalPost = worker.postMessage.bind(worker),
      originalHandler = worker.onmessage!;
    let requestId = -1;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') requestId = message.id;
      return Reflect.apply(originalPost, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (event.data.id !== requestId) return originalHandler.call(worker, event);
      (window as any).releaseFailedDraft = () =>
        originalHandler.call(
          worker,
          new MessageEvent('message', {
            data: {
              id: requestId,
              error: { code: 'WORKER_FAILED', message: 'Injected lost acknowledgement' },
            },
          }),
        );
    };
  });
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  await page.keyboard.type('first draft');
  await page.getByRole('textbox', { name: 'Edit cell' }).press('Enter');
  await page.waitForFunction(() => typeof (window as any).releaseFailedDraft === 'function');
  await page.keyboard.type('newer draft');
  const active = page.getByRole('textbox', { name: 'Edit cell' });
  await expect(active).toHaveValue('newer draft');
  await page.evaluate(() => (window as any).releaseFailedDraft());
  const recovery = page.getByRole('region', { name: 'Unconfirmed cell drafts' });
  await expect(recovery).toBeVisible();
  await expect(recovery.getByLabel(/Unconfirmed draft from .*A1/)).toHaveValue('first draft');
  await expect(active).toHaveValue('newer draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  await recovery.getByRole('button', { name: /Dismiss draft from .*A1/ }).click();
  await expect(recovery).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  await active.press('Escape');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(false);
});

test('bounds pending cell drafts and retains the next draft when the limit is reached', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    window.opensheet.selection.set({
      sheetId: sheet.id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
    book.execute = () => new Promise(() => {});
  });
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  for (let index = 0; index < 32; index++) {
    await page.keyboard.type('x');
    await page.getByRole('textbox', { name: 'Edit cell' }).press('Enter');
  }
  await page.keyboard.type('keep me');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await editor.press('Enter');
  await expect(editor).toBeVisible();
  await expect(editor).toHaveValue('keep me');
  await expect(page.locator('#notice')).toContainText('Resolve unsaved drafts');
});

test('oversized in-cell input remains editable rather than entering recovery', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  await page.keyboard.type('x');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  const input = 'x'.repeat(32770);
  await editor.fill(input);
  await editor.press('Enter');
  await expect(page.locator('#notice')).toContainText('Cell input is too long');
  await expect(editor).toHaveValue(input);
  await expect(page.getByRole('region', { name: 'Unconfirmed cell drafts' })).toBeHidden();
});
