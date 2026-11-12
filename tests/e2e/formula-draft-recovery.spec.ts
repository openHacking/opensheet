import { expect, test } from '@playwright/test';

for (const outcome of ['success', 'failure'] as const) {
  test(`a delayed formula-bar ${outcome} preserves a newer draft`, async ({ page }) => {
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
        post = worker.postMessage.bind(worker),
        handler = worker.onmessage!;
      let request = -1;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute') request = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id !== request) return handler.call(worker, event);
        (window as any).releaseFormulaWrite = (fail: boolean) =>
          handler.call(
            worker,
            fail
              ? new MessageEvent('message', {
                  data: {
                    id: request,
                    error: { code: 'WORKER_FAILED', message: 'Injected failure' },
                  },
                })
              : event,
          );
      };
    });
    const formula = page.getByRole('combobox', { name: 'Formula bar' });
    await formula.fill('first draft');
    await formula.press('Enter');
    await page.waitForFunction(() => typeof (window as any).releaseFormulaWrite === 'function');
    await formula.fill('newer draft');
    await page.evaluate((fail) => (window as any).releaseFormulaWrite(fail), outcome === 'failure');
    if (outcome === 'failure') {
      const recovery = page.getByRole('region', { name: 'Unconfirmed formula bar drafts' });
      await expect(recovery).toBeVisible();
      await expect(recovery.getByLabel(/Unconfirmed draft from .*A1/)).toHaveValue('first draft');
      await expect(formula).toHaveValue('newer draft');
      await recovery.getByRole('button', { name: /Dismiss draft from .*A1/ }).click();
      await expect(recovery).toBeHidden();
      await expect(formula).toHaveValue('newer draft');
      expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
    } else {
      await expect(page.locator('#save-state')).toHaveText('Editing…');
      await expect(formula).toBeFocused();
      await expect(formula).toHaveValue('newer draft');
    }
  });
}

test('bounds pending formula-bar submissions without discarding the active draft', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    window.opensheet.getWorkbook().execute = () => new Promise(() => {});
  });
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  await formula.fill('keep this draft');
  for (let i = 0; i < 33; i++) await formula.press('Enter');
  await expect(formula).toHaveValue('keep this draft');
  await expect(page.locator('#notice')).toContainText('Resolve unconfirmed drafts');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
});

test('oversized formula-bar input remains editable rather than entering recovery', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  const input = 'x'.repeat(32770);
  await formula.fill(input);
  await formula.press('Enter');
  await expect(page.locator('#notice')).toContainText('Cell input is too long');
  await expect(formula).toHaveValue(input);
  await expect(page.getByRole('region', { name: 'Unconfirmed formula bar drafts' })).toBeHidden();
});
