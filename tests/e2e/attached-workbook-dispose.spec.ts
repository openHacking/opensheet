import { expect, test } from '@playwright/test';

for (const action of ['deleteStorage', 'close', 'dispose'] as const) {
  test(`attached workbook ${action} stops grid work and permits replacement`, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(async (action) => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const events: string[] = [];
      const unsubscribe = book.onDispose(() => events.push('removed'));
      unsubscribe();
      book.onDispose(() => {
        events.push('disposed');
        book.dispose();
      });
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[42]]);
      const file = await book.toJSON();
      await book[action]();
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      );
      const gridRemoved = !document.querySelector('.os-grid');
      await app.load(file);
      const value = (await app.getWorkbook().getSheets()[0].range('A1').getValues())[0][0];
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      );
      return { gridRemoved, value, grids: document.querySelectorAll('.os-grid').length, events };
    }, action);
    expect(result).toEqual({ gridRemoved: true, value: 42, grids: 1, events: ['disposed'] });
    expect(errors).toEqual([]);
  });
}

test('disposal observers are isolated and transaction subscriptions expire', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const events: string[] = [];
    let expired: typeof book | undefined;
    await book.transaction({}, async (scope) => {
      expired = scope;
      scope.onDispose(() => events.push('expired-scope'));
    });
    let expiredCode = '';
    try {
      expired!.onDispose(() => {});
    } catch (error) {
      expiredCode = (error as { code: string }).code;
    }
    const originalError = console.error;
    const errors: string[] = [];
    console.error = (error) => errors.push(String(error));
    try {
      book.onDispose(() => {
        throw new Error('sync observer');
      });
      book.onDispose(async () => {
        throw new Error('async observer');
      });
      book.onDispose(() => events.push('live'));
      book.dispose();
      book.dispose();
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    } finally {
      console.error = originalError;
    }
    return { events, expiredCode, errors };
  });
  expect(result.events).toEqual(['live']);
  expect(result.expiredCode).toBe('DISPOSED');
  expect(result.errors).toEqual(
    expect.arrayContaining([
      expect.stringContaining('sync observer'),
      expect.stringContaining('async observer'),
    ]),
  );
  expect(result.errors).toHaveLength(2);
});

for (const action of ['deleteStorage', 'close', 'dispose'] as const) {
  test(`${action} on an attachment target cannot replace the current workbook`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(async (action) => {
      const app = window.opensheet;
      const original = app.getWorkbook();
      await original
        .getSheets()[0]
        .range('A1')
        .setValues([[42]]);
      const snapshot = await original.toJSON();
      const Workbook = original.constructor as new (input: typeof snapshot) => typeof original;
      const target = new Workbook({ ...snapshot, workbookId: `attach_${crypto.randomUUID()}` });
      await target.ready();
      const ready = target.ready.bind(target);
      let release!: () => void;
      const blocked = new Promise<void>((resolve) => {
        release = resolve;
      });
      target.ready = async () => {
        await blocked;
        return ready();
      };
      const attachment = app.attachWorkbook(target).then(
        () => 'attached',
        (error) => error.code,
      );
      await target[action]();
      release();
      const code = await attachment;
      const current = app.getWorkbook();
      await original
        .getSheets()[0]
        .range('A1')
        .setValues([[43]]);
      return {
        code,
        same: current === original,
        value: await original.getSheets()[0].range('A1').getValues(),
        grids: document.querySelectorAll('.os-grid').length,
      };
    }, action);
    expect(result).toEqual({ code: 'DISPOSED', same: true, value: [[43]], grids: 1 });
  });
}
