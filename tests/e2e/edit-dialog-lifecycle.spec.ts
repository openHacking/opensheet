import { expect, test } from '@playwright/test';

for (const [action, button, failure] of [
  ['paste-special', 'Apply paste', false],
  ['paste-special', 'Apply paste', true],
  ['fill-series', 'Fill series', false],
  ['fill-series', 'Fill series', true],
  ['annotation', 'Save note and link', false],
  ['annotation', 'Save note and link', true],
] as const) {
  test(`${action} ${failure ? 'failure' : 'completion'} preserves newer UI state`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page.evaluate(async (fail) => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[9]]);
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      const execute = book.execute.bind(book);
      let release!: () => void;
      let committed!: () => void;
      const wait = new Promise<void>((resolve) => (release = resolve));
      const written = new Promise<void>((resolve) => (committed = resolve));
      (window as any).editDialogGate = { release, written };
      book.execute = async (...args: Parameters<typeof book.execute>) => {
        const result = await execute(...args);
        committed();
        await wait;
        if (fail) throw new Error('Delayed failure after commit');
        return result;
      };
    }, failure);
    await page.getByLabel('Sheet actions').selectOption(action);
    if (action === 'paste-special') {
      await page.getByLabel('Paste clipboard here').evaluate((element) => {
        const clipboard = new DataTransfer();
        clipboard.setData('text/plain', 'pasted\tsecond');
        const event = new Event('paste', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: clipboard });
        element.dispatchEvent(event);
      });
    } else if (action === 'fill-series') {
      await expect(page.getByRole('button', { name: button })).toBeEnabled();
      await page.getByLabel('Start value').fill('5');
    } else {
      await expect(page.getByRole('button', { name: button })).toBeEnabled();
      await page.getByLabel('Cell note').fill('saved note');
    }
    await page.getByRole('button', { name: button }).click();
    await page.evaluate(() => (window as any).editDialogGate.written);
    const committed = await page.evaluate(async () => {
      const sheet = window.opensheet.getWorkbook().getSheets()[0];
      return {
        values: await sheet.range('A1:B1').getValues(),
        notes: await sheet.range('A1').getNotes(),
      };
    });
    if (action === 'paste-special') expect(committed.values).toEqual([['pasted', 'second']]);
    if (action === 'fill-series') expect(committed.values[0][0]).toBe(5);
    if (action === 'annotation') expect(committed.notes).toEqual([['saved note']]);
    await page.locator('#dialog-close').click();
    await page.evaluate(() => {
      const app = window.opensheet;
      const sheet = app.getWorkbook().getSheets()[0];
      app.selection.set({
        sheetId: sheet.id,
        startRow: 2,
        endRow: 3,
        startColumn: 2,
        endColumn: 3,
      });
    });
    await page.getByLabel('Sheet actions').selectOption('sort-custom');
    await expect(page.getByRole('dialog')).toBeVisible();
    await page.evaluate(async () => {
      const notice = document.querySelector<HTMLElement>('#notice')!;
      notice.hidden = true;
      notice.textContent = '';
      (window as any).editDialogGate.release();
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    await expect(page.getByRole('dialog')).toBeVisible();
    await expect(page.locator('#notice')).toBeHidden();
    expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
      startRow: 2,
      endRow: 3,
      startColumn: 2,
      endColumn: 3,
    });
  });
}

for (const back of [false, true])
  test(`paste preserves a newer selection${back ? ' after returning to the original cell' : ''}`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page.evaluate(() => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      const execute = book.execute.bind(book);
      let release!: () => void;
      let committed!: () => void;
      const wait = new Promise<void>((resolve) => (release = resolve));
      const written = new Promise<void>((resolve) => (committed = resolve));
      (window as any).editDialogGate = { release, written };
      book.execute = async (...args: Parameters<typeof book.execute>) => {
        const result = await execute(...args);
        committed();
        await wait;
        return result;
      };
    });
    await page.getByLabel('Sheet actions').selectOption('paste-special');
    await page.getByLabel('Paste clipboard here').evaluate((element) => {
      const clipboard = new DataTransfer();
      clipboard.setData('text/plain', 'pasted\tsecond');
      const event = new Event('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: clipboard });
      element.dispatchEvent(event);
    });
    await page.getByRole('button', { name: 'Apply paste' }).click();
    await page.evaluate(() => (window as any).editDialogGate.written);
    await page.evaluate((returnToOriginal) => {
      const app = window.opensheet;
      const sheet = app.getWorkbook().getSheets()[0];
      app.selection.set({
        sheetId: sheet.id,
        startRow: 2,
        endRow: 3,
        startColumn: 2,
        endColumn: 3,
      });
      if (returnToOriginal)
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 1,
        });
      (window as any).editDialogGate.release();
    }, back);
    await expect(page.getByRole('dialog')).not.toBeVisible();
    expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
      startRow: back ? 0 : 2,
      endRow: back ? 1 : 3,
      startColumn: back ? 0 : 2,
      endColumn: back ? 1 : 3,
    });
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
      ),
    ).toEqual([['pasted', 'second']]);
  });
