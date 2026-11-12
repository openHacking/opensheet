import { expect, test } from '@playwright/test';

for (const [action, button] of [
  ['paste-special', 'Apply paste'],
  ['fill-series', 'Fill series'],
  ['annotation', 'Save note and link'],
  ['sort-custom', 'Apply sort'],
  ['filter-custom', 'Apply filter'],
  ['replace', 'Replace all'],
] as const) {
  for (const mode of ['reopen', 'reattach'] as const)
    test(`${action} ignores a completion after ${mode}`, async ({ page }) => {
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      await page.evaluate(async () => {
        const app = window.opensheet;
        const book = app.getWorkbook();
        const sheet = book.getSheets()[0];
        await sheet.range('A1:A3').setValues([['replace me'], [3], [1]]);
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 3,
          startColumn: 0,
          endColumn: 1,
        });
        const execute = book.execute.bind(book);
        let release!: () => void;
        let committed!: () => void;
        const wait = new Promise<void>((resolve) => (release = resolve));
        const written = new Promise<void>((resolve) => (committed = resolve));
        (window as any).replacementGate = { release, written };
        book.execute = async (...args: Parameters<typeof book.execute>) => {
          const result = await execute(...args);
          committed();
          await wait;
          return result;
        };
      });
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
      } else if (action === 'annotation') {
        await expect(page.getByRole('button', { name: button })).toBeEnabled();
        await page.getByLabel('Cell note').fill('saved note');
      } else if (action === 'replace') {
        await page.getByLabel('Find text').fill('replace');
        await page.getByLabel('Replace with').fill('changed');
      }
      await page.getByRole('button', { name: button }).click();
      await page.evaluate(() => (window as any).replacementGate.written);
      const original = await page.evaluate(async (mode) => {
        const app = window.opensheet;
        const book = app.getWorkbook();
        const grid = app.getGrid();
        if (mode === 'reopen') await app.open(book.id);
        else await app.attachWorkbook(book);
        const current = app.getWorkbook();
        const sheet = current.getSheets()[0];
        app.selection.set({
          sheetId: sheet.id,
          startRow: 4,
          endRow: 5,
          startColumn: 2,
          endColumn: 3,
        });
        return {
          replaced: current !== book,
          gridReplaced: app.getGrid() !== grid,
          revision: current.revision,
        };
      }, mode);
      expect(original.replaced).toBe(mode === 'reopen');
      expect(original.gridReplaced).toBe(true);
      await page.evaluate(async () => {
        const notice = document.querySelector<HTMLElement>('#notice')!;
        notice.hidden = true;
        notice.textContent = '';
        (window as any).replacementGate.release();
        await new Promise((resolve) => setTimeout(resolve, 0));
      });
      await expect(page.getByRole('dialog')).toBeVisible();
      await expect(page.getByRole('alert')).toHaveText(
        'Workbook changed. Check the previous operation before trying again.',
      );
      await expect(page.locator('#notice')).toBeHidden();
      expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(
        original.revision,
      );
      expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
        startRow: 4,
        endRow: 5,
        startColumn: 2,
        endColumn: 3,
      });
    });
}

for (const action of ['fill-series', 'annotation'] as const) {
  test(`${action} ignores an initial read after workbook replacement`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const read = book.readCells.bind(book);
      let release!: () => void;
      let captured!: () => void;
      const wait = new Promise<void>((resolve) => (release = resolve));
      const readDone = new Promise<void>((resolve) => (captured = resolve));
      (window as any).replacementReadGate = { release, readDone };
      book.readCells = async (...args: Parameters<typeof book.readCells>) => {
        const result = await read(...args);
        captured();
        await wait;
        return result;
      };
    });
    await page.getByLabel('Sheet actions').selectOption(action);
    await page.evaluate(() => (window as any).replacementReadGate.readDone);
    await page.evaluate(async () => {
      const app = window.opensheet;
      await app.open(app.getWorkbook().id);
      (window as any).replacementReadGate.release();
    });
    await expect(page.getByRole('alert')).toHaveText(
      'Workbook changed. Check the previous operation before trying again.',
    );
    await expect(
      page.getByRole('button', {
        name: action === 'annotation' ? 'Save note and link' : 'Fill series',
      }),
    ).toBeDisabled();
  });
}
