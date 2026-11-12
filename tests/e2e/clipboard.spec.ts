import { test, expect } from '@playwright/test';
import { readForPoll } from './read-for-poll.js';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:B1').setValues([[42, 99]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  });
});

test('a final quoted blank clipboard row clears its target and is undoable', async ({ page }) => {
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A2')
      .setValues([['old']]);
  });
  await page.getByRole('grid', { exact: true }).evaluate((element) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', 'new\n""');
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    element.dispatchEvent(event);
  });
  await expect.poll(() => readForPoll(page, 'A1:A2')).toEqual([['new'], [null]]);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book.undo();
      return book.getSheets()[0].range('A1:A2').getValues();
    }),
  ).toEqual([[42], ['old']]);
});

test('plain clipboard fallback keeps a trailing blank copied row', async ({ page }) => {
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:A3').setValues([['first'], [null], [null]]);
    await sheet.range('C1:C3').setValues([['old'], ['restore 2'], ['restore 3']]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: async (text: string) => {
          document.body.dataset.copiedText = text;
        },
      },
    });
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('copy');
  await expect(page.locator('body')).toHaveAttribute('data-copied-text', 'first\n\n""');
  await page.evaluate(() => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 2, endColumn: 3 });
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', document.body.dataset.copiedText!);
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    document.querySelector('[role="grid"]')!.dispatchEvent(event);
  });
  await expect.poll(() => readForPoll(page, 'C1:C3')).toEqual([['first'], [null], [null]]);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book.undo();
      return book.getSheets()[0].range('C1:C3').getValues();
    }),
  ).toEqual([['old'], ['restore 2'], ['restore 3']]);
});

test('clipboard rejection preserves cut source and history', async ({ page }) => {
  const revision = await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: async () => {
          throw new Error('Clipboard permission denied');
        },
      },
    });
    return window.opensheet.getWorkbook().revision;
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('cut');
  await expect(page.locator('#notice')).toContainText('Clipboard permission denied');
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        revision: book.revision,
        values: await book.getSheets()[0].range('A1:B1').getValues(),
      };
    }),
  ).toEqual({ revision, values: [[42, 99]] });
});

test('pending cut keeps the captured source when selection moves and is undoable', async ({
  page,
}) => {
  await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: (text: string) => {
          document.body.dataset.clipboardText = text;
          return new Promise<void>((resolve) =>
            window.addEventListener('release-clipboard', () => resolve(), { once: true }),
          );
        },
      },
    });
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('cut');
  await expect(page.locator('body')).toHaveAttribute('data-clipboard-text', '42');
  expect(
    await page.evaluate(async () => {
      const app = window.opensheet,
        sheet = app.getWorkbook().getSheets()[0];
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 1,
        endColumn: 2,
      });
      return sheet.range('A1:B1').getValues();
    }),
  ).toEqual([[42, 99]]);
  await page.evaluate(() => window.dispatchEvent(new Event('release-clipboard')));
  await expect.poll(() => readForPoll(page, 'A1:B1')).toEqual([[null, 99]]);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book.undo();
      return book.getSheets()[0].range('A1:B1').getValues();
    }),
  ).toEqual([[42, 99]]);
});

test('pending cut cannot erase a newer edit', async ({ page }) => {
  await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: (text: string) => {
          document.body.dataset.clipboardText = text;
          return new Promise<void>((resolve) =>
            window.addEventListener('release-clipboard', () => resolve(), { once: true }),
          );
        },
      },
    });
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('cut');
  await expect(page.locator('body')).toHaveAttribute('data-clipboard-text', '42');
  const revision = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[7]]);
    window.dispatchEvent(new Event('release-clipboard'));
    return book.revision;
  });
  await expect(page.locator('#notice')).toBeVisible();
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        revision: book.revision,
        values: await book.getSheets()[0].range('A1:B1').getValues(),
      };
    }),
  ).toEqual({ revision, values: [[7, 99]] });
});

test('pending cut cannot clear cells after read-only is enabled', async ({ page }) => {
  const revision = await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: (text: string) => {
          document.body.dataset.copiedText = text;
          return new Promise<void>((resolve) =>
            window.addEventListener('release-clipboard', () => resolve(), { once: true }),
          );
        },
      },
    });
    return window.opensheet.getWorkbook().revision;
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('cut');
  await expect(page.locator('body')).toHaveAttribute('data-copied-text', '42');
  await page.evaluate(() => {
    window.opensheet.getWorkbook().setReadOnly(true);
    window.dispatchEvent(new Event('release-clipboard'));
  });
  await expect(page.locator('#notice')).toContainText(/read.only/i);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        revision: book.revision,
        values: await book.getSheets()[0].range('A1:B1').getValues(),
      };
    }),
  ).toEqual({ revision, values: [[42, 99]] });
});

test('pending cut cannot clear cells when only the grid becomes read-only', async ({ page }) => {
  const revision = await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: (text: string) => {
          document.body.dataset.copiedText = text;
          return new Promise<void>((resolve) =>
            window.addEventListener('release-clipboard', () => resolve(), { once: true }),
          );
        },
      },
    });
    return window.opensheet.getWorkbook().revision;
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('cut');
  await expect(page.locator('body')).toHaveAttribute('data-copied-text', '42');
  await page.evaluate(() => {
    window.opensheet.getGrid()!.setReadOnly(true);
    window.dispatchEvent(new Event('release-clipboard'));
  });
  await expect(page.locator('#notice')).toContainText(/read.only/i);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        revision: book.revision,
        readOnly: book.isReadOnly,
        values: await book.getSheets()[0].range('A1:B1').getValues(),
      };
    }),
  ).toEqual({ revision, readOnly: false, values: [[42, 99]] });
});

test('disposing the grid during a pending cut leaves the workbook and host untouched', async ({
  page,
}) => {
  const revision = await page.evaluate(() => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      grid = app.getGrid()!;
    const execute = book.execute.bind(book);
    (book as any).execute = (...args: any[]) => {
      document.body.dataset.lateExecute = 'true';
      return (execute as any)(...args);
    };
    (grid as any).options.onError = () => {
      document.body.dataset.lateError = 'true';
    };
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: (text: string) => {
          document.body.dataset.copiedText = text;
          return new Promise<void>((resolve) =>
            window.addEventListener('release-clipboard', () => resolve(), { once: true }),
          );
        },
      },
    });
    return book.revision;
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('cut');
  await expect(page.locator('body')).toHaveAttribute('data-copied-text', '42');
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook();
    app.getGrid()!.dispose();
    window.dispatchEvent(new Event('release-clipboard'));
    // Flush the clipboard continuation and its promise rejection handlers.
    await new Promise((resolve) => setTimeout(resolve, 0));
    return {
      revision: book.revision,
      values: await book.getSheets()[0].range('A1:B1').getValues(),
      execute: document.body.dataset.lateExecute ?? null,
      error: document.body.dataset.lateError ?? null,
    };
  });
  expect(result).toEqual({ revision, values: [[42, 99]], execute: null, error: null });
});

for (const format of ['text', 'html'] as const) {
  test(`late ${format} paste completion preserves a newer selection`, async ({ page }) => {
    await page.evaluate((format) => {
      const app = window.opensheet,
        book = app.getWorkbook(),
        grid = app.getGrid()!;
      const execute = book.execute.bind(book);
      let release!: () => void;
      const gate = new Promise<void>((resolve) => {
        release = resolve;
      });
      (window as any).releasePaste = release;
      (book as any).execute = async (...args: any[]) => {
        const result = await (execute as any)(...args);
        document.body.dataset.pasteCommitted = 'true';
        await gate;
        return result;
      };
      if (format === 'text') grid.pasteText('first\nsecond');
      else grid.pasteHTML('<table><tr><td>first</td></tr><tr><td>second</td></tr></table>');
    }, format);
    await expect(page.locator('body')).toHaveAttribute('data-paste-committed', 'true');
    const result = await page.evaluate(async () => {
      const app = window.opensheet,
        sheet = app.getWorkbook().getSheets()[0];
      app.selection.set({
        sheetId: sheet.id,
        startRow: 3,
        endRow: 4,
        startColumn: 3,
        endColumn: 4,
      });
      (window as any).releasePaste();
      await new Promise((resolve) => setTimeout(resolve, 0));
      return { selection: app.selection.get(), values: await sheet.range('A1:A2').getValues() };
    });
    expect(result.selection).toMatchObject({
      startRow: 3,
      endRow: 4,
      startColumn: 3,
      endColumn: 4,
    });
    expect(result.values).toEqual([['first'], ['second']]);
  });
}

for (const failure of [false, true]) {
  test(`disposed grid ignores ${failure ? 'failed' : 'successful'} pending paste callbacks`, async ({
    page,
  }) => {
    await page.evaluate((failure) => {
      const app = window.opensheet,
        book = app.getWorkbook(),
        grid = app.getGrid()!;
      const execute = book.execute.bind(book);
      let release!: () => void;
      const gate = new Promise<void>((resolve) => {
        release = resolve;
      });
      (window as any).releasePaste = release;
      (grid as any).options.onSelection = () => {
        document.body.dataset.lateSelection = 'true';
      };
      (grid as any).options.onError = () => {
        document.body.dataset.lateError = 'true';
      };
      (book as any).execute = async (...args: any[]) => {
        const result = await (execute as any)(...args);
        document.body.dataset.pasteCommitted = 'true';
        await gate;
        if (failure) throw new Error('delayed failure');
        return result;
      };
      grid.pasteText('new');
    }, failure);
    await expect(page.locator('body')).toHaveAttribute('data-paste-committed', 'true');
    const result = await page.evaluate(async () => {
      const app = window.opensheet,
        book = app.getWorkbook();
      app.getGrid()!.dispose();
      (window as any).releasePaste();
      await new Promise((resolve) => setTimeout(resolve, 0));
      return {
        selection: document.body.dataset.lateSelection ?? null,
        error: document.body.dataset.lateError ?? null,
        value: await book.getSheets()[0].range('A1').getValues(),
      };
    });
    expect(result).toEqual({ selection: null, error: null, value: [['new']] });
  });
}

test('plain and HTML paste reject a merged anchor without changing data or history', async ({
  page,
}) => {
  const revision = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:B1').clear({ all: true });
    await sheet.range('A1').setValues([['original']]);
    await sheet.range('A1:B1').merge();
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    return book.revision;
  });
  for (const html of [false, true]) {
    await page.evaluate((rich) => {
      document.querySelector('#notice')!.textContent = '';
      const clipboardData = new DataTransfer();
      clipboardData.setData('text/plain', 'replacement');
      if (rich) clipboardData.setData('text/html', '<table><tr><td>replacement</td></tr></table>');
      const event = new Event('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: clipboardData });
      document.querySelector('[role="grid"]')!.dispatchEvent(event);
    }, html);
    await expect(page.locator('#notice')).toContainText('Unmerge before pasting');
    expect(
      await page.evaluate(async () => {
        const book = window.opensheet.getWorkbook(),
          sheet = book.getSheets()[0];
        return {
          revision: book.revision,
          values: await sheet.range('A1:B1').getValues(),
          merges: book.sheetData(sheet.id).merges.length,
        };
      }),
    ).toEqual({ revision, values: [['original', null]], merges: 1 });
  }
});
