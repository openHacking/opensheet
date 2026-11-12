import { expect, test } from '@playwright/test';

test('jumps through visible data regions, extends selections and ignores stale reads', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:A8').setValues([[1], [2], [null], [3], [4], [null], [5], [6]]);
    await sheet.setRowHidden(1, true);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    app.getGrid()!.focus();
  });
  const selection = () => page.evaluate(() => window.opensheet.selection.get());
  const row = async () => (await selection())?.startRow;
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(row).toBe(3);
  await page.keyboard.press('Control+Shift+ArrowDown');
  await expect.poll(async () => (await selection())?.endRow).toBe(5);
  await page.keyboard.press('Control+Shift+ArrowDown');
  await expect.poll(async () => (await selection())?.endRow).toBe(7);
  await page.keyboard.press('Control+Shift+ArrowUp');
  await expect.poll(async () => (await selection())?.endRow).toBe(5);
  await page.evaluate(() => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      original = book.navigateData.bind(book);
    let release!: () => void;
    (window as any).releaseDataJump = () => release();
    book.navigateData = async (...args: Parameters<typeof original>) => {
      const result = await original(...args);
      await new Promise<void>((resolve) => {
        release = resolve;
        (window as any).dataJumpReady = true;
      });
      return result;
    };
    (window as any).restoreDataJump = () => {
      book.navigateData = original;
    };
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => page.evaluate(() => (window as any).dataJumpReady)).toBe(true);
  await page.evaluate(() => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 9, endRow: 10, startColumn: 1, endColumn: 2 });
    (window as any).releaseDataJump();
  });
  await page.waitForTimeout(100);
  expect(await selection()).toMatchObject({
    startRow: 9,
    endRow: 10,
    startColumn: 1,
    endColumn: 2,
  });
  await page.evaluate(() => (window as any).restoreDataJump());
});

test('serializes rapid data jumps and extends from the last completed target', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:A7').setValues([[1], [2], [null], [3], [4], [null], [5]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    const original = book.navigateData.bind(book);
    (window as any).jumpCalls = 0;
    book.navigateData = async (...args: Parameters<typeof original>) => {
      const call = ++(window as any).jumpCalls;
      const result = await original(...args);
      if (call === 1)
        await new Promise<void>((resolve) => {
          (window as any).releaseFirstJump = resolve;
        });
      return result;
    };
    app.getGrid()!.focus();
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseFirstJump))
    .toBe('function');
  await page.keyboard.press('Control+ArrowDown');
  await page.keyboard.press('Control+Shift+ArrowDown');
  expect(await page.evaluate(() => (window as any).jumpCalls)).toBe(1);
  await page.evaluate(() => (window as any).releaseFirstJump());
  await expect.poll(() => page.evaluate(() => (window as any).jumpCalls)).toBe(3);
  await expect
    .poll(() => page.evaluate(() => window.opensheet.selection.get()))
    .toMatchObject({ startRow: 3, endRow: 5, startColumn: 0, endColumn: 1 });
});

test('bounds rapid-jump backlog and reports overflow without discarding accepted jumps', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const state = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0],
      grid = app.getGrid()!;
    if (sheet.rowCount < 131) await sheet.insertRows(sheet.rowCount, 131 - sheet.rowCount);
    await sheet
      .range('A1:A131')
      .setValues(Array.from({ length: 131 }, (_, i) => [i % 2 ? null : i]));
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    grid.focus();
    const original = book.navigateData.bind(book),
      originalError = (grid as any).options.onError;
    (window as any).jumpErrors = [];
    (window as any).jumpPositions = [];
    (grid as any).options.onError = (error: any) => {
      (window as any).jumpErrors.push(error.code);
      originalError?.(error);
    };
    app.selection.onChange((selection) => (window as any).jumpPositions.push(selection.startRow));
    let calls = 0;
    book.navigateData = async (...args: Parameters<typeof original>) => {
      const current = ++calls,
        result = await original(...args);
      if (current === 1)
        await new Promise<void>((resolve) => {
          (window as any).releaseFirstJump = resolve;
        });
      return result;
    };
    const events = Array.from({ length: 66 }, () => {
      const event = new KeyboardEvent('keydown', {
        key: 'ArrowDown',
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      grid.element.dispatchEvent(event);
      return event.defaultPrevented;
    });
    return { events, revision: book.revision };
  });
  expect(state.events).toEqual(Array(66).fill(true));
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseFirstJump))
    .toBe('function');
  await expect.poll(() => page.evaluate(() => (window as any).jumpErrors)).toEqual(['BUSY']);
  await page.evaluate(() => (window as any).releaseFirstJump());
  await expect.poll(() => page.evaluate(() => (window as any).jumpPositions.length)).toBe(65);
  const result = await page.evaluate(() => ({
    positions: (window as any).jumpPositions,
    selection: window.opensheet.selection.get(),
    revision: window.opensheet.getWorkbook().revision,
  }));
  expect(result.positions).toEqual(Array.from({ length: 65 }, (_, i) => (i + 1) * 2));
  expect(result.selection).toMatchObject({
    startRow: 130,
    endRow: 131,
    startColumn: 0,
    endColumn: 1,
  });
  expect(result.revision).toBe(state.revision);
});

test('discarding a pending jump also discards queued jumps', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(() => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      original = book.navigateData.bind(book);
    (window as any).jumpCalls = 0;
    book.navigateData = async (...args: Parameters<typeof original>) => {
      (window as any).jumpCalls++;
      const result = await original(...args);
      await new Promise<void>((resolve) => {
        (window as any).releaseFirstJump = resolve;
      });
      return result;
    };
    app.getGrid()!.focus();
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseFirstJump))
    .toBe('function');
  await page.keyboard.press('Control+ArrowDown');
  await page.keyboard.press('Control+ArrowDown');
  await page.evaluate(() => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 9, endRow: 10, startColumn: 1, endColumn: 2 });
    (window as any).releaseFirstJump();
  });
  await page.waitForTimeout(100);
  expect(await page.evaluate(() => (window as any).jumpCalls)).toBe(1);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 9,
    endRow: 10,
    startColumn: 1,
    endColumn: 2,
  });
});

test('reentrant selection changes cancel later queued jumps', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0],
      original = book.navigateData.bind(book);
    await sheet.range('A1:A3').setValues([[1], [2], [null]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    (window as any).jumpCalls = 0;
    book.navigateData = async (...args: Parameters<typeof original>) => {
      (window as any).jumpCalls++;
      const result = await original(...args);
      await new Promise<void>((resolve) => {
        (window as any).releaseFirstJump = resolve;
      });
      return result;
    };
    app.selection.onChange((selection) => {
      if (selection.startRow === 1 && selection.startColumn === 0)
        app.selection.set({
          sheetId: sheet.id,
          startRow: 9,
          endRow: 10,
          startColumn: 1,
          endColumn: 2,
        });
    });
    app.getGrid()!.focus();
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseFirstJump))
    .toBe('function');
  await page.keyboard.press('Control+ArrowDown');
  await page.evaluate(() => (window as any).releaseFirstJump());
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(9);
  await page.waitForTimeout(100);
  expect(await page.evaluate(() => (window as any).jumpCalls)).toBe(1);
});

test('a failed jump drops queued commands and a later retry succeeds', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0],
      original = book.navigateData.bind(book);
    await sheet.range('A1:A3').setValues([[1], [2], [null]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    (window as any).jumpCalls = 0;
    book.navigateData = async (...args: Parameters<typeof original>) => {
      if (++(window as any).jumpCalls === 1) {
        await new Promise<void>((_resolve, reject) => {
          (window as any).failFirstJump = reject;
        });
      }
      return original(...args);
    };
    app.getGrid()!.focus();
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).failFirstJump))
    .toBe('function');
  await page.keyboard.press('Control+ArrowDown');
  await page.evaluate(() => (window as any).failFirstJump(new Error('injected failure')));
  await page.waitForTimeout(100);
  expect(await page.evaluate(() => (window as any).jumpCalls)).toBe(1);
  expect(await page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(0);
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(1);
  expect(await page.evaluate(() => (window as any).jumpCalls)).toBe(2);
});

test('cancels a 100,000-row Worker scan and successfully retries', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Large', rows: 100000, columns: 1 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A100000').setValues([[1]]);
      const revision = book.revision;
      const controller = new AbortController();
      const pending = book.navigateData(sheet.id, 0, 0, 'row', 1, { signal: controller.signal });
      setTimeout(() => controller.abort(), 5);
      const code = await pending.then(
        () => 'success',
        (error: { code?: string }) => error.code,
      );
      const start = performance.now();
      const jump = await book.navigateData(sheet.id, 0, 0, 'row', 1);
      return { code, revision, after: book.revision, jump, ms: performance.now() - start };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('ABORTED');
  expect(result.after).toBe(result.revision);
  expect(result.jump).toMatchObject({ row: 99999, column: 0, revision: result.revision });
});

test('does not apply a delayed jump after focus leaves the grid', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const before = await page.evaluate(() => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      original = book.navigateData.bind(book);
    book.navigateData = async (...args: Parameters<typeof original>) => {
      const result = await original(...args);
      await new Promise<void>((resolve) => {
        (window as any).releaseDelayedJump = resolve;
      });
      return result;
    };
    app.getGrid()!.focus();
    return app.selection.get();
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseDelayedJump))
    .toBe('function');
  await page.locator('#scene-select').focus();
  await page.evaluate(() => (window as any).releaseDelayedJump());
  await page.waitForTimeout(100);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(before);
});

for (const mode of ['programmatic', 'composition'] as const) {
  test(`does not apply a delayed jump after ${mode} editing starts`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const before = await page.evaluate(() => {
      const app = window.opensheet,
        book = app.getWorkbook(),
        original = book.navigateData.bind(book);
      book.navigateData = async (...args: Parameters<typeof original>) => {
        const result = await original(...args);
        await new Promise<void>((resolve) => {
          (window as any).releaseDelayedJump = resolve;
        });
        return result;
      };
      app.getGrid()!.focus();
      return app.selection.get();
    });
    await page.keyboard.press('Control+ArrowDown');
    await expect
      .poll(() => page.evaluate(() => typeof (window as any).releaseDelayedJump))
      .toBe('function');
    await page.evaluate((kind) => {
      const grid = window.opensheet.getGrid()!;
      if (kind === 'programmatic') grid.startEditing('draft');
      else
        grid.element.dispatchEvent(
          new CompositionEvent('compositionstart', { bubbles: true, data: '文' }),
        );
      (window as any).releaseDelayedJump();
    }, mode);
    await expect(page.locator('.os-cell-editor')).toBeVisible();
    await page.waitForTimeout(100);
    expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(before);
    if (mode === 'programmatic') await expect(page.locator('.os-cell-editor')).toHaveValue('draft');
    else expect(await page.evaluate(() => window.opensheet.getGrid()!.isComposing)).toBe(true);
  });
}

test('jumping from and into merged data keeps the whole merge selected', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:A6').clear({ all: true });
    await sheet.range('A1:A2').merge();
    await sheet.range('A1').setValues([[1]]);
    await sheet.range('A3').setValues([[2]]);
    await sheet.range('A5:A6').merge();
    await sheet.range('A5').setValues([[3]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    app.getGrid()!.focus();
  });
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(2);
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(4);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 4,
    endRow: 6,
  });
  await page.keyboard.press('Control+ArrowUp');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(2);
});

test('data jumps use restored filter state after binary import and reopen', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const state = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:B6').setValues([
      [1, 'yes'],
      [2, 'no'],
      [3, 'no'],
      [4, 'yes'],
      [null, 'yes'],
      [5, 'yes'],
    ]);
    await sheet.setFilter({
      mode: 'all',
      headerRows: 0,
      criteria: [{ column: 1, operator: 'eq', value: 'yes' }],
    });
    const parts: BlobPart[] = [];
    for await (const part of book.streamBinary()) parts.push(part as BlobPart);
    await book.importBinary(new Blob(parts));
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    app.getGrid()!.focus();
    return { filtered: book.filteredRows(sheet.id), revision: book.revision };
  });
  expect(state.filtered).toContain(1);
  expect(state.filtered).toContain(2);
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(3);
  await page.keyboard.press('Control+ArrowDown');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(5);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(state.revision);
});
