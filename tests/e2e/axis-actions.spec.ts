import { workbookTool } from './workbook-tools.js';
import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('B2:C3').setValues([
      [1, 2],
      [3, 4],
    ]);
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 });
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
});

test('inserts selected rows and columns with undo and reopen', async ({ page }) => {
  const revision = () => page.evaluate(() => window.opensheet.getWorkbook().revision);
  let before = await revision();
  await workbookTool(page, 'insert-row');
  await expect.poll(revision).toBe(before + 1);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B4:C5').getValues(),
    ),
  ).toEqual([
    [1, 2],
    [3, 4],
  ]);
  before = await revision();
  await workbookTool(page, 'insert-column');
  await expect.poll(revision).toBe(before + 1);
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const inserted = await sheet.range('D4:E5').getValues();
    await book.undo();
    const undoColumn = await sheet.range('B4:C5').getValues();
    await book.undo();
    const undoRow = await sheet.range('B2:C3').getValues();
    await book.redo();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { inserted, undoColumn, undoRow, reopened: await sheet.range('D4:E5').getValues() };
  });
  const expected = [
    [1, 2],
    [3, 4],
  ];
  expect(result).toEqual({
    inserted: expected,
    undoColumn: expected,
    undoRow: expected,
    reopened: expected,
  });
});

test('hides selected axes and shows all with history, layout and persistence', async ({ page }) => {
  const revision = () => page.evaluate(() => window.opensheet.getWorkbook().revision);
  const action = async (value: string) => {
    const before = await revision();
    await workbookTool(page, value);
    await expect.poll(revision).toBe(before + 1);
  };
  await action('hide-row');
  await action('hide-column');
  const state = () =>
    page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const sh = book.sheetData(book.getSheets()[0].id);
      return {
        rows: sh.rowOrder.slice(0, 4).map((id) => !!sh.rows[id]?.hidden),
        columns: sh.columnOrder.slice(0, 4).map((id) => !!sh.columns[id]?.hidden),
      };
    });
  expect(await state()).toEqual({
    rows: [false, true, true, false],
    columns: [false, true, true, false],
  });
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
  });
  await page.getByRole('grid', { exact: true }).focus();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 3,
    startColumn: 3,
  });
  await action('show-rows');
  await action('show-columns');
  const shownRevision = await revision();
  await workbookTool(page, 'show-rows');
  await workbookTool(page, 'show-columns');
  expect(await revision()).toBe(shownRevision);
  expect(await state()).toEqual({
    rows: [false, false, false, false],
    columns: [false, false, false, false],
  });
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    await book.undo();
    const sh = book.sheetData(book.getSheets()[0].id);
    const hiddenAfterUndo = sh.columnOrder.slice(1, 3).map((id) => !!sh.columns[id]?.hidden);
    await book.redo();
    await app.open(book.id);
    const reopened = app.getWorkbook().sheetData(app.getWorkbook().getSheets()[0].id);
    return {
      hiddenAfterUndo,
      rows: reopened.rowOrder.slice(0, 4).map((id) => !!reopened.rows[id]?.hidden),
      columns: reopened.columnOrder.slice(0, 4).map((id) => !!reopened.columns[id]?.hidden),
    };
  });
  expect(result).toEqual({
    hiddenAfterUndo: [true, true],
    rows: [false, false, false, false],
    columns: [false, false, false, false],
  });
});

test('shows only hidden axes inside the selection with history and reopen', async ({ page }) => {
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setRowHeight(1, 42);
    await sheet.setColumnWidth(1, 135);
    await sheet.setRowsHidden(1, 2, true);
    await sheet.setRowsHidden(5, 1, true);
    await sheet.setColumnsHidden(1, 2, true);
    await sheet.setColumnsHidden(5, 1, true);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 });
  });
  const revision = () => page.evaluate(() => window.opensheet.getWorkbook().revision);
  const action = async (value: string) => {
    const before = await revision();
    await workbookTool(page, value);
    await expect.poll(revision).toBe(before + 1);
  };
  await action('show-selected-rows');
  await action('show-selected-columns');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const read = () => {
      const sheet = book.sheetData(book.getSheets()[0].id);
      return {
        rows: [1, 2, 5].map((i) => sheet.rows[sheet.rowOrder[i]]),
        columns: [1, 2, 5].map((i) => sheet.columns[sheet.columnOrder[i]]),
      };
    };
    const shown = read();
    await book.undo();
    const undone = read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { shown, undone, reopened: read() };
  });
  expect(result.shown).toEqual({
    rows: [{ size: 42, hidden: false }, { hidden: false }, { hidden: true }],
    columns: [{ size: 135, hidden: false }, { hidden: false }, { hidden: true }],
  });
  expect(result.undone.columns).toEqual([
    { size: 135, hidden: true },
    { hidden: true },
    { hidden: true },
  ]);
  expect(result.reopened).toEqual(result.shown);
  const before = await revision();
  await workbookTool(page, 'show-selected-rows');
  await workbookTool(page, 'show-selected-columns');
  expect(await revision()).toBe(before);
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 4,
      endRow: 6,
      startColumn: 4,
      endColumn: 6,
    });
  });
  await page.getByLabel('Read only', { exact: true }).check();
  await workbookTool(page, 'show-selected-rows', true);
  await workbookTool(page, 'show-selected-columns', true);
  expect(await revision()).toBe(before);
  expect(
    await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.sheetData(book.getSheets()[0].id);
      return [sheet.rows[sheet.rowOrder[5]]?.hidden, sheet.columns[sheet.columnOrder[5]]?.hidden];
    }),
  ).toEqual([true, true]);
});

for (const action of ['delete-row', 'hide-column', 'show-selected-rows', 'freeze-selection']) {
  test(`${action} rejects stale selection while a structural write is pending`, async ({
    page,
  }) => {
    const before = await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let id = -1;
      (window as any).axisSent = 0;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute') {
          (window as any).axisSent++;
          if (id === -1) id = message.id;
        }
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === id) (window as any).axisInsertAck = event;
        else handler.call(worker, event);
      };
      (window as any).releaseAxisInsert = () => {
        handler.call(worker, (window as any).axisInsertAck);
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      (window as any).axisInsertWrite = book.getSheets()[0].insertRows(0);
      return book.revision;
    });
    await page.waitForFunction(() => !!(window as any).axisInsertAck);
    await workbookTool(page, action, true);
    expect(await page.evaluate(() => (window as any).axisSent)).toBe(1);
    await page.evaluate(async () => {
      (window as any).releaseAxisInsert();
      await (window as any).axisInsertWrite;
    });
    const result = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      const data = book.sheetData(sheet.id);
      return {
        revision: book.revision,
        rows: sheet.rowCount,
        cells: await sheet.range('B3:C4').getValues(),
        hiddenRows: Object.values(data.rows).filter((row) => row.hidden).length,
        hiddenColumns: Object.values(data.columns).filter((column) => column.hidden).length,
        freeze: data.freeze,
      };
    });
    expect(result).toEqual({
      revision: before + 1,
      rows: 9,
      cells: [
        [1, 2],
        [3, 4],
      ],
      hiddenRows: 0,
      hiddenColumns: 0,
      freeze: { rows: 0, columns: 0 },
    });
    if (action === 'hide-column') {
      await workbookTool(page, action);
      await expect
        .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
        .toBe(before + 2);
      expect(
        await page.evaluate(() => {
          const book = window.opensheet.getWorkbook();
          return Object.values(book.sheetData(book.getSheets()[0].id).columns).filter(
            (column) => column.hidden,
          ).length;
        }),
      ).toBe(2);
    }
  });
}

for (const [action, axis] of [
  ['show-rows', 'row'],
  ['show-columns', 'column'],
] as const) {
  test(`${action} queues behind an unacknowledged hide`, async ({ page }) => {
    await page.evaluate((kind) => {
      const book = window.opensheet.getWorkbook();
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      const ids = new Set<number>();
      const held: MessageEvent[] = [];
      (window as any).axisHeld = held;
      (window as any).releaseAxis = (index: number) => handler.call(worker, held[index]);
      (window as any).restoreAxis = () => {
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute') ids.add(message.id);
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (ids.has(event.data.id)) held.push(event);
        else handler.call(worker, event);
      };
      const sheet = book.getSheets()[0];
      (window as any).axisHide =
        kind === 'row' ? sheet.setRowsHidden(1, 1, true) : sheet.setColumnsHidden(1, 1, true);
    }, axis);
    await page.waitForFunction(() => (window as any).axisHeld.length === 1);
    await workbookTool(page, action, false, false);
    await page.waitForFunction(() => (window as any).axisHeld.length === 2);
    await page.evaluate(async () => {
      (window as any).releaseAxis(0);
      await (window as any).axisHide;
      (window as any).releaseAxis(1);
      (window as any).restoreAxis();
    });
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending))
      .toBe(0);
    const result = await page.evaluate(async (kind) => {
      const app = window.opensheet;
      let book = app.getWorkbook();
      const hidden = () => {
        const sheet = book.sheetData(book.getSheets()[0].id);
        return kind === 'row'
          ? !!sheet.rows[sheet.rowOrder[1]]?.hidden
          : !!sheet.columns[sheet.columnOrder[1]]?.hidden;
      };
      const shown = hidden();
      await book.undo();
      const undone = hidden();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      return { shown, undone, reopened: hidden() };
    }, axis);
    expect(result).toEqual({ shown: false, undone: true, reopened: false });
  });
}

for (const [action, axis] of [
  ['show-rows', 'row'],
  ['show-columns', 'column'],
] as const) {
  test(`${action} includes axes shifted by a pending insertion`, async ({ page }) => {
    await page.evaluate(async (kind) => {
      const sheet = window.opensheet.getWorkbook().getSheets()[0];
      if (kind === 'row') await sheet.setRowsHidden(7, 1, true);
      else await sheet.setColumnsHidden(7, 1, true);
    }, axis);
    await page.evaluate((kind) => {
      const book = window.opensheet.getWorkbook();
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      const ids = new Set<number>();
      const held: MessageEvent[] = [];
      (window as any).insertHeld = held;
      (window as any).releaseInsert = (index: number) => handler.call(worker, held[index]);
      (window as any).restoreInsert = () => {
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute') ids.add(message.id);
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (ids.has(event.data.id)) held.push(event);
        else handler.call(worker, event);
      };
      const sheet = book.getSheets()[0];
      (window as any).insertWrite = kind === 'row' ? sheet.insertRows(0) : sheet.insertColumns(0);
    }, axis);
    await page.waitForFunction(() => (window as any).insertHeld.length === 1);
    await workbookTool(page, action, false, false);
    await page.waitForFunction(() => (window as any).insertHeld.length === 2);
    await page.evaluate(async () => {
      (window as any).releaseInsert(0);
      await (window as any).insertWrite;
      (window as any).releaseInsert(1);
      (window as any).restoreInsert();
    });
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending))
      .toBe(0);
    const result = await page.evaluate(async (kind) => {
      const app = window.opensheet;
      let book = app.getWorkbook();
      const hidden = () => {
        const sheet = book.sheetData(book.getSheets()[0].id);
        return kind === 'row'
          ? !!sheet.rows[sheet.rowOrder[8]]?.hidden
          : !!sheet.columns[sheet.columnOrder[8]]?.hidden;
      };
      const shown = hidden();
      await book.undo();
      const undone = hidden();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      return { shown, undone, reopened: hidden() };
    }, axis);
    expect(result).toEqual({ shown: false, undone: true, reopened: false });
  });
}

test('a queued visibility no-op preserves the previous edit and redo', async ({ page }) => {
  const before = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    const ids = new Set<number>();
    const held: MessageEvent[] = [];
    (window as any).noopHeld = held;
    (window as any).releaseNoop = (index: number) => handler.call(worker, held[index]);
    (window as any).restoreNoop = () => {
      worker.postMessage = post;
      worker.onmessage = handler;
    };
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') ids.add(message.id);
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (ids.has(event.data.id)) held.push(event);
      else handler.call(worker, event);
    };
    (window as any).noopEdit = book
      .getSheets()[0]
      .range('A1')
      .setValues([[99]]);
    return book.revision;
  });
  await page.waitForFunction(() => (window as any).noopHeld.length === 1);
  await workbookTool(page, 'show-rows', false, false);
  await page.waitForFunction(() => (window as any).noopHeld.length === 2);
  await page.evaluate(async () => {
    (window as any).releaseNoop(0);
    await (window as any).noopEdit;
    (window as any).releaseNoop(1);
    (window as any).restoreNoop();
  });
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending))
    .toBe(0);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before + 1);
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await book.undo();
    const undone = await sheet.range('A1').getValues();
    await sheet.setRowsHidden(0, 8, false);
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { undone, reopened: await sheet.range('A1').getValues() };
  });
  expect(result).toEqual({ undone: [[null]], reopened: [[99]] });
});

test('visibility no-op preserves an in-flight viewport read', async ({ page }) => {
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let readId = -1;
      let reads = 0;
      let held!: MessageEvent;
      let received!: () => void;
      const waiting = new Promise<void>((resolve) => (received = resolve));
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'read') {
          reads++;
          readId = message.id;
        }
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === readId && !held) {
          held = event;
          received();
        } else handler.call(worker, event);
      };
      const first = book.prefetch(sheet.id, sheet.range('A1').bounds);
      await waiting;
      await sheet.setRowsHidden(0, 1, false);
      const second = book.prefetch(sheet.id, sheet.range('A1').bounds);
      handler.call(worker, held);
      await Promise.all([first, second]);
      return {
        reads,
        loaded: book.isLoaded(sheet.id, 0, 0),
        value: book.peekValue(sheet.id, 0, 0),
      };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ reads: 1, loaded: true, value: 42 });
});

test('a newer no-op response still advances the local workbook view', async ({ page }) => {
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
    try {
      const sheet = book.getSheets()[0];
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let firstId = -1;
      let held!: MessageEvent;
      let received!: () => void;
      const waiting = new Promise<void>((resolve) => (received = resolve));
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' && firstId === -1) firstId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === firstId) {
          held = event;
          received();
        } else handler.call(worker, event);
      };
      const edit = sheet.range('A1').setValues([[73]]);
      await waiting;
      const before = book.revision;
      await sheet.setRowsHidden(0, 1, false);
      const after = book.revision;
      const value = await sheet.range('A1').getValues();
      handler.call(worker, held);
      await edit;
      return { before, after, value };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ before: 0, after: 1, value: [[73]] });
});

test('a newer response recovers metadata from a delayed earlier acknowledgement', async ({
  page,
}) => {
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
    try {
      const sheet = book.getSheets()[0];
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let firstId = -1;
      let held!: MessageEvent;
      let received!: () => void;
      const waiting = new Promise<void>((resolve) => (received = resolve));
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' && firstId === -1) firstId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === firstId) {
          held = event;
          received();
        } else handler.call(worker, event);
      };
      const first = sheet.setRowsHidden(7, 1, true);
      await waiting;
      await sheet.setColumnsHidden(0, 1, false);
      const beforeRelease = {
        revision: book.revision,
        hidden: !!book.sheetData(sheet.id).rows[book.sheetData(sheet.id).rowOrder[7]]?.hidden,
      };
      handler.call(worker, held);
      await first;
      return beforeRelease;
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ revision: 1, hidden: true });
});

test('an elided response uses its request snapshot after local metadata changes', async ({
  page,
}) => {
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
    try {
      const sheet = book.getSheets()[0];
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      const ids = new Set<number>();
      const held: MessageEvent[] = [];
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' || message.method === 'history') ids.add(message.id);
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (ids.has(event.data.id)) held.push(event);
        else handler.call(worker, event);
      };
      const first = sheet.setRowsHidden(7, 1, true);
      // The public history API rejects concurrent writes; exercise the transport directly.
      const history = (book as any).rpc('history', [false]);
      while (held.length < 2) await new Promise((resolve) => setTimeout(resolve, 0));
      const elided = !held[1].data.result.view.file;
      handler.call(worker, held[0]);
      await first;
      const hiddenAfterFirst = !!book.sheetData(sheet.id).rows[book.sheetData(sheet.id).rowOrder[7]]
        ?.hidden;
      handler.call(worker, held[1]);
      const acknowledged = await history;
      try {
        (book as any).accept(acknowledged);
      } finally {
        // Raw history RPCs bypass the public undo/redo adoption cleanup.
        (book as any).pendingWrites--;
        (book as any).emitPersistence();
      }
      const data = book.sheetData(sheet.id);
      return {
        elided,
        hiddenAfterFirst,
        revision: book.revision,
        hidden: !!data.rows[data.rowOrder[7]]?.hidden,
      };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ elided: true, hiddenAfterFirst: true, revision: 2, hidden: false });
});

test('an older no-op response cannot roll back a newer committed view', async ({ page }) => {
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
    try {
      const sheet = book.getSheets()[0];
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let firstId = -1;
      let held!: MessageEvent;
      let received!: () => void;
      const waiting = new Promise<void>((resolve) => (received = resolve));
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' && firstId === -1) firstId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === firstId) {
          held = event;
          received();
        } else handler.call(worker, event);
      };
      const noOp = sheet.setRowsHidden(0, 1, false);
      await waiting;
      await sheet.range('A1').setValues([[73]]);
      await book.prefetch(sheet.id, sheet.range('A1').bounds);
      handler.call(worker, held);
      await noOp;
      const revision = book.revision;
      const loaded = book.isLoaded(sheet.id, 0, 0);
      const value = book.peekValue(sheet.id, 0, 0);
      await book.undo();
      const undone = await sheet.range('A1').getValues();
      await book.redo();
      return { revision, loaded, value, undone, redone: await sheet.range('A1').getValues() };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    revision: 1,
    loaded: true,
    value: 73,
    undone: [[null]],
    redone: [[73]],
  });
});

test('a delayed commit acknowledgement keeps the newer view and notifies observers', async ({
  page,
}) => {
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
    try {
      const sheet = book.getSheets()[0];
      const events: number[] = [];
      const stop = book.onCommit((commit: { revision: number }) => events.push(commit.revision));
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let firstId = -1;
      let held!: MessageEvent;
      let received!: () => void;
      const waiting = new Promise<void>((resolve) => (received = resolve));
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' && firstId === -1) firstId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === firstId) {
          held = event;
          received();
        } else handler.call(worker, event);
      };
      const first = sheet.range('A1').setValues([[11]]);
      await waiting;
      const second = await sheet.range('B1').setValues([[22]]);
      await book.prefetch(sheet.id, sheet.range('A1:B1').bounds);
      handler.call(worker, held);
      const firstCommit = await first;
      const revision = book.revision;
      const loaded = book.isLoaded(sheet.id, 0, 0);
      const values = await sheet.range('A1:B1').getValues();
      stop();
      await book.undo();
      const undone = await sheet.range('A1:B1').getValues();
      await book.redo();
      return { revision, loaded, values, undone, events, firstCommit, second };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toMatchObject({
    revision: 2,
    loaded: true,
    values: [[11, 22]],
    undone: [[11, null]],
    events: [2, 1],
    firstCommit: { revision: 1 },
    second: { revision: 2 },
  });
});

test('show all rows uses one bounded command without creating metadata for visible rows', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Large', rows: 100_000, columns: 8 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setRowHeight(0, 45);
    await sheet.setRowsHidden(0, 1, true);
    await sheet.setRowsHidden(99_999, 1, true);
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await workbookTool(page, 'show-rows');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const read = () => {
      const sheet = book.sheetData(book.getSheets()[0].id);
      return {
        count: Object.keys(sheet.rows).length,
        first: sheet.rows[sheet.rowOrder[0]],
        last: sheet.rows[sheet.rowOrder[99_999]],
      };
    };
    const shown = read();
    await book.undo();
    const undone = read();
    await book.redo();
    await app.open(book.id);
    const reopenedBook = app.getWorkbook();
    const sheet = reopenedBook.sheetData(reopenedBook.getSheets()[0].id);
    return {
      shown,
      undone,
      reopened: {
        count: Object.keys(sheet.rows).length,
        first: sheet.rows[sheet.rowOrder[0]],
        last: sheet.rows[sheet.rowOrder[99_999]],
      },
    };
  });
  expect(result.shown).toEqual({
    count: 2,
    first: { size: 45, hidden: false },
    last: { hidden: false },
  });
  expect(result.undone).toEqual({
    count: 2,
    first: { size: 45, hidden: true },
    last: { hidden: true },
  });
  expect(result.reopened).toEqual(result.shown);
});

test('resizes selected rows and columns with layout, history and reopen', async ({ page }) => {
  await workbookTool(page, 'resize-row');
  await page.getByLabel('Row height (pixels)').fill('52.5');
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  await workbookTool(page, 'resize-column');
  await page.getByLabel('Column width (pixels)').fill('174');
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await app.getGrid()!.ready();
    const layout = (app.getGrid() as any).layoutState;
    const geometry = {
      height: layout.rowOffsets[3] - layout.rowOffsets[1],
      width: layout.columnOffsets[3] - layout.columnOffsets[1],
    };
    const read = () => {
      const sh = book.sheetData(book.getSheets()[0].id);
      return {
        rows: sh.rowOrder.slice(1, 3).map((id) => sh.rows[id]?.size ?? null),
        columns: sh.columnOrder.slice(1, 3).map((id) => sh.columns[id]?.size ?? null),
      };
    };
    const resized = read();
    await book.undo();
    await book.undo();
    const undone = read();
    await book.redo();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { geometry, resized, undone, reopened: read() };
  });
  expect(result.geometry).toEqual({ height: 105, width: 348 });
  expect(result.resized).toEqual({ rows: [52.5, 52.5], columns: [174, 174] });
  expect(result.undone).toEqual({ rows: [null, null], columns: [null, null] });
  expect(result.reopened).toEqual(result.resized);
});

test('resize dialog rejects stale edits and respects read-only mode', async ({ page }) => {
  await workbookTool(page, 'resize-row');
  await page.getByLabel('Row height (pixels)').fill('50');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[99]]);
  });
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Workbook changed');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.getByRole('dialog').getByRole('button', { name: 'Close' }).click();
  await page.getByLabel('Read only', { exact: true }).check();
  await workbookTool(page, 'resize-column', true);
  await expect(page.getByRole('dialog')).not.toBeVisible();
});

test('resizes 100,000 rows and restores the prior metadata through history', async ({ page }) => {
  test.setTimeout(60_000);
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Large', rows: 100_000, columns: 8 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.setRowsHidden(500, 1, true);
    await sheet.setRowsHeight(0, 100_000, 42);
    const verify = (resized: boolean) => {
      const data = book.sheetData(sheet.id);
      for (let i = 0; i < 100_000; i++) {
        const row = data.rows[data.rowOrder[i]];
        if ((row?.size ?? null) !== (resized ? 42 : null) || !!row?.hidden !== (i === 500))
          throw new Error(`Row ${i} metadata mismatch`);
      }
    };
    verify(true);
    await book.undo();
    verify(false);
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    verify(true);
    return { count: Object.keys(book.sheetData(sheet.id).rows).length };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.count).toBe(100_000);
});
