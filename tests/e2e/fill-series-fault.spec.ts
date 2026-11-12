import { expect, test } from '@playwright/test';

test('cancelling paged series after staged writes preserves and recovers all cells', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let signalFillWrites = false, fillWrites = 0;
addEventListener('message', ({data}) => { if (data.watchFillWrites) signalFillWrites = true; });
const watchFillPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = watchFillPut.apply(this, args);
  if (signalFillWrites && this.name === 'records') {
    this.transaction.addEventListener('complete', () => {
      fillWrites++;
      if (fillWrites === 2) { signalFillWrites = false; postMessage({cancelFillNow: true}); }
    }, { once: true });
  }
  return request;
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.generate(100_000);
    let sheet = book.getSheets()[0];
    const read = async () => {
      const values = [];
      for (let row = 0; row < 1000; row += 100)
        values.push(
          ...(await sheet
            .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
            .getValues()),
        );
      return values;
    };
    const original = await read();
    const verify = async (filled: boolean) => {
      const values = await read();
      for (let r = 0; r < 1000; r++)
        for (let c = 0; c < 100; c++)
          if (values[r][c] !== (filled ? r + 1 : original[r][c]))
            throw new Error(`Mismatch ${r},${c}`);
    };
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const get = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const rawKeys = async () => {
      const db = await get(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all(
          ['records', 'sizes', 'scratch'].map((name) =>
            get(tx.objectStore(name).getAllKeys(range)),
          ),
        );
      } finally {
        db.close();
      }
    };
    const baselineKeys = await rawKeys();
    const worker = (book as any).worker as Worker;
    let reached = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.cancelFillNow) {
        reached = true;
        book.cancel();
      }
    });
    worker.postMessage({ watchFillWrites: true });
    const code = await sheet
      .range('A1:CV1000')
      .fillSeries({ start: 1, step: 1 })
      .then(
        () => 'success',
        (error) => error.code,
      );
    const failedRevision = book.revision;
    const failedBytes = (await book.storageStats()).bytes;
    const failedKeys = await rawKeys();
    await verify(false);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedRevision = book.revision;
    await verify(false);
    await sheet.range('A1:CV1000').fillSeries({ start: 1, step: 1 });
    await verify(true);
    await book.undo();
    await verify(false);
    await book.redo();
    await verify(true);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify(true);
    return {
      reached,
      code,
      revision,
      failedRevision,
      reopenedRevision,
      bytes,
      failedBytes,
      baselineKeys,
      failedKeys,
    };
  });
  expect(result.reached).toBe(true);
  expect(result.code).toBe('ABORTED');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.failedKeys).toEqual(result.baselineKeys);
});

test('large date series applies values and formatting in one undoable transaction', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    await book.generate(100_000);
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setNote('keep');
    app.selection.set({
      sheetId: sheet.id,
      startRow: 0,
      endRow: 1000,
      startColumn: 0,
      endColumn: 100,
    });
  });
  await page.getByLabel('Sheet actions').selectOption('fill-series');
  await page.getByLabel('Series type').selectOption('day');
  await page.getByLabel('Start date').fill('2024-01-01');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible({ timeout: 60000 });
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const verify = async () => {
      for (let row = 0; row < 1000; row += 100) {
        const values = await sheet
          .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
          .getValues();
        for (let r = 0; r < 100; r++)
          for (let c = 0; c < 100; c++)
            if (values[r][c] !== 45292 + row + r) throw new Error(`Mismatch ${row + r},${c}`);
      }
      const cell = await book.getCell(sheet.id, 0, 0);
      return {
        style: book.getStyle(cell?.styleId),
        note: (await sheet.range('A1').getNotes())[0][0],
      };
    };
    const filled = await verify();
    await book.undo();
    const undone = {
      value: (await sheet.range('A1').getValues())[0][0],
      style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
    };
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await verify();
    return { filled, undone, reopened };
  });
  expect(result.filled.style.numberFormat).toBe('yyyy-mm-dd');
  expect(result.filled.note).toBe('keep');
  expect(result.undone.value).not.toBe(45292);
  expect(result.undone.style.numberFormat).not.toBe('yyyy-mm-dd');
  expect(result.reopened).toEqual(result.filled);
});

test('cancelling paged style after staged writes preserves values, styles and history', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let signalFillWrites = false, fillWrites = 0;
addEventListener('message', ({data}) => { if (data.watchFillWrites) signalFillWrites = true; });
const watchFillPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = watchFillPut.apply(this, args);
  if (signalFillWrites && this.name === 'records') {
    this.transaction.addEventListener('complete', () => {
      fillWrites++;
      if (fillWrites === 2) { signalFillWrites = false; postMessage({cancelFillNow: true}); }
    }, { once: true });
  }
  return request;
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.generate(100_000);
    let sheet = book.getSheets()[0];
    const read = async () => {
      const values = [];
      for (let row = 0; row < 1000; row += 100)
        values.push(
          ...(await sheet
            .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
            .getValues()),
        );
      return values;
    };
    const original = await read();
    const initialStyle = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
    const verify = async (filled: boolean) => {
      const values = await read();
      for (let r = 0; r < 1000; r++)
        for (let c = 0; c < 100; c++)
          if (values[r][c] !== original[r][c]) throw new Error(`Mismatch ${r},${c}`);
      const style = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
      if (
        JSON.stringify(style) !==
        JSON.stringify(filled ? { ...initialStyle, bold: true } : initialStyle)
      )
        throw new Error('Style mismatch');
    };
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const get = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const rawKeys = async () => {
      const db = await get(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all(
          ['records', 'sizes', 'scratch'].map((name) =>
            get(tx.objectStore(name).getAllKeys(range)),
          ),
        );
      } finally {
        db.close();
      }
    };
    const baselineKeys = await rawKeys();
    const worker = (book as any).worker as Worker;
    let reached = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.cancelFillNow) {
        reached = true;
        book.cancel();
      }
    });
    worker.postMessage({ watchFillWrites: true });
    const code = await sheet
      .range('A1:CV1000')
      .setStyle({ bold: true })
      .then(
        () => 'success',
        (error) => error.code,
      );
    const failedRevision = book.revision;
    const failedBytes = (await book.storageStats()).bytes;
    const failedKeys = await rawKeys();
    await verify(false);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedRevision = book.revision;
    await verify(false);
    await sheet.range('A1:CV1000').setStyle({ bold: true });
    await verify(true);
    await book.undo();
    await verify(false);
    await book.redo();
    await verify(true);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify(true);
    return {
      reached,
      code,
      revision,
      failedRevision,
      reopenedRevision,
      bytes,
      failedBytes,
      baselineKeys,
      failedKeys,
    };
  });
  expect(result.reached).toBe(true);
  expect(result.code).toBe('ABORTED');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.failedKeys).toEqual(result.baselineKeys);
});

test('cancelling date transaction during style staging rolls back both commands', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let signalFillWrites = false, fillWrites = 0;
addEventListener('message', ({data}) => { if (data.watchFillWrites) signalFillWrites = true; });
const watchFillPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = watchFillPut.apply(this, args);
  if (signalFillWrites && this.name === 'records') {
    this.transaction.addEventListener('complete', () => {
      fillWrites++;
      if (fillWrites === 66) { signalFillWrites = false; postMessage({cancelFillNow: true}); }
    }, { once: true });
  }
  return request;
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.generate(100_000);
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[123]]);
    await sheet.range('A2').setValues([[456]]);
    await book.undo();
    const priorHistory = { undo: book.canUndo, redo: book.canRedo };
    const metadata = book.metadata.sheets.find((item) => item.id === sheet.id)!;
    const pageCount =
      new Set(metadata.rowOrder.slice(0, 1000).map((id) => Math.floor(id / 64))).size *
      new Set(metadata.columnOrder.slice(0, 100).map((id) => Math.floor(id / 32))).size;
    const priorStyle = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
    const read = async () => {
      const values = [];
      for (let row = 0; row < 1000; row += 100)
        values.push(
          ...(await sheet
            .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
            .getValues()),
        );
      return values;
    };
    const original = await read();
    const verify = async (filled: boolean) => {
      const values = await read();
      for (let r = 0; r < 1000; r++)
        for (let c = 0; c < 100; c++)
          if (values[r][c] !== (filled ? 45292 + r : original[r][c]))
            throw new Error(`Mismatch ${r},${c}`);
      const style = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
      if (
        JSON.stringify(style) !==
        JSON.stringify(filled ? { ...priorStyle, numberFormat: 'yyyy-mm-dd' } : priorStyle)
      )
        throw new Error('Style mismatch');
    };
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const get = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const rawKeys = async () => {
      const db = await get(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          get(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            get(tx.objectStore(name).getAllKeys(range)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const baselineKeys = await rawKeys();
    const worker = (book as any).worker as Worker;
    let reached = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.cancelFillNow) {
        reached = true;
        book.cancel();
      }
    });
    worker.postMessage({ watchFillWrites: true });
    const fill = () =>
      book.transaction({ label: 'date series' }, async (tx) => {
        const range = tx.getSheetById(sheet.id)!.range('A1:CV1000');
        await range.fillSeries({ start: 45292, step: 1 });
        await range.setStyle({ numberFormat: 'yyyy-mm-dd' });
      });
    const code = await fill().then(
      () => 'success',
      (error) => error.code,
    );
    const failedRevision = book.revision;
    const failedBytes = (await book.storageStats()).bytes;
    const failedKeys = await rawKeys();
    const failedHistory = { undo: book.canUndo, redo: book.canRedo };
    await verify(false);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedRevision = book.revision;
    await verify(false);
    await fill();
    await verify(true);
    await book.undo();
    await verify(false);
    await book.redo();
    await verify(true);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify(true);
    return {
      reached,
      pageCount,
      priorHistory,
      failedHistory,
      code,
      revision,
      failedRevision,
      reopenedRevision,
      bytes,
      failedBytes,
      baselineKeys,
      failedKeys,
    };
  });
  expect(result.reached).toBe(true);
  expect(result.pageCount).toBe(64);
  expect(result.failedHistory).toEqual(result.priorHistory);
  expect(result.priorHistory).toEqual({ undo: true, redo: true });
  expect(result.code).toBe('ABORTED');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.failedKeys).toEqual(result.baselineKeys);
});
