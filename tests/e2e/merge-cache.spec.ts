import { expect, test } from '@playwright/test';

test('merge preflights covered formulas and preserves undo and reopen', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 2, columns: 3 }] });
    file.sheets[0].blocks.push({
      row: 0,
      column: 0,
      data: {
        numbers: { positions: [0], values: [7] },
        formulas: { positions: [1], values: ['1+1'] },
        cached: { positions: [1], values: [{ type: 'number', value: 2 }] },
        blanks: [2],
        notes: { positions: [2], values: ['Keep note'] },
      },
    });
    await app.load(file);
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const read = async () => ({
      values: await sheet.range('A1:C1').getValues(),
      formulas: await sheet.range('A1:C1').getFormulas(),
      note: await sheet.range('C1').getNotes(),
      merges: book.sheetData(sheet.id).merges,
    });
    const before = await read(),
      revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    const rejected = await sheet
      .range('A1:C1')
      .merge()
      .then(
        () => 'ok',
        (error) => error.code,
      );
    const failed = await read(),
      failedRevision = book.revision,
      failedBytes = (await book.storageStats()).bytes;
    await sheet.range('A1:C1').merge({ discardCoveredValues: true });
    const merged = await read();
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await read();
    await sheet.range('A1:C1').unmerge();
    const unmerged = await read();
    return {
      before,
      revision,
      bytes,
      rejected,
      failed,
      failedRevision,
      failedBytes,
      merged,
      undo,
      reopened,
      unmerged,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.before.values).toEqual([[7, 2, null]]);
  expect(result.rejected).toBe('INVALID_ARGUMENT');
  expect(result.failed).toEqual(result.before);
  expect(result.failedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.merged.merges).toHaveLength(1);
  expect(result.undo).toEqual(result.before);
  expect(result.reopened).toEqual(result.merged);
  expect(result.unmerged.values).toEqual([[7, null, null]]);
  expect(result.unmerged.note).toEqual([['Keep note']]);
});

test('merging rewrites only changed pages and preserves untouched annotated pages', async ({
  page,
}) => {
  await page.route('**/worker-runtime.ts*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
      let enabled=false, writes=0;
      const original=Engine.prototype.storeBlock;
      Engine.prototype.storeBlock=async function(...args){if(enabled)writes++;return original.apply(this,args);};
      addEventListener('message',({data})=>{
        if(data.startMergeCount){writes=0;enabled=true;}
        if(data.stopMergeCount){enabled=false;postMessage({mergePageWrites:writes});}
      });
    `,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 1 }] });
    file.sheets[0].blocks.push(
      {
        row: 0,
        column: 0,
        data: {
          numbers: { positions: [0], values: [7] },
          formulas: { positions: [32], values: ['1+1'] },
          cached: { positions: [32], values: [{ type: 'number', value: 2 }] },
        },
      },
      {
        row: 1,
        column: 0,
        data: { blanks: [0], notes: { positions: [0], values: ['untouched'] } },
      },
      {
        row: 2,
        column: 0,
        data: {
          formulas: { positions: [0], values: ['2+2'] },
          cached: { positions: [0], values: [{ type: 'number', value: 4 }] },
        },
      },
    );
    await app.load(file);
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const read = async () => ({
      first: await sheet.range('A1:A2').getValues(),
      last: await sheet.range('A129').getValues(),
      note: await sheet.range('A65').getNotes(),
      merges: book.sheetData(sheet.id).merges,
    });
    const before = await read(),
      worker = (book as any).worker as Worker;
    worker.postMessage({ method: 'cancelRequest', args: [-1], startMergeCount: true });
    await sheet.range('A1:A130').merge({ discardCoveredValues: true });
    const writes = await new Promise<number>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Worker count timeout')), 10000);
      worker.addEventListener('message', function onMessage({ data }) {
        if (data.mergePageWrites === undefined) return;
        worker.removeEventListener('message', onMessage);
        clearTimeout(timer);
        resolve(data.mergePageWrites);
      });
      worker.postMessage({ method: 'cancelRequest', args: [-1], stopMergeCount: true });
    });
    const merged = await read();
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await read();
    await sheet.range('A1:A130').unmerge();
    const unmerged = await read();
    return { writes, before, merged, undo, reopened, unmerged };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.writes).toBe(2);
  expect(result.before.first).toEqual([[7], [2]]);
  expect(result.before.last).toEqual([[4]]);
  expect(result.undo).toEqual(result.before);
  expect(result.reopened).toEqual(result.merged);
  expect(result.unmerged.first).toEqual([[7], [null]]);
  expect(result.unmerged.last).toEqual([[null]]);
  expect(result.unmerged.note).toEqual([['untouched']]);
});

test('an aborted metadata-only merge leaves pages, history and accounting intact', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
      let armed=false;
      addEventListener('message',({data})=>{if(data.abortMergeHead)armed=true;});
      const original=IDBObjectStore.prototype.put;
      IDBObjectStore.prototype.put=function(...args){
        const request=original.apply(this,args);
        if(armed && this.name==='heads'){armed=false;this.transaction.abort();postMessage({mergeHeadAborted:true});}
        return request;
      };
    `,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet,
      file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 1 }] });
    for (let page = 0; page < 3; page++) {
      const count = Math.min(64, 130 - page * 64),
        positions = Array.from({ length: count }, (_, i) => i * 32);
      file.sheets[0].blocks.push({
        row: page,
        column: 0,
        data: {
          blanks: positions,
          notes: { positions, values: positions.map((_, i) => `note-${page * 64 + i}`) },
        },
      });
    }
    await app.load(file);
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const request = <T>(r: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    const raw = async () => {
      const db = await request(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const keys = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          request(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            request(tx.objectStore(name).getAllKeys(keys)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const check = async (expected: number) => {
      const notes = await sheet.range('A1:A130').getNotes();
      for (let i = 0; i < 130; i++)
        if (notes[i][0] !== `note-${i}`) throw new Error(`Lost note ${i}`);
      return { merges: book.sheetData(sheet.id).merges.length, notes: notes.length, expected };
    };
    const baseline = await raw(),
      revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    let aborted = false;
    const worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.mergeHeadAborted) aborted = true;
    });
    worker.postMessage({ abortMergeHead: true });
    const code = await sheet
      .range('A1:A130')
      .merge()
      .then(
        () => 'ok',
        (error) => error.code ?? 'error',
      );
    const afterRaw = await raw(),
      afterRevision = book.revision,
      afterBytes = (await book.storageStats()).bytes,
      failed = await check(0);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await check(0);
    await sheet.range('A1:A130').merge();
    const retry = await check(1);
    await book.undo();
    const undo = await check(0);
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const redo = await check(1);
    return {
      aborted,
      code,
      baseline,
      afterRaw,
      revision,
      bytes,
      afterRevision,
      afterBytes,
      failed,
      reopened,
      retry,
      undo,
      redo,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.aborted).toBe(true);
  expect(result.code).not.toBe('ok');
  expect(result.afterRaw).toEqual(result.baseline);
  expect(result.afterRevision).toBe(result.revision);
  expect(result.afterBytes).toBe(result.bytes);
  expect(result.failed).toEqual({ merges: 0, notes: 130, expected: 0 });
  expect(result.reopened).toEqual(result.failed);
  expect(result.retry).toEqual({ merges: 1, notes: 130, expected: 1 });
  expect(result.undo).toEqual(result.failed);
  expect(result.redo).toEqual(result.retry);
});

test('metadata-only merge retains pages staged earlier in the same transaction', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 2 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await book.transaction({}, async (tx) => {
      const staged = tx.getSheets()[0];
      await staged.range('A65').setNote('staged');
      await staged.range('A1:A130').merge();
      await staged.range('B1').setValues([[42]]);
    });
    const read = async () => ({
      note: await sheet.range('A65').getNotes(),
      value: await sheet.range('B1').getValues(),
      merges: book.sheetData(sheet.id).merges.length,
    });
    const committed = await read();
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.committed).toEqual({ note: [['staged']], value: [[42]], merges: 1 });
  expect(result.undo).toEqual({ note: [[null]], value: [[null]], merges: 0 });
  expect(result.reopened).toEqual(result.committed);
});

test('a transaction can create a sheet and use its handle for staged writes', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Existing', rows: 2, columns: 2 }] }));
    let book = app.getWorkbook(),
      expired: any;
    await book.transaction({}, async (tx) => {
      const created = await tx.addSheet('New', { rows: 3, columns: 2 });
      expired = created;
      if (!created) throw new Error('addSheet returned no handle');
      if (tx.getSheetByName('New')?.id !== created.id)
        throw new Error('New sheet missing from scope');
      await created.range('A1').setValues([[5]]);
      await created.range('A1:B1').merge();
    });
    const read = async () => ({
      sheets: book.getSheets().map((s) => s.name),
      value: await book.getSheetByName('New')?.range('A1').getValues(),
      merges: book.getSheetByName('New')
        ? book.sheetData(book.getSheetByName('New')!.id).merges.length
        : 0,
    });
    const committed = await read();
    let expiredCode;
    try {
      expired.range('A1');
      expiredCode = 'ok';
    } catch (error) {
      expiredCode = (error as any).code;
    }
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { committed, undo, reopened: await read(), expiredCode };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.committed).toEqual({ sheets: ['Existing', 'New'], value: [[5]], merges: 1 });
  expect(result.undo).toEqual({ sheets: ['Existing'], value: undefined, merges: 0 });
  expect(result.reopened).toEqual(result.committed);
  expect(result.expiredCode).toBe('DISPOSED');
});

test('transaction can insert repeatedly and write into expanded staged dimensions', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Existing', rows: 2, columns: 2 }] }));
    let book = app.getWorkbook(),
      inside: any;
    const outcome = await book
      .transaction({}, async (tx) => {
        const existing = tx.getSheets()[0];
        await existing.insertRows(0);
        await existing.insertRows(0);
        await existing.insertColumns(0);
        await existing.insertColumns(0);
        await existing.range('D4').setValues([[42]]);
        const fresh = await tx.addSheet('New', { rows: 2, columns: 2 });
        await fresh.insertRows(0);
        await fresh.insertRows(0);
        await fresh.insertColumns(0);
        await fresh.insertColumns(0);
        await fresh.range('D4').setValues([[17]]);
        inside = {
          existing: [existing.rowCount, existing.columnCount],
          fresh: [fresh.rowCount, fresh.columnCount],
          parent: [book.getSheets()[0].rowCount, book.getSheets()[0].columnCount],
        };
      })
      .then(
        () => 'ok',
        (error) => error.code ?? String(error),
      );
    if (outcome !== 'ok') return { outcome, inside };
    const read = async () => ({
      sizes: book.getSheets().map((s) => [s.rowCount, s.columnCount]),
      values: await Promise.all(book.getSheets().map((s) => s.range('D4').getValues())),
    });
    const committed = await read();
    await book.undo();
    const undo = {
      names: book.getSheets().map((s) => s.name),
      size: [book.getSheets()[0].rowCount, book.getSheets()[0].columnCount],
    };
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { outcome, inside, committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.outcome).toBe('ok');
  expect(result.inside).toEqual({ existing: [4, 4], fresh: [4, 4], parent: [2, 2] });
  expect(result.committed).toEqual({
    sizes: [
      [4, 4],
      [4, 4],
    ],
    values: [[[42]], [[17]]],
  });
  expect(result.undo).toEqual({ names: ['Existing'], size: [2, 2] });
  expect(result.reopened).toEqual(result.committed);
});

test('staged insert-delete-insert keeps IDs monotonic while reads remain committed', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 2, columns: 1 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:A2').setValues([[1], [2]]);
    let inside: any;
    await book.transaction({}, async (tx) => {
      const staged = tx.getSheets()[0];
      await staged.insertRows(1);
      await staged.range('A2').setValues([[9]]);
      const committedRead = await staged.range('A1:A2').getValues();
      const used = staged.getUsedRange();
      const usedRead = await used.getValues();
      used.bounds.endRow = 1;
      const usedEnd = staged.getUsedRange().bounds.endRow;
      const outside = await staged
        .range('A3')
        .getValues()
        .then(
          () => 'ok',
          (error) => error.code,
        );
      await staged.deleteRows(1);
      await staged.insertRows(2);
      await staged.range('A3').setValues([[7]]);
      inside = {
        committedRead,
        usedRead,
        usedEnd,
        outside,
        count: staged.rowCount,
        committedMetadataRows: tx.sheetData(staged.id).rowOrder.length,
        parent: book.getSheets()[0].rowCount,
      };
    });
    const read = async () => ({
      ids: book.sheetData(sheet.id).rowOrder,
      next: book.sheetData(sheet.id).nextRowId,
      values: await sheet.range('A1:A3').getValues(),
    });
    const committed = await read();
    await book.undo();
    const undo = await sheet.range('A1:A2').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { inside, committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.inside).toEqual({
    committedRead: [[1], [2]],
    usedRead: [[1], [2]],
    usedEnd: 2,
    outside: 'INVALID_RANGE',
    count: 3,
    committedMetadataRows: 2,
    parent: 2,
  });
  expect(result.committed).toEqual({ ids: [0, 1, 3], next: 4, values: [[1], [2], [7]] });
  expect(result.undo).toEqual([[1], [2]]);
  expect(result.reopened).toEqual(result.committed);
});

test('used range reads committed data after staged dimensions shrink', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 3, columns: 2 }] }));
    let book = app.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1:B3')
      .setValues([
        [1, 2],
        [3, 4],
        [5, 6],
      ]);
    let inside: any;
    await book.transaction({}, async (tx) => {
      const sheet = tx.getSheets()[0];
      await sheet.deleteRows(1, 2);
      await sheet.deleteColumns(1);
      inside = {
        dimensions: [sheet.rowCount, sheet.columnCount],
        committedUsed: await sheet.getUsedRange().getValues(),
      };
    });
    const read = async () => ({
      dimensions: [book.getSheets()[0].rowCount, book.getSheets()[0].columnCount],
      used: await book.getSheets()[0].getUsedRange().getValues(),
    });
    const committed = await read();
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { inside, committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.inside).toEqual({
    dimensions: [1, 1],
    committedUsed: [
      [1, 2],
      [3, 4],
      [5, 6],
    ],
  });
  expect(result.committed).toEqual({ dimensions: [1, 1], used: [[1]] });
  expect(result.undo).toEqual({
    dimensions: [3, 2],
    used: [
      [1, 2],
      [3, 4],
      [5, 6],
    ],
  });
  expect(result.reopened).toEqual(result.committed);
});

test('staged sheet handles remain scoped through rollback, rename and removal', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Existing', rows: 2, columns: 2 }] }));
    const book = app.getWorkbook();
    let leaked: any,
      readCode = '',
      usedRangeCode = '',
      calculationRevision = -1,
      insideParent: string[] = [],
      scopeNames: string[] = [];
    const revision = book.revision;
    const rollback = await book
      .transaction({}, async (tx) => {
        const sheet = await tx.addSheet('Draft', { rows: 2, columns: 2 });
        leaked = sheet;
        insideParent = book.getSheets().map((s) => s.name);
        scopeNames = tx.getSheets().map((s) => s.name);
        try {
          sheet.getUsedRange();
          usedRangeCode = 'ok';
        } catch (error) {
          usedRangeCode = (error as any).code;
        }
        calculationRevision = (await tx.calculation.calculate()).revision;
        try {
          await sheet.range('A1').getValues();
          readCode = 'ok';
        } catch (error) {
          readCode = (error as any).code;
        }
        await sheet.range('A1').setValues([[42]]);
        throw new Error('cancel draft');
      })
      .then(
        () => 'ok',
        (error) => error.message,
      );
    let expiredCode = '';
    try {
      leaked.range('A1');
    } catch (error) {
      expiredCode = (error as any).code;
    }
    const afterRollback = { names: book.getSheets().map((s) => s.name), revision: book.revision };
    let afterRename: string[] = [],
      afterRemove: string[] = [];
    await book.transaction({}, async (tx) => {
      const sheet = await tx.addSheet('Temp', { rows: 2, columns: 2 });
      await tx.renameSheet(sheet.id, 'Renamed');
      afterRename = tx.getSheets().map((s) => s.name);
      await tx.removeSheet(sheet.id);
      afterRemove = tx.getSheets().map((s) => s.name);
    });
    return {
      revision,
      rollback,
      insideParent,
      scopeNames,
      readCode,
      usedRangeCode,
      calculationRevision,
      expiredCode,
      afterRollback,
      afterRename,
      afterRemove,
      final: book.getSheets().map((s) => s.name),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.rollback).toBe('cancel draft');
  expect(result.insideParent).toEqual(['Existing']);
  expect(result.scopeNames).toEqual(['Existing', 'Draft']);
  expect(result.readCode).toBe('INVALID_ARGUMENT');
  expect(result.usedRangeCode).toBe('INVALID_ARGUMENT');
  expect(result.calculationRevision).toBe(result.revision);
  expect(result.expiredCode).toBe('DISPOSED');
  expect(result.afterRollback).toEqual({ names: ['Existing'], revision: result.revision });
  expect(result.afterRename).toEqual(['Existing', 'Renamed']);
  expect(result.afterRemove).toEqual(['Existing']);
  expect(result.final).toEqual(['Existing']);
});

test('merge of an unknown sheet returns a stable error without changing data', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[17]]);
    const revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    const code = await book
      .execute({
        type: 'core.cells.merge',
        payload: {
          sheetId: 'missing',
          range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
          discardCoveredValues: false,
        },
      })
      .then(
        () => 'ok',
        (error) => error.code,
      );
    return {
      code,
      revision,
      afterRevision: book.revision,
      bytes,
      afterBytes: (await book.storageStats()).bytes,
      value: await sheet.range('A1').getValues(),
    };
  });
  expect(result.code).toBe('INVALID_ARGUMENT');
  expect(result.afterRevision).toBe(result.revision);
  expect(result.afterBytes).toBe(result.bytes);
  expect(result.value).toEqual([[17]]);
});

test('merge after a staged row insertion uses the staged physical pages', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 1 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A63:A64').setValues([[10], [20]]);
    const request = <T>(r: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    const raw = async () => {
      const db = await request(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const keys = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          request(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            request(tx.objectStore(name).getAllKeys(keys)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const baseline = await raw();
    const revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    const attempt = async (discard: boolean) =>
      book.transaction({}, async (tx) => {
        const staged = tx.getSheets()[0];
        await staged.insertRows(63);
        await staged.range('A64').setValues([[42]]);
        await staged.range('A63:A64').merge({ discardCoveredValues: discard });
      });
    const code = await attempt(false).then(
      () => 'ok',
      (error) => error.code ?? 'error',
    );
    const afterFailure = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      rows: book.sheetData(sheet.id).rowOrder.length,
      values: await sheet.range('A63:A64').getValues(),
    };
    const afterRaw = await raw();
    if (code !== 'INVALID_ARGUMENT')
      return { code, revision, bytes, afterFailure, baseline, afterRaw };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedFailure = {
      rows: book.sheetData(sheet.id).rowOrder.length,
      values: await sheet.range('A63:A64').getValues(),
    };
    await attempt(true);
    const merged = {
      rows: book.sheetData(sheet.id).rowOrder.length,
      merges: book.sheetData(sheet.id).merges.length,
    };
    await sheet.range('A63:A64').unmerge();
    const unmerged = await sheet.range('A63:A65').getValues();
    await book.undo();
    await book.undo();
    const undo = {
      rows: book.sheetData(sheet.id).rowOrder.length,
      values: await sheet.range('A63:A64').getValues(),
    };
    await book.redo();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      code,
      revision,
      bytes,
      afterFailure,
      baseline,
      afterRaw,
      reopenedFailure,
      merged,
      unmerged,
      undo,
      redo: await sheet.range('A63:A65').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('INVALID_ARGUMENT');
  expect(result.afterRaw).toEqual(result.baseline);
  expect(result.afterFailure).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    rows: 130,
    values: [[10], [20]],
  });
  expect(result.reopenedFailure).toEqual({ rows: 130, values: [[10], [20]] });
  expect(result.merged).toEqual({ rows: 131, merges: 1 });
  expect(result.unmerged).toEqual([[10], [null], [20]]);
  expect(result.undo).toEqual({ rows: 130, values: [[10], [20]] });
  expect(result.redo).toEqual(result.unmerged);
});

test('clear after a staged column insertion modifies the staged physical pages', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 2, columns: 70 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('AE1:AF1').setValues([[10, 20]]);
    await book.transaction({}, async (tx) => {
      const staged = tx.getSheets()[0];
      await staged.insertColumns(31);
      await staged.range('AF1').setValues([[42]]);
      await staged.range('AE1:AF1').clear();
    });
    const read = async () => ({
      columns: book.sheetData(sheet.id).columnOrder.length,
      values: await sheet.range('AE1:AG1').getValues(),
    });
    const committed = await read();
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.committed).toEqual({ columns: 71, values: [[null, null, 20]] });
  expect(result.undo).toEqual({ columns: 70, values: [[10, 20, null]] });
  expect(result.reopened).toEqual(result.committed);
});

test('range commands can address a sheet created earlier in the same transaction', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async ({ coreUrl, modelUrl }) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ coreUrl);
      const { createSnapshot } = await import(/* @vite-ignore */ modelUrl);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Existing', rows: 2, columns: 2 }] }));
      let book = app.getWorkbook();
      const fresh = createSnapshot({ sheets: [{ name: 'New', rows: 3, columns: 2 }] }).sheets[0];
      await book.transaction({}, async (tx) => {
        await tx.execute({ type: 'core.sheet.add', payload: { sheet: fresh } });
        await tx.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: fresh.id,
            cells: [{ row: 0, column: 0, input: { type: 'number', value: 5 } }],
          },
        });
        await tx.execute({
          type: 'core.cells.merge',
          payload: {
            sheetId: fresh.id,
            range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
            discardCoveredValues: false,
          },
        });
      });
      const read = async () => ({
        sheets: book.getSheets().map((sheet) => sheet.name),
        value: await book.getSheetByName('New')?.range('A1').getValues(),
        merges: book.getSheetByName('New') ? book.sheetData(fresh.id).merges.length : 0,
      });
      const committed = await read();
      await book.undo();
      const undo = await read();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      return { committed, undo, reopened: await read() };
    },
    {
      coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      modelUrl: `/@fs${process.cwd()}/packages/core/src/model.ts`,
    },
  );
  expect(result.committed).toEqual({ sheets: ['Existing', 'New'], value: [[5]], merges: 1 });
  expect(result.undo).toEqual({ sheets: ['Existing'], value: undefined, merges: 0 });
  expect(result.reopened).toEqual(result.committed);
});
