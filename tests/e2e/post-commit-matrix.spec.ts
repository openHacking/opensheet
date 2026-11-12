import { expect, test } from '@playwright/test';

for (const operation of ['import', 'sort'] as const) {
  test(`recovers ${operation} on reopen when scratch cleanup fails`, async ({ page }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let failImportCleanup = false;
addEventListener('message', event => { if (event.data.armImportCleanup) failImportCleanup = true; });
const deleteScratch = IDBObjectStore.prototype.delete;
IDBObjectStore.prototype.delete = function(key) {
  if (failImportCleanup && this.name === 'scratch' && key instanceof IDBKeyRange) {
    failImportCleanup = false;
    throw new DOMException('Injected scratch cleanup failure', 'UnknownError');
  }
  return deleteScratch.call(this, key);
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (operation) => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setValues([[123], [1]]);
      const revision = book.revision;
      const good = await book.toJSON();
      const bad = { ...good, unsupported: true };
      (book as any).worker.postMessage({ armImportCleanup: true });
      const code = (error: unknown) => (error as { code?: string }).code;
      const first = await (
        operation === 'import' ? book.importJSON(bad) : sheet.sort('A1:A2', 0)
      ).then(() => 'success', code);
      const retry = await (
        operation === 'import' ? book.importJSON(good) : sheet.sort('A1:A2', 0)
      ).then(() => 'success', code);
      await window.opensheet.open(book.id);
      const reopened = window.opensheet.getWorkbook();
      const recoveredRevision = reopened.revision;
      const value = (await reopened.getSheets()[0].range('A1').getValues())[0][0];
      if (operation === 'import') await reopened.importJSON(good);
      else await reopened.getSheets()[0].sort('A1:A2', 0);
      const updatedValue = (await reopened.getSheets()[0].range('A1').getValues())[0][0];
      return {
        first,
        retry,
        revision,
        recoveredRevision,
        value,
        updatedValue,
        importedRevision: reopened.revision,
      };
    }, operation);
    expect(result).toMatchObject({
      first: operation === 'import' ? 'INVALID_ARGUMENT' : 'WORKER_FAILED',
      retry: 'WORKER_FAILED',
      recoveredRevision: result.revision,
      value: 123,
      updatedValue: operation === 'import' ? 123 : 1,
      importedRevision: result.revision + 1,
    });
  });
}

for (const operation of [
  'import',
  'binary',
  'generate',
  'undo',
  'redo',
  'structural',
  'structural-column',
  'paste-input',
  'paste-style',
] as const) {
  test(`recovers a committed ${operation} after a post-head index failure`, async ({ page }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let armIndexFailure = false, failIndex = false, wantScratchFailure = false, failScratch = false;
addEventListener('message', e => { if (e.data.injectIndexFailure) { armIndexFailure = true; wantScratchFailure = !!e.data.failScratch; postMessage({ faultStage: 'armed' }); } });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (armIndexFailure && this.name === 'heads') {
    armIndexFailure = false;
    const tx = this.transaction;
    const originalComplete = tx.oncomplete;
    tx.oncomplete = function(event) {
      failIndex = true;
      postMessage({ faultStage: 'committed' });
      if (originalComplete) originalComplete.call(this, event);
    };
  }
  return originalPut.apply(this, args);
};
const originalMapSet = Map.prototype.set;
Map.prototype.set = function(...args) {
  if (failIndex) { failIndex = false; failScratch = wantScratchFailure; postMessage({ faultStage: 'index-failed' }); throw new Error('Injected index failure after commit'); }
  return originalMapSet.apply(this, args);
};
const originalDelete = IDBObjectStore.prototype.delete;
IDBObjectStore.prototype.delete = function(...args) {
  if (failScratch && this.name === 'scratch') { failScratch = false; postMessage({ faultStage: 'scratch-failed' }); throw new DOMException('Injected scratch cleanup failure', 'UnknownError'); }
  return originalDelete.apply(this, args);
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async (operation) => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[10]]);
      if (
        operation === 'binary' ||
        operation.startsWith('structural') ||
        operation.startsWith('paste')
      ) {
        await sheet.range('B2').setFormulas([['A1*2']]);
        await sheet.range('A1').setStyle({ bold: true, numberFormat: '0.00' });
      }
      const snapshot = await book.toJSON();
      const parts: BlobPart[] = [];
      if (operation === 'binary')
        for await (const part of book.streamBinary()) parts.push(part as BlobPart);
      const binary = new Blob(parts);
      await sheet.range('A1').setValues([[20]]);
      if (operation === 'redo') await book.undo();
      const revision = book.revision;
      const baselineStyle = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
      const stages: string[] = [];
      (book as any).worker.addEventListener('message', (event: MessageEvent) => {
        if (event.data.faultStage) stages.push(event.data.faultStage);
      });
      (book as any).worker.postMessage({
        injectIndexFailure: true,
        failScratch: operation === 'import' || operation === 'binary',
      });
      let code = '';
      try {
        if (operation === 'import') await book.importJSON(snapshot);
        else if (operation === 'binary') await book.importBinary(binary);
        else if (operation === 'structural') await sheet.insertRows(0, 1);
        else if (operation === 'structural-column') await sheet.insertColumns(0, 1);
        else if (operation === 'generate') await book.generate(10000);
        else if (operation.startsWith('paste'))
          await book.execute({
            type: 'core.cells.paste',
            payload: {
              sheetId: sheet.id,
              cells: [
                {
                  row: 0,
                  column: 0,
                  input: { type: 'number', value: 33 },
                  note: 'pasted',
                  ...(operation === 'paste-style' ? { style: { color: '#123456' } } : {}),
                },
              ],
            },
          });
        else if (operation === 'redo') await book.redo();
        else await book.undo();
      } catch (error) {
        code = (error as { code?: string }).code ?? '';
      }
      return {
        id: book.id,
        revision,
        code,
        stages,
        baselineStyle,
        state: book.persistenceState.state,
      };
    }, operation);
    expect(before.stages).toContain('committed');
    expect(before.stages).toContain('index-failed');
    if (operation === 'import' || operation === 'binary')
      expect(before.stages).toContain('scratch-failed');
    expect(before.code).toBe('WORKER_FAILED');
    expect(before.state).toBe('failed');
    const after = await page.evaluate(
      async ({ id, operation }) => {
        await window.opensheet.open(id);
        const book = window.opensheet.getWorkbook();
        const sheet = book.getSheets()[0];
        const value = (
          await sheet
            .range(
              operation === 'structural' ? 'A2' : operation === 'structural-column' ? 'B1' : 'A1',
            )
            .getValues()
        )[0][0];
        let detail: unknown;
        if (operation === 'binary' || operation.startsWith('structural')) {
          const row = operation === 'structural' ? 1 : 0;
          const column = operation === 'structural-column' ? 1 : 0;
          const cell = await book.getCell(sheet.id, row, column);
          const formula = await book.getCell(sheet.id, row + 1, column + 1);
          const formulaValue = (
            await sheet
              .range(
                operation === 'structural' ? 'B3' : operation === 'structural-column' ? 'C2' : 'B2',
              )
              .getValues()
          )[0][0];
          detail = { style: book.getStyle(cell?.styleId), formula: formula?.input, formulaValue };
        }
        const revision = book.revision;
        const stats = await book.storageStats();
        const open = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          open.onsuccess = () => resolve(open.result);
          open.onerror = () => reject(open.error);
        });
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
        const requests = [
          tx.objectStore('heads').get(id),
          tx.objectStore('records').getAll(range),
          tx.objectStore('sizes').getAll(range),
          tx.objectStore('scratch').getAllKeys(range),
        ];
        const values = await Promise.all(
          requests.map(
            (request) =>
              new Promise<any>((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error);
              }),
          ),
        );
        db.close();
        const actualBytes =
          (values[0] as Uint8Array).byteLength +
          [...values[1], ...values[2]].reduce(
            (sum: number, value: Uint8Array) => sum + value.byteLength,
            0,
          );
        const pasteDetail = operation.startsWith('paste')
          ? {
              style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
              note: await sheet.range('A1').getNotes(),
            }
          : undefined;
        let redo: boolean | undefined;
        let redone: unknown;
        let historyDetail: unknown;
        if (operation === 'redo') {
          redo = await book.undo();
          redone = (await book.getSheets()[0].range('A1').getValues())[0][0];
          const reapplied = await book.redo();
          historyDetail = { reapplied, value: (await sheet.range('A1').getValues())[0][0] };
        } else if (operation === 'undo') {
          redo = await book.redo();
          redone = (await book.getSheets()[0].range('A1').getValues())[0][0];
        } else if (operation.startsWith('paste')) {
          redo = await book.undo();
          redone = (await sheet.range('A1').getValues())[0][0];
          const undoStyle = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
          const undoNote = await sheet.range('A1').getNotes();
          const reapplied = await book.redo();
          historyDetail = {
            undoStyle,
            undoNote,
            reapplied,
            value: (await sheet.range('A1').getValues())[0][0],
            style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
            note: await sheet.range('A1').getNotes(),
          };
        } else if (operation.startsWith('structural')) {
          redo = await book.undo();
          redone = (await book.getSheets()[0].range('A1').getValues())[0][0];
          const undoFormula = (await book.getCell(sheet.id, 1, 1))?.input;
          const reapplied = await book.redo();
          const redoFormula = (
            await book.getCell(
              sheet.id,
              operation === 'structural' ? 2 : 1,
              operation === 'structural' ? 1 : 2,
            )
          )?.input;
          const redoValue = (
            await sheet.range(operation === 'structural' ? 'B3' : 'C2').getValues()
          )[0][0];
          historyDetail = { undoFormula, reapplied, redoFormula, redoValue };
        }
        return {
          value,
          detail,
          pasteDetail,
          historyDetail,
          revision,
          redo,
          redone,
          actualBytes,
          chargedBytes: stats.bytes,
          scratch: values[3],
        };
      },
      { id: before.id, operation },
    );
    expect(after.revision).toBe(before.revision + 1);
    expect(after.chargedBytes).toBe(after.actualBytes);
    expect(after.scratch).toEqual([]);
    expect(after.value).toBe(
      operation === 'generate'
        ? 0
        : operation === 'redo' || operation.startsWith('structural')
          ? 20
          : operation.startsWith('paste')
            ? 33
            : 10,
    );
    if (operation === 'redo') {
      expect(after.redo).toBe(true);
      expect(after.redone).toBe(10);
      expect(after.historyDetail).toEqual({ reapplied: true, value: 20 });
    }
    if (operation === 'binary' || operation.startsWith('structural')) {
      expect(after.detail).toMatchObject({
        style: { bold: true, numberFormat: '0.00' },
        formula: {
          type: 'formula',
          expression:
            operation === 'structural'
              ? "('Launch budget'!A2*2)"
              : operation === 'structural-column'
                ? "('Launch budget'!B1*2)"
                : 'A1*2',
        },
        formulaValue: operation.startsWith('structural') ? 40 : 20,
      });
    }
    if (operation.startsWith('structural')) {
      expect(after.redo).toBe(true);
      expect(after.redone).toBe(20);
      expect(after.historyDetail).toMatchObject({
        undoFormula: { type: 'formula', expression: 'A1*2' },
        reapplied: true,
        redoFormula: {
          type: 'formula',
          expression:
            operation === 'structural' ? "('Launch budget'!A2*2)" : "('Launch budget'!B1*2)",
        },
        redoValue: 40,
      });
    }
    if (operation.startsWith('paste')) {
      const style = operation === 'paste-style' ? { color: '#123456' } : before.baselineStyle;
      expect(after.pasteDetail).toEqual({ style, note: [['pasted']] });
      expect(after.redo).toBe(true);
      expect(after.redone).toBe(20);
      expect(after.historyDetail).toEqual({
        undoStyle: before.baselineStyle,
        undoNote: [[null]],
        reapplied: true,
        value: 33,
        style,
        note: [['pasted']],
      });
    }
    if (operation === 'undo') {
      expect(after.redo).toBe(true);
      expect(after.redone).toBe(20);
    }
  });
}

for (const operation of [
  'import',
  'binary',
  'generate',
  'undo',
  'redo',
  'structural',
  'structural-column',
  'paste-input',
  'paste-style',
] as const) {
  test(`head transaction abort preserves the old ${operation} revision`, async ({ page }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let abortNextHead = false;
addEventListener('message', e => { if (e.data.injectHeadAbort) abortNextHead = true; });
const originalHeadPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (abortNextHead && this.name === 'heads') {
    abortNextHead = false;
    this.transaction.abort();
    postMessage({faultStage: 'head-aborted'});
    throw new DOMException('Injected head transaction abort', 'AbortError');
  }
  return originalHeadPut.apply(this, args);
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async (operation) => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[10]]);
      if (
        operation === 'binary' ||
        operation.startsWith('structural') ||
        operation.startsWith('paste')
      ) {
        await sheet.range('B2').setFormulas([['A1*2']]);
        await sheet.range('A1').setStyle({ bold: true, numberFormat: '0.00' });
      }
      const snapshot = await book.toJSON();
      const parts: BlobPart[] = [];
      if (operation === 'binary')
        for await (const part of book.streamBinary()) parts.push(part as BlobPart);
      const binary = new Blob(parts);
      await sheet.range('A1').setValues([[20]]);
      if (operation === 'redo') await book.undo();
      const revision = book.revision;
      const baselineStyle = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
      const stages: string[] = [];
      (book as any).worker.addEventListener('message', (event: MessageEvent) => {
        if (event.data.faultStage) stages.push(event.data.faultStage);
      });
      (book as any).worker.postMessage({ injectHeadAbort: true });
      let code = '';
      try {
        if (operation === 'import') await book.importJSON(snapshot);
        else if (operation === 'binary') await book.importBinary(binary);
        else if (operation === 'structural') await sheet.insertRows(0, 1);
        else if (operation === 'structural-column') await sheet.insertColumns(0, 1);
        else if (operation === 'generate') await book.generate(10000);
        else if (operation.startsWith('paste'))
          await book.execute({
            type: 'core.cells.paste',
            payload: {
              sheetId: sheet.id,
              cells: [
                {
                  row: 0,
                  column: 0,
                  input: { type: 'number', value: 33 },
                  note: 'pasted',
                  ...(operation === 'paste-style' ? { style: { color: '#123456' } } : {}),
                },
              ],
            },
          });
        else if (operation === 'redo') await book.redo();
        else await book.undo();
      } catch (error) {
        code = (error as { code?: string }).code ?? '';
      }
      return { id: book.id, revision, afterFailure: book.revision, code, stages };
    }, operation);
    expect(before.stages).toContain('head-aborted');
    expect(before.code).toBe('AbortError');
    expect(before.afterFailure).toBe(before.revision);
    const after = await page.evaluate(
      async ({ id, operation }) => {
        await window.opensheet.open(id);
        const book = window.opensheet.getWorkbook();
        const sheet = book.getSheets()[0];
        const value = (await sheet.range('A1').getValues())[0][0];
        const cell = await book.getCell(sheet.id, 0, 0);
        const formula = await book.getCell(sheet.id, 1, 1);
        const detail = {
          style: book.getStyle(cell?.styleId),
          formula: formula?.input,
          formulaValue: (await sheet.range('B2').getValues())[0][0],
          note: await sheet.range('A1').getNotes(),
        };
        const revision = book.revision;
        const stats = await book.storageStats();
        const open = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          open.onsuccess = () => resolve(open.result);
          open.onerror = () => reject(open.error);
        });
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
        const requests = [
          tx.objectStore('heads').get(id),
          tx.objectStore('records').getAll(range),
          tx.objectStore('sizes').getAll(range),
          tx.objectStore('scratch').getAllKeys(range),
        ];
        const values = await Promise.all(
          requests.map(
            (request) =>
              new Promise<any>((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error);
              }),
          ),
        );
        db.close();
        const bytes =
          (values[0] as Uint8Array).byteLength +
          [...values[1], ...values[2]].reduce(
            (sum: number, item: Uint8Array) => sum + item.byteLength,
            0,
          );
        const recoveredRedo = operation === 'redo' ? await book.redo() : undefined;
        const recoveredRedoValue =
          operation === 'redo' ? (await sheet.range('A1').getValues())[0][0] : undefined;
        const undo = await book.undo();
        const undone = (await book.getSheets()[0].range('A1').getValues())[0][0];
        const redo = await book.redo();
        const redone = (await book.getSheets()[0].range('A1').getValues())[0][0];
        return {
          value,
          detail,
          revision,
          bytes,
          chargedBytes: stats.bytes,
          scratch: values[3],
          undo,
          undone,
          redo,
          redone,
          recoveredRedo,
          recoveredRedoValue,
        };
      },
      { id: before.id, operation },
    );
    expect(after).toMatchObject({
      value: operation === 'redo' ? 10 : 20,
      revision: before.revision,
      scratch: [],
      undo: true,
      undone: 10,
      redo: true,
      redone: 20,
    });
    if (operation === 'redo') {
      expect(after.recoveredRedo).toBe(true);
      expect(after.recoveredRedoValue).toBe(20);
    }
    expect(after.chargedBytes).toBe(after.bytes);
    if (
      operation === 'binary' ||
      operation.startsWith('structural') ||
      operation.startsWith('paste')
    )
      expect(after.detail).toMatchObject({
        style: { bold: true, numberFormat: '0.00' },
        formula: { type: 'formula', expression: 'A1*2' },
        formulaValue: 40,
      });
  });
}
