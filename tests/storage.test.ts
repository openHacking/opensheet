import { encodeWorkbookFile, decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook, key, type Workbook } from '@opensheetjs/core';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import { Engine } from '../packages/core/src/engine.js';
import { ByteCache, Database, encodedBytes, recordBytes } from '../packages/core/src/storage.js';
import { streamExport } from '@opensheetjs/formats';
import {
  probeCapacity,
  percentile,
  type Measurement,
} from '../apps/playground/src/performance-model.js';
const books: Workbook[] = [];
const make = async (rows = 200, columns = 100, options = {}) => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows, columns }] }, options);
  books.push(book);
  return book;
};
afterEach(() => {
  for (const book of books.splice(0)) book.dispose();
});
describe('paged storage contracts', () => {
  it('reads raw cells, inputs and formula text without triggering calculation', async () => {
    const book = await make(),
      sheet = book.getSheets()[0];
    await sheet.range('B1').setValues([[2]]);
    await sheet.range('A1').setFormulas([['SUM(B1:B100)']]);
    const calculate = vi.spyOn(AsyncCalculation.prototype, 'value');
    try {
      expect((await book.getCell(sheet.id, 0, 0))!.input.type).toBe('formula');
      expect(await sheet.range('A1').getFormulas()).toEqual([['SUM(B1:B100)']]);
      expect(await sheet.range('A1').getInputs()).toEqual([
        [{ type: 'formula', expression: 'SUM(B1:B100)' }],
      ]);
      expect(calculate).not.toHaveBeenCalled();
      expect(await sheet.range('A1').getValues()).toEqual([[2]]);
      expect(calculate).toHaveBeenCalled();
    } finally {
      calculate.mockRestore();
    }
  });
  it('rejects large blank-range allocation before loading or creating a working set', async () => {
    const book = await make(10000, 100),
      revision = book.revision;
    const blocks = vi.spyOn(
      Engine.prototype as unknown as { block: (...args: unknown[]) => Promise<unknown> },
      'block',
    );
    try {
      await expect(
        book.getSheets()[0].range('A1:CV10000').setStyle({ bold: true }),
      ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
      expect(blocks).not.toHaveBeenCalled();
    } finally {
      blocks.mockRestore();
    }
    expect(book.revision).toBe(revision);
    expect(book.canUndo).toBe(false);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
  });
  it('charges the exact persisted UTF-8 payloads rather than a decoded-object estimate', async () => {
    const db = await Database.open(`codec-${crypto.randomUUID()}`),
      snapshot = createSnapshot(),
      engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    let actual = 0;
    for (const store of ['records', 'sizes', 'heads']) {
      const payloads = await new Promise<unknown[]>((resolve, reject) => {
        const request = db.db.transaction(store).objectStore(store).getAll();
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      for (const payload of payloads) {
        expect(payload).toBeInstanceOf(Uint8Array);
        actual += (payload as Uint8Array).byteLength;
      }
    }
    expect(engine.view().bytes).toBe(actual);
    engine.close();
  });
  it('cleans an interrupted initial creation even when no committed head exists', async () => {
    const db = await Database.open(`initial-${crypto.randomUUID()}`),
      snapshot = createSnapshot();
    const orphan = `${snapshot.workbookId}/orphan`;
    await db.write('records', [[orphan, { partial: true }]]);
    const engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    expect(await db.get('records', orphan)).toBeUndefined();
    engine.close();
  });
  it('aborts queued data writes if the matching byte-ledger write throws synchronously', async () => {
    const db = await Database.open(`atomic-${crypto.randomUUID()}`);
    const original = IDBObjectStore.prototype.put;
    const put = vi.spyOn(IDBObjectStore.prototype, 'put').mockImplementation(function (
      this: IDBObjectStore,
      value,
      key,
    ) {
      if (this.name === 'sizes')
        throw new DOMException('synchronous failure', 'QuotaExceededError');
      return original.call(this, value, key);
    });
    try {
      await expect(db.write('records', [['atomic/record', { value: 1 }]])).rejects.toMatchObject({
        name: 'QuotaExceededError',
      });
    } finally {
      put.mockRestore();
    }
    expect(await db.get('records', 'atomic/record')).toBeUndefined();
    expect(await db.get('sizes', 'atomic/record')).toBeUndefined();
    db.close();
  });
  it('streams and validates more than the former 200000-cell range boundary', async () => {
    const book = await make();
    await book.generate(210000);
    const blocks = vi.spyOn(
      Engine.prototype as unknown as { block: (...args: unknown[]) => Promise<unknown> },
      'block',
    );
    try {
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[0]]);
      expect(blocks.mock.calls.length).toBeLessThanOrEqual(4);
    } finally {
      blocks.mockRestore();
    }
    let count = 0;
    for await (const batch of book.getSheets()[0].range('A1:CV2100').stream({ values: false }))
      count += Object.keys(batch.cells).length;
    expect(count).toBe(210000);
  });
  it('keeps durable commits successful when post-commit reclamation fails', async () => {
    const book = await make();
    await book.generate(10000);
    const original = Database.prototype.write;
    const write = vi.spyOn(Database.prototype, 'write').mockImplementation(function (
      this: Database,
      store,
      puts = [],
      deletes = [],
    ) {
      if (store === 'records' && deletes.length)
        return Promise.reject(new Error('reclamation interrupted'));
      return original.call(this, store, puts, deletes);
    });
    try {
      await book.generate(20000);
    } finally {
      write.mockRestore();
    }
    const revision = book.revision,
      id = book.id;
    await book.close();
    const reopened = await openWorkbook(id);
    books.push(reopened);
    expect(reopened.revision).toBe(revision);
    expect((await reopened.getSheets()[0].range('A200').getValues())[0]).toHaveLength(1);
  });
  it('distinguishes unloaded, empty, zero and false across blocks, eviction and reopen', async () => {
    const book = await make(200, 100, { viewportCacheBytes: 2048 });
    const sheet = book.getSheets()[0];
    expect(book.peekValue(sheet.id, 63, 31)).toEqual({ error: '#LOADING' });
    await book.transaction({}, async (tx) => {
      await tx
        .getSheetById(sheet.id)!
        .range('AF64:AG65')
        .setValues([
          [0, false],
          ['', null],
        ]);
    });
    expect(await sheet.range('AF64:AG65').getValues()).toEqual([
      [0, false],
      ['', null],
    ]);
    await book.prefetch(sheet.id, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    expect(book.stats.viewportCacheBytes).toBeLessThanOrEqual(2048);
    const id = book.id;
    await book.close();
    const restored = await openWorkbook(id);
    books.push(restored);
    expect(await restored.getSheets()[0].range('AF64:AG65').getValues()).toEqual([
      [0, false],
      ['', null],
    ]);
  });
  it('evaluates unloaded cross-block and cross-sheet formulas, ranges, lazy branches and cycles', async () => {
    const book = await make();
    const data = book.getSheets()[0],
      other = await book.addSheet('Other', { rows: 200, columns: 100 });
    await data.range('AG65').setValues([[7]]);
    await other.range('A1').setFormulas([['SUM(Data!AF64:AG65)+IF(FALSE,1/0,3)']]);
    expect(await other.range('A1').getValues()).toEqual([[10]]);
    await data.range('AG65').setValues([[8]]);
    expect(await other.range('A1').getValues()).toEqual([[11]]);
    await data.range('A1:B1').setFormulas([['B1', 'A1']]);
    expect(await data.range('A1').getValues()).toEqual([[{ error: '#CYCLE!' }]]);
    const id = book.id;
    await book.close();
    const restored = await openWorkbook(id);
    books.push(restored);
    expect(await restored.getSheetByName('Other')!.range('A1').getValues()).toEqual([[11]]);
    expect(await restored.getSheetByName('Data')!.range('A1').getValues()).toEqual([
      [{ error: '#CYCLE!' }],
    ]);
  });
  it('isolates asynchronous transaction collectors from concurrent callers and cancellation', async () => {
    const book = await make();
    const sheet = book.getSheets()[0];
    let release!: () => void;
    const waiting = new Promise<void>((r) => {
      release = r;
    });
    const transaction = book.transaction({}, async (tx) => {
      await tx
        .getSheets()[0]
        .range('A1')
        .setValues([[9]]);
      await waiting;
    });
    await expect(sheet.range('A2').setValues([[2]])).rejects.toMatchObject({ code: 'BUSY' });
    book.cancel();
    release();
    await expect(transaction).rejects.toMatchObject({ code: 'ABORTED' });
    expect(await sheet.range('A1:A2').getValues()).toEqual([[null], [null]]);
  });
  it('does not publish failed or cancelled staging and cleans orphan records on recovery', async () => {
    const dbName = `recovery-${crypto.randomUUID()}`,
      snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 100, columns: 100 }] });
    let engine: Engine;
    let cancel = false;
    const db = await Database.open(dbName);
    engine = new Engine(db, snapshot.workbookId, undefined, () => {
      if (cancel) engine.cancel();
    });
    await engine.open(snapshot);
    cancel = true;
    await expect(engine.generate(10000, 256 * 1048576)).rejects.toMatchObject({ code: 'ABORTED' });
    engine.begin();
    expect(
      (
        await engine.read(snapshot.sheets[0].id, {
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 1,
        })
      ).calculated,
    ).toEqual({ [key(snapshot.sheets[0].rowOrder[0], snapshot.sheets[0].columnOrder[0])]: null });
    await db.write('records', [[`${snapshot.workbookId}/orphan`, { partial: true }]]);
    engine.close();
    const db2 = await Database.open(dbName),
      reopened = new Engine(db2, snapshot.workbookId);
    await reopened.open();
    expect(await db2.get('records', `${snapshot.workbookId}/orphan`)).toBeUndefined();
    reopened.close();
  });
  it('accounts encoded records exactly and retains the original workbook on budget failure', async () => {
    const dbName = `budget-${crypto.randomUUID()}`,
      snapshot = createSnapshot(),
      db = await Database.open(dbName),
      engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    const initial = engine.view().bytes;
    await expect(engine.generate(10000, initial + 4096)).rejects.toMatchObject({
      code: 'STORAGE_BUDGET',
    });
    expect(engine.view().snapshot.revision).toBe(0);
    expect(engine.view().bytes).toBe(initial);
    let actual = recordBytes(snapshot.workbookId, await db.get('heads', snapshot.workbookId));
    for await (const batch of db.scan(
      'records',
      IDBKeyRange.bound(`${snapshot.workbookId}/`, `${snapshot.workbookId}/\uffff`),
    ))
      for (const entry of batch) {
        actual += recordBytes(entry[0], entry[1]);
        const size = await db.get<number>('sizes', entry[0]);
        actual += recordBytes(entry[0], size);
      }
    expect(engine.view().bytes).toBe(actual);
    engine.close();
  });
  it('preserves edits across whole-block structural moves and bounded external sort', async () => {
    const book = await make(200, 100),
      sheet = book.getSheets()[0];
    await sheet.range('AF64:AG65').setValues([
      [3, 30],
      [1, 10],
    ]);
    await sheet.range('AH64:AH65').setFormulas([['AG64*2'], ['AG65*2']]);
    await sheet.insertRows(0);
    await sheet.sort('AF65:AH66', 31);
    expect(await sheet.range('AF65:AH66').getValues()).toEqual([
      [1, 10, 20],
      [3, 30, 60],
    ]);
    await book.undo();
    expect(await sheet.range('AF65:AH66').getValues()).toEqual([
      [3, 30, 60],
      [1, 10, 20],
    ]);
  });
  it('keeps cache reservations bounded and never retains an oversized entry', () => {
    const cache = new ByteCache<string>(10);
    cache.set('a', 'first', 6);
    cache.set('b', 'second', 6);
    expect(cache.get('a')).toBeUndefined();
    cache.set('huge', 'x', 11);
    expect(cache.bytes).toBe(6);
    cache.clear();
    expect(cache.bytes).toBe(0);
  });
  it('streams exports beyond the preview cap and preserves merges across batches', async () => {
    const book = await make(300, 100),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['Header']]);
    await sheet.range('A300').setValues([[0]]);
    await sheet.range('A160:B170').merge();
    for (const format of ['json', 'csv', 'tsv', 'markdown', 'html', 'latex'] as const) {
      let text = '';
      const stream =
        format === 'json'
          ? book.streamJSON()
          : streamExport(book, { sheetId: sheet.id, format, range: 'A1:CV300' });
      for await (const chunk of stream) text += chunk;
      if (format === 'json')
        expect(Object.values(decodeWorkbookFile(JSON.parse(text)).sheets[0].cells).length).toBe(2);
      else expect(text).toContain('0');
      if (format === 'html') {
        expect(text.match(/<table>/g)?.length).toBe(1);
        expect(text).toContain('rowspan="11"');
      }
      if (format === 'latex') {
        expect(text.match(/begin\{tabular\}/g)?.length).toBe(1);
        expect(text).toContain('multirow{11}');
      }
    }
  });
  it('rejects stale streaming revisions instead of combining versions', async () => {
    const book = await make(300, 100),
      sheet = book.getSheets()[0];
    const stream = book.streamRange(sheet.id, {
      startRow: 0,
      endRow: 300,
      startColumn: 0,
      endColumn: 100,
    });
    await stream.next();
    await sheet.range('A1').setValues([[1]]);
    await expect(stream.next()).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
  });
  it('checks envelope revisions and concurrent retries in the Worker and emits the requested id once', async () => {
    const book = await make(),
      sheet = book.getSheets()[0],
      listener = vi.fn();
    book.onCommit(listener);
    const command = {
      protocolVersion: 1 as const,
      commandId: 'once',
      workbookId: book.id,
      baseRevision: 0,
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 1 } }],
      },
    };
    const results = await Promise.all([book.execute(command), book.execute(command)]);
    expect(results[0]).toEqual(results[1]);
    expect(listener).toHaveBeenCalledTimes(1);
    expect(listener.mock.calls[0][0].commandId).toBe('once');
    const next = { ...command, commandId: 'next', baseRevision: 1 };
    const race = await Promise.allSettled([
      book.execute(next),
      book.execute({ ...next, commandId: 'stale' }),
    ]);
    expect(race.filter((r) => r.status === 'fulfilled')).toHaveLength(1);
    expect(race.filter((r) => r.status === 'rejected')).toHaveLength(1);
  });
  it('normalizes native quota errors rather than exposing the legacy numeric DOMException code', async () => {
    const book = await make();
    const write = vi
      .spyOn(Database.prototype, 'write')
      .mockRejectedValueOnce(new DOMException('quota full', 'QuotaExceededError'));
    try {
      await expect(book.generate(10000)).rejects.toMatchObject({ code: 'QuotaExceededError' });
    } finally {
      write.mockRestore();
    }
    expect(book.revision).toBe(0);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
  });
  it('terminates an unresponsive worker and rejects initialization within the configured deadline', async () => {
    let terminated = false;
    const dead = {
      postMessage() {},
      terminate() {
        terminated = true;
      },
    } as unknown as Worker;
    const { Workbook } = await import('@opensheetjs/core');
    await expect(
      new Workbook(encodeWorkbookFile(createSnapshot()), {
        workerFactory: () => dead,
        operationTimeoutMs: 5,
      }).ready(),
    ).rejects.toMatchObject({ code: 'WORKER_TIMEOUT' });
    expect(terminated).toBe(true);
  });
  it('rejects pending requests on disposal and reports unavailable storage', async () => {
    const book = await make();
    const pending = book.getCell(book.getSheets()[0].id, 0, 0);
    book.dispose();
    await expect(pending).rejects.toMatchObject({ code: 'DISPOSED' });
  });
});
describe('capacity search', () => {
  it('brackets and refines in 10000-cell steps without claiming untested maxima', async () => {
    const visited: number[] = [];
    const found = await probeCapacity(
      async (cells) => {
        visited.push(cells);
        return { cells, passed: cells <= 70000, rounds: [] } as Measurement;
      },
      () => false,
      100000,
    );
    expect(found.largestPassed).toBe(70000);
    expect(found.smallestFailed).toBe(80000);
    expect(new Set(visited).size).toBe(visited.length);
  });
  it('stops immediately on real quota failure, cancellation and dimension boundaries', async () => {
    const quota = await probeCapacity(
      async (cells) => ({ cells, passed: false, failure: 'QuotaExceededError', rounds: [] }),
      () => false,
    );
    expect(quota.results).toHaveLength(1);
    const boundary = await probeCapacity(
      async (cells) => ({ cells, passed: true, rounds: [] }),
      () => false,
      20000,
    );
    expect(boundary.results.at(-1)?.boundary).toBe('Dimension boundary');
    const stopped = await probeCapacity(
      async (cells) => ({ cells, passed: true, rounds: [] }),
      () => true,
    );
    expect(stopped.results).toHaveLength(0);
    const cancelled = await probeCapacity(
      async (cells) => ({ cells, passed: false, failure: 'ABORTED', rounds: [] }),
      () => false,
    );
    expect(cancelled.smallestFailed).toBeUndefined();
    expect(percentile([5, 1, 3])).toBe(5);
  });
});
