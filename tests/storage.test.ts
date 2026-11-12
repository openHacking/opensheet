import { encodeWorkbookFile, decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  createWorkbook,
  openWorkbook,
  key,
  type Workbook,
  type OperationTiming,
} from '@opensheetjs/core';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import { Engine } from '../packages/core/src/engine.js';
import { upgradeValidationFile } from '../packages/core/src/validation-file-migration.js';
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
  it('cancels an uncommitted record transaction with its ledger and preserves committed writes', async () => {
    const db = await Database.open(`abort-write-${crypto.randomUUID()}`);
    try {
      const pre = new AbortController();
      pre.abort();
      await expect(
        db.write('records', [['pre', { value: 1 }]], [], pre.signal),
      ).rejects.toMatchObject({ name: 'AbortError' });
      const abort = new AbortController();
      const pending = db.write(
        'records',
        [
          ['first', { value: 1 }],
          ['second', { value: 2 }],
        ],
        [],
        abort.signal,
      );
      abort.abort();
      await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
      for (const key of ['pre', 'first', 'second']) {
        expect(await db.get('records', key)).toBeUndefined();
        expect(await db.get('sizes', key)).toBeUndefined();
      }
      const after = new AbortController();
      await db.write('records', [['committed', { value: 3 }]], [], after.signal);
      after.abort();
      expect(await db.get('records', 'committed')).toEqual({ value: 3 });
      expect(await db.recordSizes(['committed'])).toEqual([recordBytes('committed', { value: 3 })]);
    } finally {
      db.close();
    }
  });
  it('persists exact sparse masks at both 32-bit row boundaries', async () => {
    const name = `page-mask-${crypto.randomUUID()}`;
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 64, columns: 32 }] });
    const sheet = snapshot.sheets[0];
    for (const [row, column] of [
      [0, 0],
      [31, 31],
      [32, 0],
      [63, 31],
    ]) {
      sheet.cells[key(row, column)] = {
        rowId: row,
        columnId: column,
        input: { type: 'number', value: row },
      };
    }
    const db = await Database.open(name);
    const engine = new Engine(db, snapshot.workbookId);
    try {
      await engine.open(snapshot);
      const head = await db.get<{ directory: Record<string, string> }>(
        'heads',
        snapshot.workbookId,
      );
      const directory = await db.get<Record<string, unknown>>(
        'records',
        Object.values(head!.directory)[0],
      );
      expect(directory?.[`${sheet.id}/0/0`]).toMatchObject({
        rowMask: '8000000180000001',
        columnMask: 0x80000001,
        count: 4,
      });
    } finally {
      engine.close();
    }
  });
  it('publishes a streamed rule-free backup atomically and preserves the old head on loss or commit failure', async () => {
    const name = `validation-rollback-${crypto.randomUUID()}`;
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 80, columns: 2 }] });
    const id = snapshot.workbookId;
    const db = await Database.open(name);
    const engine = new Engine(db, id);
    await engine.open(snapshot);
    const backup = upgradeValidationFile(encodeWorkbookFile(snapshot));
    backup.sheets[0].blocks.push({
      row: 0,
      column: 0,
      data: { numbers: { positions: [0], values: [42] } },
    });
    const baseline = await db.get('heads', id);
    const scope = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
    const source = () => new Blob([JSON.stringify(backup)]);
    backup.sheets[0].validation.push({
      id: 'rule',
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      rule: { kind: 'list', values: [42], allowBlank: false, policy: 'reject' },
    });
    await expect(engine.importRuleFreeValidationBackup(source())).rejects.toMatchObject({
      code: 'UNSUPPORTED_FEATURE',
    });
    expect(await db.get('heads', id)).toEqual(baseline);
    expect(await db.hasAny('scratch', scope)).toBe(false);
    backup.sheets[0].validation = [];
    const write = db.write.bind(db);
    const fault = vi.spyOn(db, 'write').mockImplementation(async (store, puts, deletes) => {
      if (store === 'heads') throw new DOMException('Injected quota', 'QuotaExceededError');
      return write(store, puts, deletes);
    });
    await expect(engine.importRuleFreeValidationBackup(source())).rejects.toMatchObject({
      name: 'QuotaExceededError',
    });
    fault.mockRestore();
    expect(await db.get('heads', id)).toEqual(baseline);
    expect(await db.hasAny('scratch', scope)).toBe(false);
    await engine.importRuleFreeValidationBackup(source());
    engine.close();
    const reopened = new Engine(await Database.open(name), id);
    try {
      const view = await reopened.open();
      expect(view.revision).toBe(1);
      const cells = await reopened.read(view.snapshot.sheets[0].id, {
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      expect(cells.calculated).toEqual({ '0:0': 42 });
    } finally {
      reopened.close();
    }
  });
  it('reads record and ledger from a consistent transaction during a concurrent update', async () => {
    const db = await Database.open(`ledger-snapshot-${crypto.randomUUID()}`);
    try {
      const id = 'book/record';
      const before = new Uint8Array([1, 2, 3]);
      const after = new Uint8Array([4, 5, 6, 7, 8, 9]);
      await db.write('records', [[id, before]]);
      const reading = db.recordSizes([id]);
      const writing = db.write('records', [[id, after]]);
      expect(await reading).toEqual([recordBytes(id, before)]);
      await writing;
      expect(await db.recordSizes([id])).toEqual([recordBytes(id, after)]);
    } finally {
      db.close();
    }
  });
  it('searches scalar cells directly while retaining formula and formatted-value matches', async () => {
    const book = await make();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A4').setValues([['Alpha'], [42], [false], [null]]);
    await sheet.range('A4').setFormulas([['2+3']]);
    await sheet.range('A2').setStyle({ numberFormat: '0.00' });
    const calculate = vi.spyOn(AsyncCalculation.prototype, 'value');
    try {
      expect(await book.search(sheet.id, 'ALPHA')).toEqual({
        count: 1,
        match: { row: 0, column: 0 },
      });
      expect(calculate).toHaveBeenCalledTimes(1);
      expect(await book.search(sheet.id, 'alpha', 0, { direction: 'previous' })).toEqual({
        count: 1,
        match: { row: 0, column: 0 },
      });
      expect(calculate).toHaveBeenCalledTimes(1);
      await sheet.range('A1').setValues([['changed']]);
      expect(await book.search(sheet.id, 'alpha')).toEqual({ count: 0, match: undefined });
      calculate.mockClear();
      expect(await book.search(sheet.id, '42.00')).toEqual({
        count: 1,
        match: { row: 1, column: 0 },
      });
      expect(calculate).toHaveBeenCalledTimes(1);
      expect(await book.search(sheet.id, '5')).toEqual({ count: 1, match: { row: 3, column: 0 } });
    } finally {
      calculate.mockRestore();
    }
  });
  it('tracks search positions after axis changes, undo and reopen', async () => {
    let book = await make();
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1:A3').setValues([['needle'], ['other'], ['needle']]);
      expect(await book.search(sheet.id, 'needle')).toEqual({
        count: 2,
        match: { row: 0, column: 0 },
      });
      await sheet.insertRows(0, 2);
      expect(await book.search(sheet.id, 'needle')).toEqual({
        count: 2,
        match: { row: 2, column: 0 },
      });
      await book.undo();
      expect(await book.search(sheet.id, 'needle')).toEqual({
        count: 2,
        match: { row: 0, column: 0 },
      });
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      sheet = book.getSheets()[0];
      expect(await book.search(sheet.id, 'needle')).toEqual({
        count: 2,
        match: { row: 2, column: 0 },
      });
      expect(await book.search(sheet.id, 'needle', 2 * 100)).toEqual({
        count: 2,
        match: { row: 4, column: 0 },
      });
    } finally {
      await book.close();
    }
  });
  it('replaces literal text atomically across pages and preserves history on overflow', async () => {
    let book = await make();
    try {
      const sheet = book.getSheets()[0];
      const large = 'a'.repeat(32768);
      await sheet.range('A1').setValues([['Alpha alpha']]);
      await sheet.range('A100').setValues([[large]]);
      await sheet.range('B1').setFormulas([['"Alpha"']]);
      const revision = book.revision;
      await expect(book.replaceText(sheet.id, 'a', 'aa')).rejects.toMatchObject({
        code: 'LIMIT_EXCEEDED',
      });
      expect(book.revision).toBe(revision);
      expect(await sheet.range('A1:B1').getValues()).toEqual([['Alpha alpha', 'Alpha']]);
      expect(await sheet.range('A100').getValues()).toEqual([[large]]);
      await book.replaceText(sheet.id, 'Alpha', 'Omega');
      expect(await sheet.range('A1:B1').getValues()).toEqual([['Omega alpha', 'Alpha']]);
      await book.undo();
      expect(await sheet.range('A1').getValues()).toEqual([['Alpha alpha']]);
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      expect(await book.getSheets()[0].range('A1:B1').getValues()).toEqual([
        ['Omega alpha', 'Alpha'],
      ]);
      expect(await book.getSheets()[0].range('A100').getValues()).toEqual([[large]]);
    } finally {
      await book.close();
    }
  });
  it('bounds replacement expansion before allocation and treats dollar sequences literally', async () => {
    const book = await make();
    const sheet = book.getSheets()[0];
    const large = 'a'.repeat(32768);
    await sheet.range('A1:A2').setValues([[large], ['aa']]);
    const revision = book.revision;
    const original = String.prototype.replaceAll;
    let dangerousCalls = 0;
    const replace = vi.spyOn(String.prototype, 'replaceAll').mockImplementation(function (
      this: string,
      search,
      replacement,
    ) {
      if (String(this) === large && search === 'a') {
        dangerousCalls++;
        throw new Error('unbounded allocation attempted');
      }
      return original.call(this, search, replacement);
    });
    try {
      await expect(book.replaceText(sheet.id, 'a', 'x'.repeat(32768))).rejects.toMatchObject({
        code: 'LIMIT_EXCEEDED',
      });
      expect(dangerousCalls).toBe(0);
    } finally {
      replace.mockRestore();
    }
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1').getValues()).toEqual([[large]]);
    await sheet.range('A1').setValues([['unchanged']]);
    await book.replaceText(sheet.id, 'a', '$&');
    expect(await sheet.range('A2').getValues()).toEqual([['$&$&']]);
    await book.undo();
    expect(await sheet.range('A2').getValues()).toEqual([['aa']]);
  });
  it('keeps revision and redo history for replacements that change no text', async () => {
    let book = await make();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([['alpha']]);
      await sheet.range('A2').setValues([['beta']]);
      await book.undo();
      const revision = book.revision;
      expect(book.canRedo).toBe(true);
      expect(await book.replaceText(sheet.id, 'missing', 'new')).toBeUndefined();
      expect(await book.replaceText(sheet.id, 'alpha', 'alpha')).toBeUndefined();
      const envelope = {
        protocolVersion: 1 as const,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision: revision,
        type: 'core.cells.replace' as const,
        payload: { sheetId: sheet.id, query: 'absent', replacement: 'new' },
      };
      expect(await book.execute(envelope)).toBeUndefined();
      (book as any).seen.clear(); // Exercise the Worker-side retry cache, not only the client cache.
      expect(await book.execute(envelope)).toBeUndefined();
      await expect(
        book.execute({ ...envelope, payload: { ...envelope.payload, query: 'other' } }),
      ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      expect(book.revision).toBe(revision);
      expect(book.canRedo).toBe(true);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      expect(book.revision).toBe(revision);
      expect(book.canRedo).toBe(true);
      expect(await book.redo()).toBe(true);
      expect(await book.getSheets()[0].range('A1:A2').getValues()).toEqual([['alpha'], ['beta']]);
      await expect(book.replaceText('missing-sheet', 'x', 'x')).rejects.toMatchObject({
        code: 'INVALID_ARGUMENT',
      });
    } finally {
      await book.close();
    }
  });
  it('opens an unversioned directory head without rewriting it and versions the next commit', async () => {
    const name = `head-version-${crypto.randomUUID()}`;
    const snapshot = createSnapshot();
    let db = await Database.open(name);
    let engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    const head = await db.get<Record<string, unknown>>('heads', snapshot.workbookId);
    expect(head?.engineFormatVersion).toBe(1);
    engine.close();
    db = await Database.open(name);
    delete head!.engineFormatVersion;
    await db.write('heads', [[snapshot.workbookId, head]]);
    engine = new Engine(db, snapshot.workbookId);
    await engine.open();
    expect(await db.get('heads', snapshot.workbookId)).toEqual(head);
    const actualBytes = async () => {
      let total = 0;
      for (const store of ['records', 'sizes', 'heads']) {
        const request = db.db.transaction(store).objectStore(store).getAll();
        const values = await new Promise<Uint8Array[]>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        total += values.reduce((sum, value) => sum + value.byteLength, 0);
      }
      return total;
    };
    expect(engine.view().bytes).toBe(await actualBytes());
    await engine.replace(snapshot);
    expect(
      (await db.get<Record<string, unknown>>('heads', snapshot.workbookId))?.engineFormatVersion,
    ).toBe(1);
    expect(engine.view().bytes).toBe(await actualBytes());
    engine.close();
  });
  it.each([
    [2, 'UNSUPPORTED_VERSION'],
    [null, 'CORRUPT_STORAGE'],
    ['1', 'CORRUPT_STORAGE'],
    [1.5, 'CORRUPT_STORAGE'],
    [0, 'CORRUPT_STORAGE'],
  ])('rejects engine version %s without reclaiming records', async (version, code) => {
    const name = `head-reject-${crypto.randomUUID()}`;
    const snapshot = createSnapshot();
    let db = await Database.open(name);
    let engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    const head = await db.get<Record<string, unknown>>('heads', snapshot.workbookId);
    engine.close();
    db = await Database.open(name);
    head!.engineFormatVersion = version;
    await db.write('heads', [[snapshot.workbookId, head]]);
    const orphan = `${snapshot.workbookId}/future-record`;
    await db.write('records', [[orphan, { retained: true }]]);
    engine = new Engine(db, snapshot.workbookId);
    try {
      await expect(engine.open()).rejects.toMatchObject({ code });
      expect(await db.get('heads', snapshot.workbookId)).toEqual(head);
      expect(await db.get('records', orphan)).toEqual({ retained: true });
    } finally {
      engine.close();
    }
  });
  it('reads scalar viewport values without routing them through formula calculation', async () => {
    const book = await make(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:F1').setValues([[0, false, '', null, 'text', -12.5]]);
    const cell = vi.spyOn(Engine.prototype, 'cell');
    const calculate = vi.spyOn(AsyncCalculation.prototype, 'value');
    try {
      expect(await sheet.range('A1:F1').getValues()).toEqual([[0, false, '', null, 'text', -12.5]]);
      expect(cell).not.toHaveBeenCalled();
      expect(calculate).not.toHaveBeenCalled();
    } finally {
      cell.mockRestore();
      calculate.mockRestore();
    }
    await sheet.range('G1').setFormulas([['SUM(A1:F1)']]);
    expect(await sheet.range('F1:G1').getValues()).toEqual([[-12.5, -12.5]]);
    await sheet.range('F1').setValues([[4]]);
    expect(await sheet.range('F1:G1').getValues()).toEqual([[4, 4]]);
    await book.undo();
    expect(await sheet.range('F1:G1').getValues()).toEqual([[-12.5, -12.5]]);
  });

  it('reuses a decoded page throughout a scalar range read', async () => {
    const book = await make(),
      sheet = book.getSheets()[0];
    await sheet
      .range('A1:AF64')
      .setValues(
        Array.from({ length: 64 }, (_, row) =>
          Array.from({ length: 32 }, (_, column) => row * 32 + column),
        ),
      );
    const blocks = vi.spyOn(
      Engine.prototype as unknown as { block: (...args: unknown[]) => Promise<unknown> },
      'block',
    );
    try {
      const values = await sheet.range('A1:AF64').getValues();
      expect(values.flat()).toEqual(Array.from({ length: 2048 }, (_, i) => i));
      expect(blocks).toHaveBeenCalledTimes(1);
    } finally {
      blocks.mockRestore();
    }
  });

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
  it('cleans an interrupted initial creation when explicitly allowed', async () => {
    const db = await Database.open(`initial-${crypto.randomUUID()}`),
      snapshot = createSnapshot();
    const orphan = `${snapshot.workbookId}/orphan`;
    await db.write('records', [[orphan, { partial: true }]]);
    const engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot, undefined, false);
    expect(await db.get('records', orphan)).toBeUndefined();
    engine.close();
  });
  it.each(['records', 'sizes', 'scratch'])(
    'preserves %s by default when the head is missing',
    async (store) => {
      const db = await Database.open(`missing-head-${crypto.randomUUID()}`),
        snapshot = createSnapshot();
      const orphan = `${snapshot.workbookId}/orphan`;
      await db.write(store, [[orphan, { partial: true }]]);
      const engine = new Engine(db, snapshot.workbookId);
      try {
        await expect(engine.open(snapshot)).rejects.toMatchObject({
          code: 'CORRUPT_STORAGE',
        });
        expect(await db.get(store, orphan)).toEqual({ partial: true });
        expect(await db.get('heads', snapshot.workbookId)).toBeUndefined();
      } finally {
        engine.close();
      }
    },
  );
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
  }, 30000);
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
  it('preserves the committed head if post-publication history loading fails', async () => {
    const book = await make();
    const sheet = book.getSheets()[0];
    const revision = book.revision;
    const writeOriginal = Database.prototype.write;
    const getOriginal = Database.prototype.get;
    let published = false;
    const write = vi.spyOn(Database.prototype, 'write').mockImplementation(async function (
      this: Database,
      store,
      puts = [],
      deletes = [],
    ) {
      await writeOriginal.call(this, store, puts, deletes);
      if (store === 'heads' && puts.length) published = true;
    });
    const get = vi.spyOn(Database.prototype, 'get').mockImplementation(function (
      this: Database,
      store,
      id,
    ) {
      if (published && store === 'records') {
        published = false;
        return Promise.reject(new Error('post-publication read failed'));
      }
      return getOriginal.call(this, store, id);
    });
    try {
      await expect(sheet.range('A1').setValues([[99]])).rejects.toMatchObject({
        code: 'WORKER_FAILED',
      });
      expect(book.persistenceState.state).toBe('failed');
    } finally {
      write.mockRestore();
      get.mockRestore();
    }
    const id = book.id;
    await book.close();
    const reopened = await openWorkbook(id);
    books.push(reopened);
    expect(reopened.revision).toBe(revision + 1);
    expect(await reopened.getSheets()[0].range('A1').getValues()).toEqual([[99]]);
    expect(await reopened.undo()).toBe(true);
    expect(await reopened.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
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
    const orphans = [
      ...Array.from({ length: 600 }, (_, i) => `${snapshot.workbookId}/orphan-${i}`),
      `${snapshot.workbookId}/\uffff`,
    ];
    await db.write(
      'records',
      orphans.map((id) => [id, { partial: true }]),
    );
    engine.close();
    const db2 = await Database.open(dbName),
      reopened = new Engine(db2, snapshot.workbookId);
    await reopened.open();
    for (const store of ['records', 'sizes']) {
      const surviving: IDBValidKey[] = [];
      for await (const batch of db2.scanKeys(
        store,
        IDBKeyRange.bound(`${snapshot.workbookId}/`, `${snapshot.workbookId}/\uffff`),
      ))
        surviving.push(...batch);
      expect(surviving.length).toBeGreaterThan(0);
      expect(surviving.some((id) => orphans.includes(String(id)))).toBe(false);
    }
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
  it('reports opt-in worker timings without letting observers break commits', async () => {
    const timings: { method: string; queueMs: number; executionMs: number; roundTripMs: number }[] =
      [];
    const errorLog = vi.spyOn(console, 'error').mockImplementation(() => {});
    const book = await make(200, 100, {
      onOperationTiming: (timing: (typeof timings)[number]) => {
        timings.push(timing);
        throw new Error('observer');
      },
    });
    try {
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[123]]);
      const timing = timings.find((item) => item.method === 'execute');
      expect(timing).toBeDefined();
      expect(timing!.queueMs).toBeGreaterThanOrEqual(0);
      expect(timing!.executionMs).toBeGreaterThanOrEqual(0);
      expect(timing!.roundTripMs).toBeGreaterThanOrEqual(0);
      expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[123]]);
      expect(errorLog).toHaveBeenCalled();
    } finally {
      errorLog.mockRestore();
    }
  });
  it('reports cold and absent page reads through opt-in Worker timing', async () => {
    const timings: OperationTiming[] = [];
    const book = await make(200, 100, {
      cacheBytes: 0,
      onOperationTiming: (value: OperationTiming) => timings.push(value),
    });
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[123]]);
    expect(await sheet.range('A1').getValues()).toEqual([[123]]);
    const cold = timings.filter((item) => item.method === 'read').at(-1)!.read!;
    expect(cold).toMatchObject({ cacheHits: 0, cacheMisses: 1, absentPages: 0 });
    expect(cold.storageMs).toBeGreaterThanOrEqual(0);
    expect(cold.decodeMs).toBeGreaterThanOrEqual(0);
    expect(cold.inflateMs + cold.pageDecodeMs).toBeCloseTo(cold.decodeMs, 5);
    expect(await sheet.range('A65').getValues()).toEqual([[null]]);
    const absent = timings.filter((item) => item.method === 'read').at(-1)!.read!;
    expect(absent).toMatchObject({
      cacheHits: 0,
      cacheMisses: 0,
      absentPages: 1,
      storageMs: 0,
      decodeMs: 0,
      inflateMs: 0,
      pageDecodeMs: 0,
    });
  });
  it('reports a cached page read through opt-in Worker timing', async () => {
    const timings: OperationTiming[] = [];
    const book = await make(200, 100, {
      onOperationTiming: (value: OperationTiming) => timings.push(value),
    });
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[123]]);
    expect(await sheet.range('A1').getValues()).toEqual([[123]]);
    expect(await sheet.range('A1').getValues()).toEqual([[123]]);
    expect(timings.filter((item) => item.method === 'read').at(-1)!.read).toMatchObject({
      cacheHits: 1,
      cacheMisses: 0,
      absentPages: 0,
      storageMs: 0,
      decodeMs: 0,
      inflateMs: 0,
      pageDecodeMs: 0,
    });
  });
  it('continues progress delivery and commits when an option callback throws', async () => {
    const errorLog = vi.spyOn(console, 'error').mockImplementation(() => {});
    const onProgress = vi.fn(() => {
      throw new Error('progress observer failed');
    });
    const book = await make(200, 100, { onProgress });
    const listener = vi.fn();
    const stop = book.onProgress(listener);
    try {
      await book.generate(1000);
      expect(onProgress).toHaveBeenCalled();
      expect(listener).toHaveBeenCalled();
      expect(errorLog).toHaveBeenCalled();
      expect(book.persistenceState.state).toBe('saved');
      expect(book.revision).toBeGreaterThan(0);
    } finally {
      stop();
      await book.close();
      errorLog.mockRestore();
    }
  });
  it('preserves the options receiver for progress and timing callbacks', async () => {
    const receivers: unknown[] = [];
    const options = {
      onProgress(this: unknown) {
        receivers.push(this);
      },
      onOperationTiming(this: unknown) {
        receivers.push(this);
      },
    };
    const book = await make(200, 100, options);
    await book.generate(1000);
    expect(receivers.length).toBeGreaterThan(1);
    expect(receivers.every((receiver) => receiver === options)).toBe(true);
  });
  it('isolates synchronous and asynchronous workbook observers from reads and commits', async () => {
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    const onProgress = vi.fn(async () => {
      throw new Error('progress option');
    });
    const onOperationTiming = vi.fn(async () => {
      throw new Error('timing option');
    });
    const book = await make(200, 100, { onProgress, onOperationTiming });
    const data = vi.fn(),
      progress = vi.fn(),
      commits = vi.fn(),
      persistence = vi.fn();
    const stops = [
      book.onData(() => {
        throw new Error('data sync');
      }),
      book.onData(async () => {
        throw new Error('data async');
      }),
      book.onData(data),
      book.onProgress(async () => {
        throw new Error('progress async');
      }),
      book.onProgress(progress),
      book.onCommit(async () => {
        throw new Error('commit async');
      }),
      book.onCommit(commits),
      book.onPersistence(async () => {
        throw new Error('persistence async');
      }),
      book.onPersistence(persistence),
    ];
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[4321]]);
      await book.prefetch(sheet.id, sheet.range('A1').bounds);
      expect(data).toHaveBeenCalled();
      expect(commits).toHaveBeenCalled();
      expect(persistence).toHaveBeenCalled();
      expect(await sheet.range('A1').getValues()).toEqual([[4321]]);
      await book.generate(1000);
      expect(progress).toHaveBeenCalled();
      expect(onProgress).toHaveBeenCalled();
      expect(onOperationTiming).toHaveBeenCalled();
      await vi.waitFor(() => expect(log.mock.calls.length).toBeGreaterThanOrEqual(5));
      expect(book.persistenceState.state).toBe('saved');
    } finally {
      stops.forEach((stop) => stop());
      await book.close();
      log.mockRestore();
    }
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
