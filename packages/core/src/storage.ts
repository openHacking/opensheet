import { assert, OpenSheetError } from './types.js';
export { BLOCK_ROWS, BLOCK_COLUMNS } from './block-layout.js';
export const encodedBytes = (value: unknown) =>
  value instanceof Uint8Array
    ? value.byteLength + 1
    : new TextEncoder().encode(JSON.stringify(value)).byteLength;
export const recordBytes = (id: IDBValidKey, value: unknown) =>
  value instanceof Uint8Array ? value.byteLength + 1 : encodedBytes([id, value]) + 1;
export class ByteCache<T> {
  private entries = new Map<string, { value: T; bytes: number }>();
  bytes = 0;
  constructor(
    readonly budget: number,
    private onDelete?: (id: string) => void,
  ) {}
  get(id: string): T | undefined {
    const entry = this.entries.get(id);
    if (!entry) return;
    this.entries.delete(id);
    this.entries.set(id, entry);
    return entry.value;
  }
  set(id: string, value: T, bytes = encodedBytes(value) * 4) {
    this.delete(id);
    if (bytes > this.budget) return;
    this.entries.set(id, { value, bytes });
    this.bytes += bytes;
    while (this.bytes > this.budget) this.delete(this.entries.keys().next().value!);
  }
  delete(id: string) {
    const entry = this.entries.get(id);
    if (entry) this.bytes -= entry.bytes;
    this.entries.delete(id);
    if (entry) this.onDelete?.(id);
  }
  clear() {
    for (const id of this.entries.keys()) this.onDelete?.(id);
    this.entries.clear();
    this.bytes = 0;
  }
}
export const request = <T>(req: IDBRequest<T>) =>
  new Promise<T>((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
export function completed(tx: IDBTransaction) {
  return new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onabort = () =>
      reject(tx.error ?? new OpenSheetError('ABORTED', 'Storage transaction aborted'));
    tx.onerror = () => {}; // onabort carries the final result
  });
}
const encode = (id: IDBValidKey, value: unknown) => {
  const data =
    value instanceof Uint8Array ? value : new TextEncoder().encode(JSON.stringify([id, value]));
  const result = new Uint8Array(data.length + 1);
  result[0] = value instanceof Uint8Array ? 1 : 0;
  result.set(data, 1);
  return result;
};
const decode = <T>(value: unknown): T | undefined => {
  if (value === undefined) return undefined;
  assert(value instanceof Uint8Array, 'CORRUPT_STORAGE', 'Invalid encoded storage record');
  assert(value[0] === 0 || value[0] === 1, 'CORRUPT_STORAGE', 'Unsupported storage format');
  return value[0] === 1
    ? (value.subarray(1) as T)
    : (JSON.parse(new TextDecoder().decode(value.subarray(1)))[1] as T);
};
export class Database {
  private constructor(readonly db: IDBDatabase) {}
  static async open(name = 'opensheet') {
    try {
      assert(typeof indexedDB !== 'undefined', 'STORAGE_UNAVAILABLE', 'IndexedDB is required');
      const req = indexedDB.open(name, 1);
      req.onupgradeneeded = () => {
        req.result.createObjectStore('sizes');
        req.result.createObjectStore('heads');
        req.result.createObjectStore('records');
        req.result.createObjectStore('scratch');
      };
      const db = await request(req);
      db.onversionchange = () => db.close();
      return new Database(db);
    } catch (error) {
      if ((error as { name?: unknown })?.name === 'SecurityError')
        throw new OpenSheetError(
          'STORAGE_UNAVAILABLE',
          'IndexedDB access was denied; enable site storage or use a supported browser context',
        );
      throw error;
    }
  }
  get<T>(store: string, id: IDBValidKey) {
    return request<unknown>(this.db.transaction(store).objectStore(store).get(id)).then((value) =>
      decode<T>(value),
    );
  }
  async hasAny(store: string, range: IDBKeyRange) {
    return (
      (await request(this.db.transaction(store).objectStore(store).getKey(range))) !== undefined
    );
  }
  /** Compare raw record lengths with their ledger entries in bounded read transactions. */
  async recordSizes(ids: readonly string[]): Promise<number[]> {
    const result: number[] = [];
    for (let start = 0; start < ids.length; start += 128) {
      const batch = ids.slice(start, start + 128);
      const tx = this.db.transaction(['records', 'sizes']);
      const records = tx.objectStore('records'),
        sizes = tx.objectStore('sizes');
      await new Promise<void>((resolve, reject) => {
        let index = 0;
        const next = () => {
          if (index === batch.length) return resolve();
          const id = batch[index];
          const ledger = sizes.get(id);
          const record = records.get(id);
          let pending = 2;
          const check = () => {
            if (--pending !== 0) return;
            try {
              const size = decode<number>(ledger.result);
              assert(
                Number.isSafeInteger(size) && size! > 0,
                'CORRUPT_STORAGE',
                'Missing or invalid storage ledger; restore a backup',
              );
              assert(
                record.result instanceof Uint8Array && record.result.byteLength === size,
                'CORRUPT_STORAGE',
                'Record length differs from storage ledger; restore a backup',
              );
              result.push(size!);
              index++;
              next();
            } catch (error) {
              reject(error);
            }
          };
          ledger.onerror = () => reject(ledger.error);
          record.onerror = () => reject(record.error);
          ledger.onsuccess = record.onsuccess = check;
        };
        next();
      });
    }
    return result;
  }
  async write(
    store: string,
    puts: Array<[IDBValidKey, unknown]>,
    deletes: Array<IDBValidKey | IDBKeyRange> = [],
    signal?: AbortSignal,
  ) {
    signal?.throwIfAborted();
    const stores = store === 'records' ? [store, 'sizes'] : [store];
    const tx = this.db.transaction(stores, 'readwrite');
    const done = completed(tx);
    const target = tx.objectStore(store);
    const ledger = store === 'records' ? tx.objectStore('sizes') : undefined;
    let abortedBySignal = false;
    const abort = () => {
      try {
        tx.abort();
        abortedBySignal = true;
      } catch {
        // A completed or already-aborted transaction cannot be cancelled here.
      }
    };
    signal?.addEventListener('abort', abort, { once: true });
    if (signal?.aborted) abort();
    try {
      for (const id of deletes) {
        target.delete(id);
        ledger?.delete(id);
      }
      for (const [id, value] of puts) {
        target.put(encode(id, value), id);
        ledger?.put(encode(id, recordBytes(id, value)), id);
      }
      await done;
    } catch (error) {
      // Synchronous put/delete exceptions must not auto-commit earlier queued writes.
      try {
        tx.abort();
      } catch {
        /* Transaction may already be aborted. */
      }
      await done.catch(() => {});
      if (abortedBySignal) throw signal!.reason;
      throw error;
    } finally {
      signal?.removeEventListener('abort', abort);
    }
  }
  /** Remove one workbook atomically, including its size ledger and staging records. */
  async deleteWorkbook(id: string, range: IDBKeyRange) {
    const tx = this.db.transaction(['heads', 'records', 'sizes', 'scratch'], 'readwrite');
    const done = completed(tx);
    try {
      tx.objectStore('heads').delete(id);
      tx.objectStore('records').delete(range);
      tx.objectStore('sizes').delete(range);
      tx.objectStore('scratch').delete(range);
      await done;
    } catch (error) {
      try {
        tx.abort();
      } catch {
        /* Transaction may already be aborted. */
      }
      await done.catch(() => {});
      throw error;
    }
  }
  /** Read bounded key batches without decoding values. The consumer may delete yielded keys. */
  async *scanKeys(store: string, range: IDBKeyRange, batch = 256) {
    assert(Number.isSafeInteger(batch) && batch > 0, 'INVALID_ARGUMENT', 'Invalid scan batch');
    let remaining = range;
    while (true) {
      const target = this.db.transaction(store).objectStore(store);
      const keys = await request(target.getAllKeys(remaining, batch));
      if (!keys.length) return;
      const last = keys.at(-1)!;
      const finished =
        keys.length < batch ||
        (range.upper !== undefined && indexedDB.cmp(last, range.upper) === 0);
      yield keys;
      if (finished) return;
      remaining =
        range.upper === undefined
          ? IDBKeyRange.lowerBound(last, true)
          : IDBKeyRange.bound(last, range.upper, true, range.upperOpen);
    }
  }
  async *scan<T>(
    store: string,
    range: IDBKeyRange,
    batch = 32,
    maxBytes = Infinity,
  ): AsyncGenerator<Array<[IDBValidKey, T]>> {
    assert(
      Number.isSafeInteger(batch) && batch > 0 && maxBytes > 0,
      'INVALID_ARGUMENT',
      'Invalid scan batch',
    );
    let after: IDBValidKey | undefined;
    while (true) {
      const tx = this.db.transaction(store);
      const target = tx.objectStore(store);
      const rows: Array<[IDBValidKey, T]> = [];
      let bytes = 0;
      await new Promise<void>((resolve, reject) => {
        const req = target.openCursor(
          after === undefined
            ? range
            : range.upper === undefined
              ? IDBKeyRange.lowerBound(after, true)
              : IDBKeyRange.bound(after, range.upper, true, range.upperOpen),
        );
        req.onerror = () => reject(req.error);
        req.onsuccess = () => {
          try {
            const cursor = req.result;
            if (!cursor || rows.length === batch) {
              resolve();
              return;
            }
            const encoded = cursor.value;
            assert(
              encoded instanceof Uint8Array,
              'CORRUPT_STORAGE',
              'Invalid encoded storage record',
            );
            // Always allow one record so a large page cannot prevent cursor progress.
            if (rows.length && bytes + encoded.byteLength > maxBytes) {
              resolve();
              return;
            }
            rows.push([cursor.key, decode<T>(encoded)!]);
            bytes += encoded.byteLength;
            cursor.continue();
          } catch (error) {
            reject(error);
          }
        };
      });
      if (!rows.length) return;
      after = rows.at(-1)![0];
      yield rows;
      if (range.upper !== undefined && indexedDB.cmp(after, range.upper) === 0) return;
    }
  }
  close() {
    this.db.close();
  }
}
