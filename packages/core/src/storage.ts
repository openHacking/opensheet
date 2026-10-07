import { assert, OpenSheetError } from './types.js';
export const BLOCK_ROWS = 64;
export const BLOCK_COLUMNS = 32;
export const encodedBytes = (value: unknown) =>
  new TextEncoder().encode(JSON.stringify(value)).byteLength;
export class ByteCache<T> {
  private entries = new Map<string, { value: T; bytes: number }>();
  bytes = 0;
  constructor(readonly budget: number) {}
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
  }
  clear() {
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
const encode = (id: IDBValidKey, value: unknown) =>
  new TextEncoder().encode(JSON.stringify([id, value]));
const decode = <T>(value: unknown): T | undefined => {
  if (value === undefined) return undefined;
  assert(value instanceof Uint8Array, 'CORRUPT_STORAGE', 'Invalid encoded storage record');
  return JSON.parse(new TextDecoder().decode(value))[1] as T;
};
export class Database {
  private constructor(readonly db: IDBDatabase) {}
  static async open(name = 'opensheet-v4') {
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
  }
  get<T>(store: string, id: IDBValidKey) {
    return request<unknown>(this.db.transaction(store).objectStore(store).get(id)).then((value) =>
      decode<T>(value),
    );
  }
  async write(
    store: string,
    puts: Array<[IDBValidKey, unknown]>,
    deletes: Array<IDBValidKey | IDBKeyRange> = [],
  ) {
    const stores = store === 'records' ? [store, 'sizes'] : [store];
    const tx = this.db.transaction(stores, 'readwrite');
    const done = completed(tx);
    const target = tx.objectStore(store);
    const ledger = store === 'records' ? tx.objectStore('sizes') : undefined;
    try {
      for (const id of deletes) {
        target.delete(id);
        ledger?.delete(id);
      }
      for (const [id, value] of puts) {
        target.put(encode(id, value), id);
        ledger?.put(encode(id, encodedBytes([id, value])), id);
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
      throw error;
    }
  }
  async *scan<T>(
    store: string,
    range: IDBKeyRange,
    batch = 32,
  ): AsyncGenerator<Array<[IDBValidKey, T]>> {
    let after: IDBValidKey | undefined;
    while (true) {
      const tx = this.db.transaction(store);
      const target = tx.objectStore(store);
      const rows: Array<[IDBValidKey, T]> = [];
      await new Promise<void>((resolve, reject) => {
        const req = target.openCursor(
          after === undefined
            ? range
            : IDBKeyRange.bound(after, range.upper, true, range.upperOpen),
        );
        req.onerror = () => reject(req.error);
        req.onsuccess = () => {
          const cursor = req.result;
          if (!cursor || rows.length === batch) {
            resolve();
            return;
          }
          rows.push([cursor.key, decode<T>(cursor.value)!]);
          cursor.continue();
        };
      });
      if (!rows.length) return;
      after = rows.at(-1)![0];
      yield rows;
    }
  }
  close() {
    this.db.close();
  }
}
