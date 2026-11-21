import { expect, it } from 'vitest';
import { Engine } from '../packages/core/src/engine.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { completed, Database } from '../packages/core/src/storage.js';

const wrap = (text: string) => new Uint8Array([0, ...new TextEncoder().encode(text)]);
async function putRaw(db: Database, store: string, key: IDBValidKey, value: Uint8Array) {
  const tx = db.db.transaction(store, 'readwrite');
  const done = completed(tx);
  tx.objectStore(store).put(value, key);
  await done;
}

it.each([
  ['invalid UTF-8', new Uint8Array([0, 0x5b, 0x22, 0xff, 0x22, 0x2c, 0x31, 0x5d])],
  ['invalid JSON', wrap('{')],
  ['object in place of tuple', wrap('{}')],
  ['missing tuple value', wrap('["entry"]')],
  ['extra tuple value', wrap('["entry",1,2]')],
  ['different tuple key', wrap('["other",1]')],
])('rejects a corrupt storage envelope with %s', async (_label, bytes) => {
  const db = await Database.open(`bad-envelope-${crypto.randomUUID()}`);
  try {
    await putRaw(db, 'heads', 'entry', bytes);
    await expect(db.get('heads', 'entry')).rejects.toMatchObject({ code: 'CORRUPT_STORAGE' });
    const stored = await new Promise<unknown>((resolve, reject) => {
      const request = db.db.transaction('heads').objectStore('heads').get('entry');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    expect(stored).toEqual(bytes);
  } finally {
    db.close();
  }
});

it.each([null, false, 0, ''])(
  'does not replace an existing invalid head %j or delete records',
  async (invalidHead) => {
    const db = await Database.open(`null-head-${crypto.randomUUID()}`);
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 3, columns: 3 }] });
    const engine = new Engine(db, snapshot.workbookId);
    const id = `${snapshot.workbookId}/existing`;
    try {
      await db.write('heads', [[snapshot.workbookId, invalidHead]]);
      await db.write('records', [[id, { kept: true }]]);
      await expect(engine.open(snapshot, 1024 * 1024, false)).rejects.toMatchObject({
        code: 'CORRUPT_STORAGE',
      });
      expect(await db.get('heads', snapshot.workbookId)).toBe(invalidHead);
      expect(await db.get('records', id)).toEqual({ kept: true });
    } finally {
      engine.close();
    }
  },
);

it('retains valid JSON and binary records for string, numeric and array keys', async () => {
  const db = await Database.open(`valid-envelope-${crypto.randomUUID()}`);
  try {
    for (const id of ['entry', 23, ['array', 42]] as IDBValidKey[]) {
      await db.write('heads', [[id, { text: '你好' }]]);
      expect(await db.get('heads', id)).toEqual({ text: '你好' });
    }
    await db.write('records', [['bytes', new Uint8Array([0, 128, 255])]]);
    expect(await db.get('records', 'bytes')).toEqual(new Uint8Array([0, 128, 255]));
    expect(await db.get('heads', 'absent')).toBeUndefined();
  } finally {
    db.close();
  }
});

it('rejects wrong envelope keys in ledger reads and cursor scans', async () => {
  const db = await Database.open(`bad-envelope-paths-${crypto.randomUUID()}`);
  try {
    await db.write('records', [['entry', { value: 1 }]]);
    await putRaw(db, 'sizes', 'entry', wrap('["different",123]'));
    await expect(db.recordSizes(['entry'])).rejects.toMatchObject({ code: 'CORRUPT_STORAGE' });
    await putRaw(db, 'heads', 'entry', wrap('["different",{"revision":1}]'));
    await expect(db.scan('heads', IDBKeyRange.only('entry')).next()).rejects.toMatchObject({
      code: 'CORRUPT_STORAGE',
    });
  } finally {
    db.close();
  }
});
