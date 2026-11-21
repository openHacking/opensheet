/** Internal, transactional reference primitive. Not yet used by the playground routing. */
export interface SceneReference {
  value: string | null;
  revision: number;
}

const databaseName = 'opensheet-playground-references';
const storeName = 'references';

function validValue(value: unknown): value is string | null {
  return (
    value === null ||
    (typeof value === 'string' &&
      value.length > 0 &&
      value.length <= 100 &&
      /^[\w-]+$/.test(value) &&
      !['__proto__', 'constructor', 'prototype'].includes(value))
  );
}

function checkKey(key: string) {
  if (!/^[a-z][a-z0-9-]{0,63}$/.test(key)) throw new Error('Invalid scene reference key');
}

function checkedReference(value: unknown): SceneReference {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('Corrupt scene reference');
  const record = value as Record<string, unknown>;
  if (
    Object.keys(record).sort().join(',') !== 'revision,value' ||
    !validValue(record.value) ||
    !Number.isSafeInteger(record.revision) ||
    (record.revision as number) < 0
  )
    throw new Error('Corrupt scene reference');
  return { value: record.value, revision: record.revision as number };
}

async function openReferences(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(databaseName, 1);
    let blocked = false;
    request.onupgradeneeded = () => request.result.createObjectStore(storeName);
    request.onblocked = () => {
      blocked = true;
      reject(new Error('Scene reference database is blocked'));
    };
    request.onerror = () => reject(request.error ?? new Error('Cannot open scene references'));
    request.onsuccess = () => {
      const db = request.result;
      if (blocked) {
        db.close();
        return;
      }
      db.onversionchange = () => db.close();
      resolve(db);
    };
  });
}

/** Reads an existing record without initializing or consulting legacy storage. */
export async function readSceneReference(key: string): Promise<SceneReference | null> {
  checkKey(key);
  const db = await openReferences();
  return new Promise((resolve, reject) => {
    let result: SceneReference | null = null;
    let failure: unknown;
    try {
      const tx = db.transaction(storeName, 'readonly');
      const request = tx.objectStore(storeName).get(key);
      request.onsuccess = () => {
        try {
          result = request.result === undefined ? null : checkedReference(request.result);
        } catch (error) {
          failure = error;
          tx.abort();
        }
      };
      tx.oncomplete = () => {
        db.close();
        resolve(result);
      };
      tx.onabort = () => {
        db.close();
        reject(failure ?? tx.error ?? new Error('Scene reference transaction aborted'));
      };
      tx.onerror = () => {
        /* onabort reports the failure */
      };
    } catch (error) {
      db.close();
      reject(error);
    }
  });
}

/** Atomically initializes a missing record from a caller-captured legacy value. */
export async function initializeSceneReference(
  key: string,
  legacy: string | null,
): Promise<SceneReference> {
  checkKey(key);
  if (!validValue(legacy)) throw new Error('Invalid legacy scene reference');
  const result = await update(key, (current) => current ?? { value: legacy, revision: 0 });
  if (!result) throw new Error('Scene reference initialization failed');
  return result;
}

/** Compares the full versioned record, preventing a stale update after an ABA change. */
export async function compareSceneReference(
  key: string,
  expected: SceneReference,
  next: string | null,
): Promise<{ updated: boolean; current: SceneReference | null }> {
  checkKey(key);
  const snapshot = checkedReference(expected);
  if (!validValue(next)) throw new Error('Invalid scene reference value');
  let updated = false;
  const current = await update(key, (existing) => {
    if (!existing || existing.revision !== snapshot.revision || existing.value !== snapshot.value)
      return existing;
    if (existing.revision === Number.MAX_SAFE_INTEGER)
      throw new Error('Scene reference revision exhausted');
    updated = true;
    return { value: next, revision: existing.revision + 1 };
  });
  return { updated, current };
}

async function update(
  key: string,
  transform: (current: SceneReference | null) => SceneReference | null,
): Promise<SceneReference | null> {
  const db = await openReferences();
  return new Promise((resolve, reject) => {
    let result: SceneReference | null = null;
    let failure: unknown;
    let tx: IDBTransaction;
    try {
      tx = db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      const request = store.get(key);
      request.onsuccess = () => {
        try {
          const current = request.result === undefined ? null : checkedReference(request.result);
          result = transform(current);
          if (
            result &&
            (!current || result.value !== current.value || result.revision !== current.revision)
          )
            store.put(result, key);
        } catch (error) {
          failure = error;
          tx.abort();
        }
      };
      tx.oncomplete = () => {
        db.close();
        resolve(result);
      };
      tx.onabort = () => {
        db.close();
        reject(failure ?? tx.error ?? new Error('Scene reference transaction aborted'));
      };
      tx.onerror = () => {
        /* onabort reports the failure */
      };
    } catch (error) {
      db.close();
      reject(error);
    }
  });
}
