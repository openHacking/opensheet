import { expect, test } from '@playwright/test';

test('unversioned head survives an aborted first publication and upgrades on retry', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let abortNextHead = false;
addEventListener('message', e => { if (e.data.injectHeadAbort) abortNextHead = true; });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (abortNextHead && this.name === 'heads') {
    abortNextHead = false;
    this.transaction.abort();
    postMessage({faultStage: 'head-aborted'});
    throw new DOMException('Injected head transaction abort', 'AbortError');
  }
  return originalPut.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const fixture = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[10]]);
    await sheet.range('B1').setFormulas([['A1*2']]);
    await sheet.range('A1').setStyle({ bold: true });
    await sheet.range('A1').setValues([[20]]);
    const id = book.id;
    const revision = book.revision;
    await book.close();
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('opensheet');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const raw = await new Promise<Uint8Array>((resolve, reject) => {
      const request = db.transaction('heads').objectStore('heads').get(id);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const tuple = JSON.parse(new TextDecoder().decode(raw.subarray(1)));
    delete tuple[1].engineFormatVersion;
    const encoded = new TextEncoder().encode(JSON.stringify(tuple));
    const legacy = new Uint8Array(encoded.length + 1);
    legacy.set(encoded, 1);
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').put(legacy, id);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    return { id, revision, legacy: Array.from(legacy) };
  });
  const result = await page.evaluate(async ({ id }) => {
    const readHead = async () => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('opensheet');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const value = await new Promise<Uint8Array>((resolve, reject) => {
        const request = db.transaction('heads').objectStore('heads').get(id);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      db.close();
      return Array.from(value);
    };
    await window.opensheet.open(id);
    let book = window.opensheet.getWorkbook();
    const afterOpen = await readHead();
    const stages: string[] = [];
    (book as any).worker.addEventListener('message', (event: MessageEvent) => {
      if (event.data.faultStage) stages.push(event.data.faultStage);
    });
    (book as any).worker.postMessage({ injectHeadAbort: true });
    let code = '';
    try {
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[30]]);
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    const afterAbort = await readHead();
    await window.opensheet.open(id);
    book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const revision = book.revision;
    const values = await sheet.range('A1:B1').getValues();
    const formula = await sheet.range('B1').getFormulas();
    const cell = await book.getCell(sheet.id, 0, 0);
    const style = book.getStyle(cell?.styleId);
    const accounting = async () => {
      const stats = await book.storageStats();
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('opensheet');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
      const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
      const requests = [
        tx.objectStore('heads').get(id),
        tx.objectStore('records').getAll(range),
        tx.objectStore('sizes').getAll(range),
        tx.objectStore('scratch').getAllKeys(range),
      ];
      const data = await Promise.all(
        requests.map(
          (request) =>
            new Promise<any>((resolve, reject) => {
              request.onsuccess = () => resolve(request.result);
              request.onerror = () => reject(request.error);
            }),
        ),
      );
      db.close();
      return {
        charged: stats.bytes,
        actual:
          (data[0] as Uint8Array).byteLength +
          [...data[1], ...data[2]].reduce(
            (sum: number, value: Uint8Array) => sum + value.byteLength,
            0,
          ),
        scratch: data[3],
      };
    };
    const beforeUpgrade = await accounting();
    const undo = await book.undo();
    const undone = await sheet.range('A1:B1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1:B1').getValues();
    const head = new Uint8Array(await readHead());
    const version = JSON.parse(new TextDecoder().decode(head.subarray(1)))[1].engineFormatVersion;
    const afterUpgrade = await accounting();
    await window.opensheet.open(id);
    const reopened = await window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues();
    return {
      afterOpen,
      afterAbort,
      stages,
      code,
      revision,
      values,
      formula,
      style,
      beforeUpgrade,
      afterUpgrade,
      undo,
      undone,
      redo,
      redone,
      version,
      reopened,
    };
  }, fixture);
  expect(result.afterOpen).toEqual(fixture.legacy);
  expect(result.afterAbort).toEqual(fixture.legacy);
  expect(result.beforeUpgrade).toEqual({
    charged: result.beforeUpgrade.actual,
    actual: result.beforeUpgrade.actual,
    scratch: [],
  });
  expect(result.afterUpgrade).toEqual({
    charged: result.afterUpgrade.actual,
    actual: result.afterUpgrade.actual,
    scratch: [],
  });
  expect(result).toMatchObject({
    stages: ['head-aborted'],
    code: 'AbortError',
    revision: fixture.revision,
    values: [[20, 40]],
    formula: [['A1*2']],
    style: { bold: true },
    undo: true,
    undone: [[10, 20]],
    redo: true,
    redone: [[20, 40]],
    version: 1,
    reopened: [[20, 40]],
  });
});

for (const variant of [
  'future',
  'unknown-versioned',
  'unknown-unversioned',
  'invalid-directory',
  'invalid-bucket',
  'foreign-directory',
  'invalid-directory-record',
  'invalid-pointer',
  'foreign-pointer',
  'wrong-bucket',
  'unknown-pointer-field',
  'unknown-sheet',
  'unknown-history-field',
  'foreign-history-reference',
  'invalid-history-pointer',
  'invalid-history-address',
  'history-accounting',
] as const) {
  test(`${variant} engine heads are rejected before any persisted records are reclaimed`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const fixture = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[456]]);
      const id = book.id;
      const sheetId = book.getSheets()[0].id;
      await book.close();
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('opensheet');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const head = await new Promise<Uint8Array>((resolve, reject) => {
        const request = db.transaction('heads').objectStore('heads').get(id);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const tuple = JSON.parse(new TextDecoder().decode(head.subarray(1)));
      if (kind === 'future') tuple[1].engineFormatVersion = 99;
      else if (
        [
          'unknown-history-field',
          'foreign-history-reference',
          'invalid-history-pointer',
          'invalid-history-address',
          'history-accounting',
        ].includes(kind)
      ) {
        tuple[1].past = [`${id}/future-record`];
        tuple[1].future = [];
        tuple[1].historyBytes = kind === 'history-accounting' ? 1 : 0;
      } else if (kind === 'invalid-directory') tuple[1].directory = [];
      else if (kind === 'invalid-bucket') tuple[1].directory = { '256': `${id}/future-record` };
      else if (kind === 'foreign-directory') tuple[1].directory = { '0': 'other/record' };
      else if (
        [
          'invalid-directory-record',
          'invalid-pointer',
          'foreign-pointer',
          'wrong-bucket',
          'unknown-pointer-field',
          'unknown-sheet',
        ].includes(kind)
      )
        tuple[1].directory = { '0': `${id}/future-record` };
      else {
        if (kind === 'unknown-unversioned') delete tuple[1].engineFormatVersion;
        tuple[1].requiredFutureData = { record: `${id}/future-record` };
      }
      const body = new TextEncoder().encode(JSON.stringify(tuple));
      const replacement = new Uint8Array(body.length + 1);
      replacement.set(body, 1);
      const orphan = `${id}/future-record`;
      let recordBody = new Uint8Array([42]);
      if (kind === 'invalid-directory-record')
        recordBody = new TextEncoder().encode(JSON.stringify([orphan, []]));
      else if (
        [
          'unknown-history-field',
          'foreign-history-reference',
          'invalid-history-pointer',
          'invalid-history-address',
          'history-accounting',
        ].includes(kind)
      ) {
        const history = {
          before: kind === 'foreign-history-reference' ? 'other/record' : tuple[1].meta,
          after: tuple[1].meta,
          changes:
            kind === 'invalid-history-pointer' || kind === 'invalid-history-address'
              ? [
                  [
                    `${sheetId}/${kind === 'invalid-history-address' ? '999999999999' : '0'}/0`,
                    null,
                    {
                      id: `${id}/page`,
                      bytes: kind === 'invalid-history-pointer' ? -1 : 100,
                      count: 1,
                      endRow: 1,
                      endColumn: 1,
                    },
                  ],
                ]
              : [],
          bytes: 0,
          ranges: [],
          ...(kind === 'unknown-history-field'
            ? { futureReference: `${id}/important-record` }
            : {}),
        };
        recordBody = new TextEncoder().encode(JSON.stringify([orphan, history]));
      } else if (
        [
          'invalid-pointer',
          'foreign-pointer',
          'wrong-bucket',
          'unknown-pointer-field',
          'unknown-sheet',
        ].includes(kind)
      ) {
        let row = 0;
        const bucket = (address: string) => {
          let hash = 2166136261;
          for (const c of address) hash = Math.imul(hash ^ c.charCodeAt(0), 16777619);
          return (hash >>> 0) % 256;
        };
        const pageSheet = kind === 'unknown-sheet' ? 'missing-sheet' : sheetId;
        while (bucket(`${pageSheet}/${row}/0`) !== (kind === 'wrong-bucket' ? 1 : 0)) row++;
        const pointer = {
          id: kind === 'foreign-pointer' ? 'other/record' : `${id}/page`,
          bytes: 100,
          count: 1,
          endRow: 1,
          endColumn: 1,
        };
        const value =
          kind === 'invalid-pointer'
            ? null
            : kind === 'unknown-pointer-field'
              ? { ...pointer, future: true }
              : pointer;
        recordBody = new TextEncoder().encode(
          JSON.stringify([orphan, { [`${pageSheet}/${row}/0`]: value }]),
        );
      }
      const orphanBytes = new Uint8Array(recordBody.length + 1);
      orphanBytes.set(recordBody, 1);
      const tx = db.transaction(['heads', 'records'], 'readwrite');
      tx.objectStore('heads').put(replacement, id);
      tx.objectStore('records').put(orphanBytes, orphan);
      await new Promise<void>((resolve, reject) => {
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id, orphan, head: Array.from(replacement), orphanBytes: Array.from(orphanBytes) };
    }, variant);
    const other = await context.newPage();
    await other.goto('/#budget');
    await expect(other.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await expect(other.locator('#startup-recovery-message')).toContainText(
      variant === 'future'
        ? 'UNSUPPORTED_VERSION'
        : variant === 'history-accounting'
          ? 'Workbook history accounting differs'
          : [
                'unknown-history-field',
                'foreign-history-reference',
                'invalid-history-pointer',
                'invalid-history-address',
              ].includes(variant)
            ? 'Invalid workbook history'
            : variant === 'unknown-sheet'
              ? 'Block directory references an unknown sheet'
              : [
                    'invalid-directory-record',
                    'invalid-pointer',
                    'foreign-pointer',
                    'wrong-bucket',
                    'unknown-pointer-field',
                  ].includes(variant)
                ? 'Invalid block directory'
                : ['invalid-directory', 'invalid-bucket', 'foreign-directory'].includes(variant)
                  ? 'Unsupported workbook database format'
                  : 'Unknown workbook head fields',
    );
    expect(await other.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
      fixture.id,
    );
    const state = await other.evaluate(
      async ({ id, orphan }) => {
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          const request = indexedDB.open('opensheet');
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const data = await new Promise<Uint8Array>((resolve, reject) => {
          const request = db.transaction('heads').objectStore('heads').get(id);
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const record = await new Promise<Uint8Array>((resolve, reject) => {
          const request = db.transaction('records').objectStore('records').get(orphan);
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        db.close();
        return { head: Array.from(data), record: Array.from(record) };
      },
      { id: fixture.id, orphan: fixture.orphan },
    );
    expect(state).toEqual({ head: fixture.head, record: fixture.orphanBytes });
    await other.close();
  });
}

test('reopening after undo preserves redo history accounting', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const id = book.id;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[10]]);
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[20]]);
    await book.undo();
    await app.open(id);
    book = app.getWorkbook();
    const before = await book.getSheets()[0].range('A1').getValues();
    const redo = await book.redo();
    const after = await book.getSheets()[0].range('A1').getValues();
    return { before, redo, after };
  });
  expect(result).toEqual({ before: [[10]], redo: true, after: [[20]] });
});

test('unsupported native versions are explicit and leave current data and history intact', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const fixture = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const original = await sheet.range('A1').getValues();
    await sheet.range('A1').setValues([[987]]);
    const file = await book.toJSON();
    const parts: BlobPart[] = [];
    for await (const chunk of book.streamBinary()) parts.push(chunk as BlobPart);
    const binary = new Uint8Array(await new Blob(parts).arrayBuffer());
    binary[3] = 4;
    return {
      id: book.id,
      revision: book.revision,
      original,
      json: JSON.stringify({ ...file, schemaVersion: 4 }),
      binary: Array.from(binary),
    };
  });
  await page.locator('#file').setInputFiles({
    name: 'future.json',
    mimeType: 'application/json',
    buffer: Buffer.from(fixture.json),
  });
  await expect(page.locator('#notice')).toContainText('Workbook schema 4 is unsupported');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(fixture.id);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(fixture.revision);
  await page.locator('#file').setInputFiles({
    name: 'future.opensheet',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(fixture.binary),
  });
  await expect(page.locator('#notice')).toContainText('Binary workbook version 4 is unsupported');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const value = await book.getSheets()[0].range('A1').getValues();
    const revision = book.revision;
    const undo = await book.undo();
    const undone = await book.getSheets()[0].range('A1').getValues();
    return { id: book.id, revision, value, undo, undone };
  });
  expect(result).toEqual({
    id: fixture.id,
    revision: fixture.revision,
    value: [[987]],
    undo: true,
    undone: fixture.original,
  });
});

test('unknown nested native fields cannot replace the current workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const fixture = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const original = await sheet.range('A1').getValues();
    await sheet.range('A1').setValues([[987]]);
    const file = structuredClone(await book.toJSON());
    Object.assign(file.sheets[0].freeze, { unexpected: true });
    return { id: book.id, revision: book.revision, original, json: JSON.stringify(file) };
  });
  await page.locator('#file').setInputFiles({
    name: 'unknown.json',
    mimeType: 'application/json',
    buffer: Buffer.from(fixture.json),
  });
  await expect(page.locator('#notice')).toContainText('Invalid workbook file');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const value = await sheet.range('A1').getValues(),
      revision = book.revision;
    const undo = await book.undo();
    return { id: book.id, revision, value, undo, undone: await sheet.range('A1').getValues() };
  });
  expect(result).toEqual({
    id: fixture.id,
    revision: fixture.revision,
    value: [[987]],
    undo: true,
    undone: fixture.original,
  });
});

test('corrupt binary directory marker cannot replace the current workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const fixture = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const original = await sheet.range('A1').getValues();
    await sheet.range('A1').setValues([[987]]);
    const parts: BlobPart[] = [];
    for await (const chunk of book.streamBinary()) parts.push(chunk as BlobPart);
    const binary = new Uint8Array(await new Blob(parts).arrayBuffer());
    const footer = new DataView(binary.buffer, binary.length - 24, 24);
    const directory = Number(footer.getBigUint64(4, true));
    binary[directory - 8] ^= 1;
    return { id: book.id, revision: book.revision, original, binary: Array.from(binary) };
  });
  await page.locator('#file').setInputFiles({
    name: 'damaged.opensheet',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(fixture.binary),
  });
  await expect(page.locator('#notice')).toContainText('Unknown binary page sheet');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const value = await sheet.range('A1').getValues(),
      revision = book.revision;
    const undo = await book.undo();
    return { id: book.id, revision, value, undo, undone: await sheet.range('A1').getValues() };
  });
  expect(result).toEqual({
    id: fixture.id,
    revision: fixture.revision,
    value: [[987]],
    undo: true,
    undone: fixture.original,
  });
});

test('future persisted metadata is rejected without rewriting its head or records', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const fixture = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[321]]);
    const id = book.id;
    await book.close();
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const read = (store: string, key: IDBValidKey) =>
      new Promise<Uint8Array>((resolve, reject) => {
        const request = db.transaction(store).objectStore(store).get(key);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const head = await read('heads', id);
    const headObject = JSON.parse(new TextDecoder().decode(head.subarray(1)))[1];
    const metadataId = headObject.meta as string;
    const metadata = await read('records', metadataId);
    const tuple = JSON.parse(new TextDecoder().decode(metadata.subarray(1)));
    tuple[1].snapshot.schemaVersion = 4;
    const body = new TextEncoder().encode(JSON.stringify(tuple));
    const future = new Uint8Array(body.length + 1);
    future[0] = 0;
    future.set(body, 1);
    const tx = db.transaction('records', 'readwrite');
    tx.objectStore('records').put(future, metadataId);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    const list = db.transaction('records').objectStore('records').getAllKeys();
    const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
      list.onsuccess = () => resolve(list.result);
      list.onerror = () => reject(list.error);
    });
    db.close();
    return {
      id,
      metadataId,
      head: Array.from(head),
      future: Array.from(future),
      keys: keys.map(String).sort(),
    };
  });
  const other = await context.newPage();
  await other.goto('/#budget');
  await expect(other.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(other.locator('#startup-recovery-message')).toContainText('UNSUPPORTED_VERSION');
  expect(await other.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
    fixture.id,
  );
  const after = await other.evaluate(
    async ({ id, metadataId }) => {
      const open = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
      });
      const tx = db.transaction(['heads', 'records']);
      const requests = [
        tx.objectStore('heads').get(id),
        tx.objectStore('records').get(metadataId),
        tx.objectStore('records').getAllKeys(),
      ];
      const data = await Promise.all(
        requests.map(
          (request) =>
            new Promise<any>((resolve, reject) => {
              request.onsuccess = () => resolve(request.result);
              request.onerror = () => reject(request.error);
            }),
        ),
      );
      db.close();
      return {
        head: Array.from(data[0] as Uint8Array),
        future: Array.from(data[1] as Uint8Array),
        keys: (data[2] as IDBValidKey[]).map(String).sort(),
      };
    },
    { id: fixture.id, metadataId: fixture.metadataId },
  );
  expect(after).toEqual({ head: fixture.head, future: fixture.future, keys: fixture.keys });
  await other.close();
});

for (const location of ['current', 'undo', 'redo'] as const) {
  test(`unsupported ${location} metadata is rejected before recovery touches records or scratch`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const fixture = await page.evaluate(async (location) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[246]]);
      if (location !== 'current') await book.renameSheet(book.getSheets()[0].id, 'Renamed');
      if (location === 'redo') await book.undo();
      const id = book.id;
      await book.close();
      const open = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
      });
      const get = (store: string, key: string) =>
        new Promise<Uint8Array>((resolve, reject) => {
          const request = db.transaction(store).objectStore(store).get(key);
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
      const head = await get('heads', id);
      const decode = (bytes: Uint8Array) => JSON.parse(new TextDecoder().decode(bytes.subarray(1)));
      const state = decode(head)[1];
      const metadataId =
        location === 'current'
          ? (state.meta as string)
          : location === 'undo'
            ? (decode(await get('records', state.past.at(-1)))[1].before as string)
            : (decode(await get('records', state.future.at(-1)))[1].after as string);
      const metadata = await get('records', metadataId);
      const tuple = JSON.parse(new TextDecoder().decode(metadata.subarray(1)));
      if (location === 'current') tuple[1].mandatoryFutureRules = { required: true };
      else if (location === 'undo') tuple[1].snapshot.schemaVersion = 4;
      else tuple[1].snapshot.sheets[0].validation = [{ required: true }];
      const body = new TextEncoder().encode(JSON.stringify(tuple));
      const altered = new Uint8Array(body.length + 1);
      altered.set(body, 1);
      const scratchId = `${id}/unrecognized-scratch`;
      const tx = db.transaction(['records', 'sizes', 'scratch'], 'readwrite');
      tx.objectStore('records').put(altered, metadataId);
      const sizeBody = new TextEncoder().encode(JSON.stringify([metadataId, altered.byteLength]));
      const size = new Uint8Array(sizeBody.length + 1);
      size.set(sizeBody, 1);
      tx.objectStore('sizes').put(size, metadataId);
      tx.objectStore('scratch').put(new Uint8Array([0, 42]), scratchId);
      await new Promise<void>((resolve, reject) => {
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id, metadataId, scratchId, head: Array.from(head), metadata: Array.from(altered) };
    }, location);
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await expect(recovery.locator('#startup-recovery-message')).toContainText(
      location === 'current'
        ? 'Unknown stored workbook metadata'
        : location === 'undo'
          ? 'Stored workbook schema 4 is unsupported'
          : 'Invalid historical workbook metadata',
    );
    expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
      fixture.id,
    );
    const after = await recovery.evaluate(async ({ id, metadataId, scratchId }) => {
      const open = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
      });
      const tx = db.transaction(['heads', 'records', 'scratch']);
      const requests = [
        tx.objectStore('heads').get(id),
        tx.objectStore('records').get(metadataId),
        tx.objectStore('scratch').get(scratchId),
      ];
      const values = await Promise.all(
        requests.map(
          (request) =>
            new Promise<Uint8Array>((resolve, reject) => {
              request.onsuccess = () => resolve(request.result);
              request.onerror = () => reject(request.error);
            }),
        ),
      );
      db.close();
      return values.map((value) => Array.from(value));
    }, fixture);
    expect(after).toEqual([fixture.head, fixture.metadata, [0, 42]]);
    await recovery.close();
  });
}
