# OpenSheet API

The default workbook engine runs in a module Worker and persists data in IndexedDB. Use a bundler that packages `new Worker(new URL(..., import.meta.url))`, such as Vite. Browser Workers, IndexedDB and Web Locks are required. Pure schemas, formula functions, snapshot readers and file adapters can be imported in Node; a workbook in a custom runtime requires an injected `workerFactory` and IndexedDB in that worker.

## Create, restore and close

```ts
import { createOpenSheet, createWorkbook, openWorkbook } from 'opensheet';
import 'opensheet/style.css';
const app = createOpenSheet({ container: '#sheet' });
const book = await app.createWorkbook({ sheets: [{ name: 'Sales' }] });
await book.getSheets()[0].range('A1:B2').setValues([['Month', 'Revenue'], ['January', 1200]]);
await app.ready(); // first viewport is loaded and painted
const id = book.id; // save this identifier to restore later
app.dispose();

const restored = await openWorkbook(id);
const nextApp = createOpenSheet({ container: '#sheet' });
await nextApp.attachWorkbook(restored);
```

`app.load(snapshot)` imports a materialized snapshot as a new local workbook and awaits its first viewport. `app.open(id, storageOptions)` restores a persistent workbook. `book.close()` releases the writer lock and closes the worker; `dispose()` terminates it immediately and rejects outstanding requests. Neither deletes saved data. `await book.deleteStorage()` explicitly deletes this workbook only.

Storage options: `database`, `budgetBytes` (256 MiB default), `cacheBytes` (64 MiB Worker cache reservation), `viewportCacheBytes` (8 MiB), `operationTimeoutMs` (60 seconds), `workerFactory` and `onProgress`. Budget accounting includes encoded records, metadata, manifest, history and staged data. Browser physical usage and quota are separate estimates.

## Reads and writes

Metadata methods (`getSheets`, `getSheetById`, `sheetData`, `getStyle`, `usedRange`) are synchronous and expose immutable metadata. They do not enumerate stored cells. `getUsedRange()` uses persisted block extents.

```ts
const sheet = book.getSheets()[0];
const values = await sheet.range('A1:B2').getValues();
const inputs = await sheet.range('A1:B2').getInputs();
await sheet.range('C2').setFormulas([['B2*2']]);
await sheet.range('A1:C1').setStyle({ bold: true });
await book.undo();
await book.redo();
```

`getCell`, `value`, `display`, range reads and mutations are asynchronous. A single read batch is limited to 16,384 cells. Stream larger ranges:

```ts
for await (const batch of sheet.getUsedRange().stream({ signal: controller.signal })) {
  // batch.cells, batch.calculated, batch.display, batch.range and batch.revision
}
```

Streaming reads reject with `REVISION_CONFLICT` if the workbook changes during the stream. `values: false` skips calculation. Renderer-only `peekCell`, `peekValue`, `peekDisplay` and `isLoaded` consult cached data; unloaded values return `#LOADING` and are never treated as blank. `prefetch` loads complete 64 × 32 blocks; `onData` schedules repainting.

## Atomic transactions and events

Transactions use the scoped workbook passed to the callback. Reads inside the callback see committed data. Commands are collected, validated and atomically published after the callback succeeds. The scope expires when the callback finishes; do not retain its sheet handles.

```ts
await book.transaction({ label: 'Budget update' }, async tx => {
  const sheet = tx.getSheets()[0];
  await sheet.range('A1').setValues([[10]]);
  await sheet.range('A2').setValues([[20]]);
});
book.onCommit(({ revision, changedRanges, source }) => {
  // Acknowledged only after IndexedDB commits the revision.
});
```

Concurrent writes through the parent workbook are rejected while a transaction collects commands. Nested transactions are rejected. `cancel()` aborts active staging or cancels an unsubmitted transaction. Failed staging never changes the committed revision. Undo/redo persists across reopen, with at most 100 steps and a 32 MiB delta budget. Commands exceeding the working-set or undo budget fail atomically; bulk generation/import creates a new workbook without individual-cell history.

The workbook permits one writer across tabs, coordinated by Web Locks. Other tabs are read-only and receive revision invalidations through BroadcastChannel. Reopen after the writer closes to acquire write access. `isReadOnly` reflects the lock as well as application mode and enforced workbook restrictions.

## Export and integrations

`await book.toJSON()` explicitly materializes a full snapshot and can require substantial memory; pass a selection to materialize a bounded preview. Normal rendering, events and framework callbacks do not invoke it. Use `streamJSON()` or `streamExport(book, config)` from `@opensheetjs/formats` for bounded input batches and feed the yielded strings to a file/Blob sink. A Blob download still retains its output bytes until download completes.

`search(sheetId, query, afterPosition)` scans in the Worker and returns the next match and total count. `replaceText` is atomic; sorting uses disk merge passes; structural edits rewrite affected references in staged blocks. Formula range aggregates stream referenced cells and calculation caches are bounded.

Plugin range/snapshot reads and commands return promises; toolbar actions may be async. React `onChange` and Vue `change` now receive a commit rather than a full snapshot. React accepts `onError`; Vue emits `error`. See [plugin guide](plugins.md).

Errors include `ABORTED`, `BUSY`, `READ_ONLY`, `STORAGE_BUDGET`, `QuotaExceededError`, `REVISION_CONFLICT`, `CORRUPT_STORAGE`, `WORKER_FAILED`, `WORKER_TIMEOUT`, input/range/operation limits and disposal. Worker timeouts stop the worker; reopen the last committed workbook to recover.
