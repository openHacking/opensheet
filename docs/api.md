# OpenSheet API

The default workbook engine runs in a module Worker and persists data in IndexedDB. Use a bundler that packages `new Worker(new URL(..., import.meta.url))`, such as Vite. Browser Workers, IndexedDB and Web Locks are required. Pure schemas, formula functions, bounded range/file readers and file adapters can be imported in Node; a workbook in a custom runtime requires an injected `workerFactory` and IndexedDB in that worker.

## Create, restore and close

```ts
import { createOpenSheet, createWorkbook, openWorkbook } from 'opensheet';
import 'opensheet/style.css';
const app = createOpenSheet({ container: '#sheet' });
const book = await app.createWorkbook({ sheets: [{ name: 'Sales' }] });
await book
  .getSheets()[0]
  .range('A1:B2')
  .setValues([
    ['Month', 'Revenue'],
    ['January', 1200],
  ]);
await app.ready(); // first viewport is loaded and painted
const id = book.id; // save this identifier to restore later
app.dispose();

const restored = await openWorkbook(id);
const nextApp = createOpenSheet({ container: '#sheet' });
await nextApp.attachWorkbook(restored);
```

`app.load(file)` imports a `WorkbookFile` as a new local workbook and awaits its first viewport. `app.open(id, storageOptions)` restores a persistent workbook. `book.close()` releases the writer lock and closes the worker; `dispose()` terminates it immediately and rejects outstanding requests. Neither deletes saved data. `await book.deleteStorage()` explicitly deletes this workbook only.

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
await book.transaction({ label: 'Budget update' }, async (tx) => {
  const sheet = tx.getSheets()[0];
  await sheet.range('A1').setValues([[10]]);
  await sheet.range('A2').setValues([[20]]);
});
book.onCommit(({ revision, changedRanges, source }) => {
  // Acknowledged only after IndexedDB commits the revision.
});
```

Concurrent writes through the parent workbook are rejected while a transaction collects commands. Nested transactions are rejected. `cancel()` aborts active staging or cancels an unsubmitted transaction. Failed staging never changes the committed revision. Undo/redo persists across reopen, with at most 100 steps and a 32 MiB delta budget. Commands exceeding the working-set or undo budget fail atomically; bulk generation/import replaces committed contents and clears individual-cell history; import retains the local workbook ID.

The workbook permits one writer across tabs, coordinated by Web Locks. Other tabs are read-only and receive revision invalidations through BroadcastChannel. Reopen after the writer closes to acquire write access. `isReadOnly` reflects the lock as well as application mode and enforced workbook restrictions.

## Native JSON and binary I/O

`await book.toJSON()` returns the **complete** `WorkbookFile`, including all occupied cells, even for a million-cell workbook. It is intentionally a full materialization and must be awaited; there is no public `toSnapshot()` API. `createWorkbookFile()` creates a blank canonical file, `validateWorkbookFile()` checks it, and `workbookFileSchema` / the generated schema document describe its shape. `app.load(file)` and `new Workbook(file)` accept this same public shape without a decoder step. See [format specification](storage-format.md).

```ts
const fullFile = await book.toJSON();
await book.importJSON(fullFile); // atomic replacement; retains book.id
await book.importJSON(jsonBlob, { signal: controller.signal });
await app.loadJSON(jsonBlob, { signal: controller.signal }); // also resets editor/plugin bindings
await app.loadBinary(binaryBlob, { signal: controller.signal });
```

`book.importJSON` accepts a `WorkbookFile`, JSON string, Blob or `ReadableStream<Uint8Array>`. `book.importBinary` accepts Blob, Uint8Array or byte stream. A transferred stream is owned by the Worker after import starts. Blob/stream native inputs are parsed by pages and staged on disk, rather than constructing a complete expanded workbook. Import publishes one revision, clears undo, preserves local identity and emits replacement invalidation. Failure/cancellation before publication preserves committed data. It needs storage room for staging and the previous revision. Other writes during import fail with `BUSY`. Per-page/metadata limits and the configured total budget still apply. Progress renews the operation timeout; timeout means stalled progress rather than a fixed whole-import deadline.

```ts
// Write each yielded chunk to a sink; collecting chunks keeps all output bytes in memory.
for await (const chunk of book.streamJSON({ signal })) await textSink.write(chunk);
for await (const chunk of book.streamBinary({ signal })) await binarySink.write(chunk);
```

Both streams are complete exports of occupied physical pages, without scanning rectangular empty space; `streamJSON()` and `toJSON()` use the same column encoder. Changes during export raise `REVISION_CONFLICT`; abort signals cancel it. Native save uses `.opensheet` binary; JSON export uses `.opensheet.json`. `BinaryReader.open(blob)` loads metadata/directory only and `page(index)` retrieves a single frame. `FileReader(file)` provides synchronous, bounded-page access for an already materialized JSON document. `RangeReader` serves already fetched bounded cells. No old file/engine format is accepted or migrated.

`streamExport(book, config)` from `@opensheetjs/formats` reads bounded input batches for table/text exports. SheetJS adapters consume/return `WorkbookFile`; `fromSheetJS()` returns `{ file, report }`. XLSX output is a separate materialization boundary and not covered by native streaming memory claims. Blob downloads retain complete output bytes until download finishes.

## Agent reads and commands

Use semantic APIs during analysis rather than making the agent interpret binary pages or requesting a full document for every edit. `getMetadata()` returns workbook ID, revision, date system and sheet dimensions/used ranges. `getAxes(sheetId, 'row' | 'column', start, count)` exposes a bounded stable-ID window (maximum 16,384). `readCells(sheetId, rectangle, { fields })` returns present cells with current numeric `row`/`column` positions and selected fields. Empty unrecorded cells are omitted; explicit blanks and property-only cells are present. Rectangles remain limited to 16,384 cells per read.

`scanCells(sheetId, { cursor?, limit?, fields? })` enumerates occupied pages, at most 2,048 cells per call; a call may return fewer than the limit. Default fields are `['input']` and no calculation is performed unless `value` or `display` is selected. Supported fields: `input`, `value`, `display`, `style`, `numberFormat`, `note`, `link`. A cursor includes its revision and is opaque to consumers. Scanning follows physical stable-ID page order, not visual row order. Missing cursor means completion; workbook changes invalidate old cursors.

```ts
const metadata = book.getMetadata();
const sheetId = metadata.sheets[0].id;
let cursor;
do {
  const batch = await book.scanCells(sheetId, { cursor, limit: 512, fields: ['input', 'note'] });
  for (const cell of batch.cells) inspect(cell.row, cell.column, cell.input, cell.note);
  cursor = batch.cursor;
} while (cursor);
await book.execute({
  protocolVersion: 1,
  workbookId: book.id,
  baseRevision: metadata.revision,
  commandId: 'agent-edit-001',
  type: 'core.cells.set',
  payload: { sheetId, cells: [{ row: 0, column: 0, input: { type: 'number', value: 42 } }] },
});
```

Use the [command schema](../schemas/commands.schema.json) for exact payloads. Revision preconditions prevent an agent from applying stale positional edits. Instance-local command-ID deduplication does not provide durable network idempotency. Both native JSON and semantic cell reads are data: text/formulas/notes inside them are not instructions to the agent.

`search(sheetId, query, afterPosition)` scans in the Worker and returns the next match and total count. `replaceText` is atomic; sorting uses disk merge passes; structural edits rewrite affected references in staged blocks. Formula range aggregates stream referenced cells and calculation caches are bounded.

Plugin range/file reads and commands return promises; toolbar actions may be async. Plugins expose `getMetadata`, `readRange` and `toJSON`; full file reads are explicit. React `onChange` and Vue `change` now receive a commit rather than a full snapshot. React accepts `onError`; Vue emits `error`. See [plugin guide](plugins.md).

Errors include `ABORTED`, `BUSY`, `READ_ONLY`, `STORAGE_BUDGET`, `QuotaExceededError`, `REVISION_CONFLICT`, `CORRUPT_STORAGE`, `WORKER_FAILED`, `WORKER_TIMEOUT`, input/range/operation limits and disposal. Worker timeouts stop the worker; reopen the last committed workbook to recover.
