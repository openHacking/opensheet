# OpenSheet API

The default workbook engine runs in a module Worker and persists data in IndexedDB. Use a bundler that packages `new Worker(new URL(..., import.meta.url))`, such as Vite. Browser Workers, IndexedDB, Web Locks and BroadcastChannel are required. Synchronous denial of the default Worker constructor reports `WORKER_UNAVAILABLE`; an asynchronous startup failure reports `WORKER_FAILED` and suggests checking the Worker asset path and site security policy. These are possible causes, not a diagnosis of the failure. Opaque-origin browser frames report `STORAGE_UNAVAILABLE` before allocating a Worker; open the app on HTTPS or localhost or use a trusted same-origin frame. If BroadcastChannel construction is denied, `BROADCAST_CHANNEL_UNAVAILABLE` is raised before the Worker is created. Later startup failures may still have created storage, so retry the same workbook ID rather than assuming nothing was written. Pure schemas, formula functions, bounded range/file readers and file adapters can be imported in Node; a workbook in a custom runtime requires compatible Worker, storage and coordination facilities. The [Node guide](node-usage.md) includes a runnable JSON reader and distinguishes it from the full editable engine.

`Workbook.persistenceState` reports `saving`, `saved`, `failed` or `read-only`, a pending-operation count, and optional error details. `onPersistence(listener)` subscribes to subsequent changes and returns an unsubscribe function; read the property for the initial state. The editor forwards changes as `workbook:persistence`. `saved` means no tracked write is pending and the last worker acknowledgement was received and adopted. Commit observers run after adoption while the acknowledged write is still pending; saving-state observers can wait for `saved` before starting a selection-sensitive follow-up. Browser storage can still be evicted and is not a backup. Export a native or JSON copy separately.

If a revision notification throws after a Worker acknowledges a commit, that operation still succeeds: the revision was saved. The instance enters a non-acknowledgeable `SYNC_UNAVAILABLE` failure and terminates its Worker; reopen the workbook and check any other pending operations. Other tabs may show stale data until they refresh or acquire the writer lock. The failure does not guarantee that a notification already delivered to another tab can be recalled.

The editor retains in-cell and formula-bar drafts whose asynchronous writes fail in separate copyable unconfirmed-draft regions. The failure may mean the commit succeeded but its acknowledgement was lost; verify the reopened workbook before applying a draft again. Dismiss retained drafts explicitly before replacing the workbook. Pending and retained drafts share a limit of 32 per editing surface; recovery drafts are in memory and do not survive reload or disposal.

Recognizable unsupported native JSON schema or binary container versions report `UNSUPPORTED_VERSION`; malformed data remains invalid. No automatic migration or downgrade is provided. See [the version policy](adr/0001-native-format-versions.md).

A failed or unconfirmed write stays visible even if later writes succeed. `acknowledgePersistenceFailure()` clears an acknowledged nonfatal failure indicator; it does not retry the failed operation or prove what happened after a lost acknowledgement. Fatal Worker failure remains failed and requires reopening the last committed workbook. Replacement/open/attachment and native import reject with `BUSY` while tracked writes are pending. The playground also blocks demo switching and backup downloads while saving. On startup recovery, its separate Verify backup action checks native JSON or binary structure without replacing the saved workbook and offers cancellation; it does not establish intended completeness or authenticity.

Read-only opening and reload compare the persisted head before and after loading. If another writer publishes during the read, they retry; sustained changes eventually return `REVISION_CONFLICT` and the caller can retry. After subscribing to cross-tab notifications, initialization checks the head again to cover commits during subscription setup. A visible reader also checks on window focus and visibility changes; overlapping checks are coalesced, and an unchanged head skips reloading. A missing head discovered without a deletion notification is reported as `CORRUPT_STORAGE`, because its cause is unknown. This is not a multi-writer or offline collaboration protocol.

For Vite dependency optimization, nested deployment paths, strict CSP and React/Vue server-render boundaries, use the [installed-package host recipe](../examples/framework-host/README.md). Its CSP configuration initializes Zod's jitless mode before schemas are constructed; a default install may otherwise report a blocked dynamic-code capability probe. The recipe documents the global configuration and performance tradeoffs. Server-rendered wrappers produce an empty host; Worker-backed editing starts on the client.

With esbuild, bundle the application and the installed Worker separately, placing the Worker beside the emitted application module. For example, with esbuild installed and an entry named `main.js`:

```sh
esbuild main.js --bundle --format=esm --platform=browser --target=es2022 --outfile=dist/main.js
esbuild node_modules/@opensheetjs/core/dist/engine.worker.js --bundle --format=esm --platform=browser --target=es2022 --outfile=dist/engine.worker.js
```

Serve both files from the same directory, including under a deployment subpath. The installed Worker contains bare package imports, so copying it without bundling is insufficient for this setup. This recipe assumes a single application output module; splitting, hashed Worker names and other output layouts require the host to arrange a matching Worker URL, for example using `workerFactory`.

A local [installed-package test](../benchmarks/esbuild-editor-keyboard-csp-retry-2026-11-12.json) additionally uses the [Zod jitless shim](../examples/framework-host/csp-zod.js) as an esbuild alias for `zod` in both bundles and sends a CSP without `unsafe-eval` or `unsafe-inline` on the page and Worker responses. It verifies editor CSS, create, calculation, automated keyboard editing, disposal and reopen under a subpath in three browsers. A negative control confirms dynamic code is blocked in the page and Worker in all three; the Worker violation event was observed in Chromium and Firefox, but not WebKit within one second. The test injects telemetry and does not establish complete CSP reporting, native input, SSR framework or other bundler support. The commands above alone do not configure the shim or CSP.

`formulaHelp(source, cursor?)` is a pure, non-evaluating helper exported from formula, core and opensheet. Pass an incomplete formula including `=` and a UTF-16 cursor offset (default: end). It returns an optional completion replacement range with at most eight function definitions, and the innermost recognized call with a zero-based argument index. Metadata comes from the evaluator registry. Quoted text/sheet names and nested parentheses are scanned without reading any workbook. Invalid offsets or input longer than 32,767 UTF-16 units return no help; this is editing assistance, not syntax validation.

`normalizeFormulaPrecedence(expression, source)` is exported from formula and core for formula interchange. `source` is `native` or `spreadsheet`; on supported input it returns a parenthesized expression that preserves the source syntax tree under the other profile. It returns `undefined` for unsupported syntax or functions, error literals, oversized results or a failed syntax-tree round trip. It does not evaluate the formula or establish complete Excel compatibility.

`formulaReferenceHints(source)` extracts up to 64 A1 references or ranges from an incomplete formula, including absolute coordinates and quoted sheet names. It skips quoted strings and invalid/out-of-bounds addresses, returns zero-based exclusive bounds, and never evaluates the formula. Input over 32,767 UTF-16 units returns an empty array. It is a bounded editing hint, not a syntax validator.

`formulaErrorExplanation(code)` returns general guidance for supported calculation error codes, or `undefined` for unknown codes. It does not identify the actual source of an error: the error can propagate from a referenced cell. The formula bar offers the same guidance for a selected formula with a calculated error; changing the draft, selection or calculated result dismisses stale guidance.

The formula bar exposes this helper as an editable combobox: Up/Down chooses a function, Enter/Tab accepts it without committing, and Enter with no completion list commits the draft. Escape first dismisses a completion list, then cancels editing. Argument hints include the supported signature and description. IME composition suppresses completion acceptance/submission. While a formula-bar draft expects an operand, click or drag cells to insert a reference; the grid shows the selected range and the draft keeps focus. Clicking a sheet tab first allows a cross-sheet reference, with sheet names quoted and escaped. Alternatively, Alt+Down starts keyboard reference selection: arrows move, Shift+arrows extend, Page Up/Down changes sheets, Enter accepts the reference into the draft, and Escape cancels the selection. A subsequent Enter commits the finished formula. Up to 64 references on the visible sheet have colored outlines; outlines are removed when editing ends. The original edit target is restored on commit or cancel, even if the grid selection changes. Changing to read-only cancels the draft. Workbook replacement is rejected while an uncommitted draft is active.

The multiline in-cell editor also offers function completion, argument hints and reference outlines. It retains textbox semantics and uses the same accept/cancel behavior; Alt+Enter inserts a newline. Click or drag on the same sheet at an operand position to insert a reference. Selecting another sheet tab or starting Alt+Down keyboard picking moves the uncommitted draft to the formula bar so cross-sheet references retain the original target. Synthetic composition events are protected from grid and sheet switching; native input and assistive-technology acceptance are still pending.

In either editor, F4 cycles the reference endpoint at the caret through `A1`, `$A$1`, `A$1`, `$A1` and back. Selecting both endpoints of a range cycles each endpoint; sheet qualifiers and quoted text remain unchanged. This changes only the draft and is suppressed during composition. It scans the draft up to the selected reference without retaining preceding references; highlighting remains capped at 64 references. Drafts are limited to 32,767 UTF-16 units, and a change that would exceed that limit is ignored. Native keyboard and assistive-technology acceptance remain open.

In the grid, Shift+arrow extends or shrinks the selection from its keyboard anchor, including when the moving edge crosses that anchor. Ctrl/⌘+arrow moves to a visible data-region boundary; adding Shift extends the selection. Rapid data-region jumps run in order, with at most 64 waiting; excess requests report `BUSY` without changing the existing queue. Selection changes, focus leaving the grid, editing or a navigation failure discard pending jumps. Page Up/Down moves approximately one scrollable viewport; adding Shift extends the selection. Shift+Home extends to the first visible column; Ctrl+Shift+Home or ⌘+Shift+Home extends to the first visible row and column. Ctrl+End or ⌘+End moves to the bottom-right corner of the current used range; adding Shift extends from the keyboard anchor. Hidden rows and columns are skipped during directional movement, including at worksheet edges, and merged cells remain whole. If an entire axis is hidden, movement uses the nominal target because no visible target exists. A new selection or workbook commit starts a new keyboard extension sequence.

Ctrl/⌘+arrow data-region jumps are not implemented. Modified navigation keys without a supported grid action are left to the browser or host and do not move the selection as plain arrow keys. Browser and operating-system shortcuts differ; synthetic tests do not establish native keyboard or assistive-technology behavior.

The cell address box accepts A1 cells and ranges on the current sheet, or a sheet-qualified address such as `'O''Brien Q1'!B2:D4`. Enter validates the target sheet and bounds before navigating; Escape restores the current selection. These are addresses, not named ranges. Native input-method and assistive-technology acceptance remain pending.

## Workbook controls

The editor toolbar provides Sort, Filter, Note & link, and Freeze. Right-click a cell, row number, or column letter for actions on that target. Shift+F10 and the Menu key open the same grid menus. Pass `contextMenu: false` to `createOpenSheet` to leave grid right-click and context-menu keys available to the host. The toolbar remains independently configurable through `toolbar`.

Paste special accepts values, formula inputs, formats, or all available content. When a browser restricts clipboard reads, the editor opens a paste dialog with the user's chosen mode and accepts Ctrl/Cmd+V there.

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

`app.load(file)` imports a `WorkbookFile` as a new local workbook and awaits its first viewport. `app.open(id, storageOptions)` restores a persistent workbook. `book.close()` releases the writer lock and closes the worker; `dispose()` terminates it immediately and rejects outstanding requests. Neither deletes saved data. `await book.deleteStorage()` atomically deletes this workbook and disposes the instance on success; new Worker operations are rejected while deletion is pending. If the deletion transaction aborts, the instance remains open; a Worker failure requires reopening. Other open readers check the head after a deletion notification, clear their cached cells and report `WORKBOOK_DELETED` when the head is absent. Delivery depends on BroadcastChannel; a suspended or disconnected reader may temporarily retain stale display data.

`book.onDispose(listener)` registers a synchronous notification and returns an unsubscribe function. It fires once when the instance is disposed, including after successful `close()` or `deleteStorage()`; it does not fire for a rejected deletion. The instance is already disposed when listeners run, and registering on a disposed instance throws. An attached grid releases its resources at that point; load or attach another workbook to restore the editor view.

Repeated `app.attachWorkbook(book)` replaces the editor's previous subscriptions, including when attaching the same workbook. If the editor is disposed while waiting for an external workbook's readiness, attachment rejects with `DISPOSED` before creating a grid; the unattached external workbook remains usable by its owner.

If an attachment target is closed, deleted or disposed while its readiness is pending, attachment rejects with `DISPOSED` before replacing the current workbook. The current workbook remains attached and usable.

Storage options: `database`, `budgetBytes` (256 MiB default), `cacheBytes` (64 MiB Worker cache reservation), `viewportCacheBytes` (8 MiB), `operationTimeoutMs` (60 seconds), `workerFactory` and `onProgress`. A formula-bearing range read may additionally use an estimated 16–64 KiB of temporary scalar scratch outside the Worker cache reservation; even `cacheBytes: 0` permits the 16 KiB read-scoped scratch, which is discarded when the read completes. These estimates do not bound total heap. `preserveOrphans` defaults to `true`: creation refuses to remove records, size entries or scratch data when the workbook head is missing. Set it to `false` only to explicitly discard remnants of an interrupted creation. Budget accounting includes encoded records, metadata, manifest, history and staged data. Browser physical usage and quota are separate estimates.

`onOperationTiming` is an optional diagnostic callback. When supplied, completed Worker requests report time in the Worker queue, Worker execution time through response preparation, round-trip time measured on the caller, the active Worker task on arrival and the number of tasks already waiting. These are observations, not hard latency guarantees; caller-to-Worker message delivery before arrival is outside the queue measurement. Callback exceptions are isolated from the operation. Timings are not enabled when the callback is omitted.

## Reads and writes

Metadata methods (`getSheets`, `getSheetById`, `sheetData`, `getStyle`, `usedRange`) are synchronous and do not enumerate stored cells. Shared sheet metadata, used-range bounds and filtered-row lists are frozen; `getStyle` returns a detached object. An empty filtered-row list may be a detached empty array. `usedRange()` returns `undefined` when no committed range exists, including unknown or newly staged sheet IDs; `getUsedRange()` uses persisted block extents.

```ts
const sheet = book.getSheets()[0];
const values = await sheet.range('A1:B2').getValues();
const inputs = await sheet.range('A1:B2').getInputs();
await sheet.range('C2').setFormulas([['B2*2']]);
await sheet.range('A1:C1').setStyle({ bold: true });
await book.undo();
await book.redo();
```

`setValues` and `setFormulas` require a dense rectangular array matching the target dimensions; sparse rows or cells are rejected before submission. Use explicit `null` in `setValues` to clear a cell.

Range matrix and single-cell reads capture their bounds or position and row/column identities when invoked. If the Worker returns a different revision, the read rejects with `REVISION_CONFLICT`; a result already obtained for the captured revision can still resolve after a later edit. Callers that require the latest revision should compare or retry explicitly. Streamed reads capture their bounds, value option and abort signal when iteration starts and reject if the revision changes between batches.

`getCell`, `value`, `display`, range reads and mutations are asynchronous. A single read batch is limited to 16,384 cells. Stream larger ranges:

```ts
for await (const batch of sheet.getUsedRange().stream({ signal: controller.signal })) {
  // batch.cells, batch.calculated, batch.display, batch.range and batch.revision
}
```

Streaming reads reject with `REVISION_CONFLICT` if the workbook changes during the stream. `values: false` skips calculation. Renderer-only `peekCell`, `peekValue`, `peekDisplay` and `isLoaded` consult cached data; unloaded values return `#LOADING` and are never treated as blank. Cached cell and error objects returned by peeks are immutable. `prefetch` loads complete 64 × 32 blocks; `onData` schedules repainting.

## Atomic transactions and events

Transactions use the scoped workbook passed to the callback. Cell reads inside the callback see committed data. Commands are collected, validated and atomically published after the callback succeeds. A sheet added through the scope is visible in that scope's sheet list and its handle can enqueue writes; reading its cells fails until the transaction commits because it has no committed revision yet. Sheet lists, name lookups and handle names reflect staged additions, renames and removals, while the parent workbook remains unchanged until commit. Sheet handles use staged row and column dimensions and ID allocation for constructing subsequent commands. For existing sheets, `sheetData()` and cell reads still describe the committed revision; for a new sheet, `sheetData()` exposes its initial metadata. If a removed sheet ID is re-added within the transaction, reads by that ID still see the old committed sheet until commit. Other staged metadata is not reflected until commit. A range in newly inserted space can be used for a write, but a committed read of that range fails if it is outside the committed sheet. `copyTo` commands, including selective `values`, `formulas` and `formats` modes, read the staged source when executed in a transaction. Selective copies are limited to 200,000 cells and a 32 MiB serialized working-set estimate. The scope expires when the transaction finishes; do not retain its sheet handles.

The public `addSheet` helper creates an empty sheet. A validated `core.sheet.add` command may also carry initial cells keyed by their physical row and column IDs; they are staged and committed with the sheet, including in transactions. Initial cells are limited by both the 32 MiB traversal estimate (512 bytes per cell) and a separate 32 MiB working estimate (four times the UTF-8 serialized cell sizes), in addition to storage and undo budgets. These estimates are not process-memory limits. Invalid identities, missing styles or exceeded budgets reject the operation; existing data is not partially published.

Subscriptions registered through the scoped workbook are removed when its transaction finishes, including on failure. Register on the parent workbook for notifications that must outlive the transaction.

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

Concurrent writes through the parent workbook are rejected while a transaction collects commands. Nested transactions are rejected, as are `close()` and `deleteStorage()` through either the parent or its active transaction scope. `cancel()` aborts active staging or cancels an unsubmitted transaction. Failed staging never changes the committed revision. Undo/redo persists across reopen, with at most 100 steps and a 32 MiB delta budget. Commands exceeding the working-set or undo budget fail atomically; bulk generation/import replaces committed contents and clears individual-cell history; import retains the local workbook ID.

Transactions require pending writes to finish before they start. Bulk generation is exclusive: it cannot start inside a transaction or while a write is pending, and edits, imports, history operations and deletion return `BUSY` until it completes.

The workbook permits one writer across tabs, coordinated by Web Locks. Other tabs are read-only and receive revision invalidations through BroadcastChannel. A read-only instance can call `tryAcquireWriteLock()` to attempt a nonblocking handover; it returns `false` while another writer holds the lock. On success it reloads the latest committed head and recovers storage before enabling writes. `canRequestWriteLock` distinguishes lock contention from an explicit read-only mode. `isReadOnly` reflects the lock as well as application mode and enforced workbook restrictions.

## Native JSON and binary I/O

`await book.toJSON()` returns the **complete** `WorkbookFile`, including all occupied cells, even for a million-cell workbook. The returned file has independent, editable metadata; changing it does not alter the open workbook until it is imported. It is intentionally a full materialization and must be awaited; there is no public `toSnapshot()` API. `createWorkbookFile()` creates a blank canonical file, `validateWorkbookFile()` checks it, and `workbookFileSchema` / the generated schema document describe its shape. `app.load(file)` and `new Workbook(file)` accept this same public shape without a decoder step. See [format specification](storage-format.md).

```ts
const fullFile = await book.toJSON();
await book.importJSON(fullFile); // atomic replacement; retains book.id
await book.importJSON(jsonBlob, { signal: controller.signal });
await app.loadJSON(jsonBlob, { signal: controller.signal }); // also resets editor/plugin bindings
await app.loadBinary(binaryBlob, { signal: controller.signal });
```

`book.importJSON` accepts a `WorkbookFile`, JSON string, Blob or `ReadableStream<Uint8Array>`. `book.importBinary` accepts Blob, Uint8Array or byte stream. A transferred stream is owned by the Worker after import starts. After an early failure the reader requests cancellation and releases its lock without waiting for the source cancellation promise, which can remain pending. Blob/stream native inputs are parsed by pages and staged on disk, rather than constructing a complete expanded workbook. Import publishes one revision, clears undo, preserves local identity and emits replacement invalidation. Failure/cancellation before publication preserves committed data. It needs storage room for staging and the previous revision. Other writes during import fail with `BUSY`. Per-page/metadata limits and the configured total budget still apply. Progress renews the operation timeout; timeout means stalled progress rather than a fixed whole-import deadline.

```ts
// Write each yielded chunk to a sink; collecting chunks keeps all output bytes in memory.
for await (const chunk of book.streamJSON({ signal })) await textSink.write(chunk);
for await (const chunk of book.streamBinary({ signal })) await binarySink.write(chunk);
```

Both streams are complete exports of occupied physical pages, without scanning rectangular empty space; `streamJSON()` and `toJSON()` use the same column encoder. Changes during export raise `REVISION_CONFLICT`; abort signals cancel it. Native save uses `.opensheet` binary; JSON export uses `.opensheet.json`. `BinaryReader.open(blob)` loads metadata/directory only and `page(index)` retrieves a single frame; encoded frames above 32 MiB are rejected before page materialization. `FileReader(file)` provides synchronous, bounded-page access for an already materialized JSON document. `RangeReader` serves already fetched bounded cells. Pre-0.1.1 persisted-engine layouts and native schema versions other than 3 are unsupported; an unversioned directory head in the current layout is accepted and marked on its next successful publication. There is no general migration or downgrade guarantee.

For a large rectangular read, `reader.getSheetById(id).range('A1:C50000').batches({ maxCells: 4096 })` returns a synchronous iterator of `{ bounds, values }` in row order. Each matrix contains at most `maxCells` cells; the option defaults to 16,384 and must be between the range width and 16,384. Set `display: true` for formatted text. Options and bounds are captured when `batches()` is called, and iteration can stop early. A batch limit bounds the returned matrix only: the materialized file, metadata, calculation cache, individual formula work and consumer-retained batches can consume additional memory. Existing `getValues()` and `getDisplayValues()` still materialize the full requested matrix and do not impose that batch limit.

`await verifyBinary(blobOrBytes, { signal? })` validates a complete native binary backup without opening or modifying IndexedDB. It checks metadata, directory, page framing/checksums, decoded cells, style references and row/column identities, and returns `{ workbookId, sheets, pages, cells }`. Cancellation actively cancels a pending input or decompression read; synchronous work within a page may still delay it. Both verification helpers cooperatively yield between completed pages when their task budget is exceeded; an individual page or metadata parse can still delay the main thread. Verification uses bounded page buffers but retains directory and axis indexes, so it is not constant-memory. CRC detects accidental corruption, not malicious tampering; a structurally valid file can still omit data the caller intended to back up. Keep and compare a separate trusted manifest or reference when completeness/authenticity matters.

`await compareBinaryBackup(book, blob, { signal? })` first verifies a binary backup, then compares it byte for byte with a fresh streamed export of the current workbook. It returns `{ matches, revision }` and rejects with `REVISION_CONFLICT` if the workbook changes during the comparison. Cancellation stops waiting for a pending export chunk, though already-started Worker work may finish later. This is an exact-byte check, not a semantic equivalence test: different encoder or compression implementations can produce different bytes for equivalent data. It reads the backup twice, retains the verification indexes described above, and does not authenticate either the backup or the live workbook. Local 200k-cell timings and their limits are in [performance](performance.md); peak Worker/process memory remains unmeasured.

`await verifyJSON(textOrBlob, { signal? })` provides the analogous read-only check for native schema-3 JSON. It parses the replayable input twice because metadata may follow blocks, validates every block against the sheet axes and styles, and returns the same counts. Non-replayable streams are rejected; cancellation actively cancels a pending input read but processing a long chunk can still delay it. It retains metadata and page identities and does not prove intended completeness or authenticity. Local large-file measurements and their limitations are in [performance](performance.md).

`streamExport(book, config)` from `@opensheetjs/formats` reads bounded input batches for table/text exports. SheetJS adapters consume/return `WorkbookFile`; `fromSheetJS()` returns `{ file, report }`. XLSX output is a separate materialization boundary and not covered by native streaming memory claims. Blob downloads retain complete output bytes until download finishes.

Single-column blank CSV/TSV rows are quoted as `""`, allowing OpenSheet's delimited parser to preserve a final empty row. First-column text beginning with U+FEFF is also quoted so its data is not mistaken for a file BOM. Other spreadsheet applications may still trim trailing blanks on import or save.

For CSV/TSV, `streamExport(book, config, { onReport })` calls `onReport` after the consumer fully reads a successful stream at one revision. The report contains `format`, `revision`, the total `sanitized` text cells, up to 32 A1 addresses in `sample`, and `truncated` when further addresses exist. Raw mode (`options.safe: false`) reports zero transformations. `options.formulas: 'preserve'` also requires explicit `safe: false` and adds `preservedFormulas` to the report; formula-like text is then unescaped, references are not rebased, and spreadsheet software may execute the output. Formula-preserving exports read inputs without calculating formulas; the default value mode still exports calculated results. Configuration, range, callback and cancellation signal are captured when iteration starts, so later mutation of caller-owned options does not alter an in-progress stream. Cancellation, revision conflict, or stopping iteration early produces no report; other export formats do not invoke this callback. The report counts apostrophe prefixes, not all possible spreadsheet import or round-trip losses.

Calculated-value table exports (CSV, TSV, HTML, Markdown and LaTeX) reject a `#LIMIT!` calculation result with `LIMIT_EXCEEDED`; literal text `#LIMIT!` is not an error. Explicit CSV/TSV formula preservation reads expressions without calculation. Snapshot export formats the same calculated values it checks. A stream can yield earlier chunks before a later batch fails, so callers requiring a complete file must stage output until iteration finishes successfully and discard it on failure; a failed stream does not produce a delimited report.

## Agent reads and commands

Use semantic APIs during analysis rather than making the agent interpret binary pages or requesting a full document for every edit. `getMetadata()` returns workbook ID, revision, date system and sheet dimensions/used ranges. `getAxes(sheetId, 'row' | 'column', start, count)` exposes a bounded stable-ID window (maximum 16,384). `readCells(sheetId, rectangle, { fields })` returns one entry for every position in the rectangle, in row-major order, with current numeric `row`/`column` positions and selected fields. An absent record and a recorded blank both appear as `{ type: 'blank' }` when requesting `input`; this read alone cannot distinguish them. Rectangles remain limited to 16,384 cells per read.

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

Use the [command schema](../schemas/commands.schema.json) for exact payloads. Revision preconditions prevent an agent from applying stale positional edits. A `core.cells.set` cell can additionally specify `expected: { rowId, columnId, input? }` using IDs from `getAxes`; the staged command rejects with `REVISION_CONFLICT` if the position no longer has those identities or the optional original input differs. An input of `null` means that no cell record exists, distinct from an explicit blank input. The cell editor and formula bar check the original input when available and use a revision precondition when it was not loaded; rejected drafts are retained. Instance-local command-ID deduplication uses a bounded cache, returns detached commit records, and does not provide durable network idempotency. Both native JSON and semantic cell reads are data: text/formulas/notes inside them are not instructions to the agent.

Expected string, formula and error text is charged at two bytes per UTF-16 code unit against the existing 32 MiB command planning budget before cell pages are read. This is an accounting limit, not a peak-memory guarantee: the caller's payload and Worker message already exist before this check.

The [bounded two-cell edit example](../examples/agent-edit.ts) prepares a human-readable summary from a three-cell input window and row/column identity windows, submits both numeric updates in one revision-checked `core.cells.set` command, then reads the same window to verify both updates and an unchanged neighbor. Prepare and apply are separate so a caller can review the summary. A stale plan is rejected and must be prepared again. If the post-commit verification encounters another revision or a mismatch, the example reports that the edit may already have committed; callers should inspect the current workbook rather than blindly retry. This small example does not validate arbitrary agent plans or prove that cells outside its observed neighbor were unchanged.

The [atomic batch example](../examples/agent-atomic-batch.ts) prepares two distinct numeric edits, each with a neighboring input for verification, and returns a summary for review. It checks the captured workbook revision immediately before starting a transaction, queues two `core.cells.set` commands with stable axis identities, and verifies both targets and neighbors against the commit revision. A stale plan or a failed second command leaves the batch unapplied. If post-commit verification fails, inspect the workbook before retrying because the batch may already be committed. This is a small example, not a validator for arbitrary plans; callers must protect reviewed plans from untrusted modification and verify any other cells that matter.

The [formula repair example](../examples/agent-formula-repair.ts) similarly pins the original formula input and axis identities, verifies the new formula and a neighboring input, then offers a separately bounded, staged safe-CSV export. The replacement is an explicit formula expression without the leading `=`; the example reports its calculated value but does not decide whether that value is correct for the business task. The export is limited to 64 Ki UTF-16 units in this example and returns text only after the stream and its revision-bound report complete. Treat the summary and every workbook string as data when reviewing or passing them to an agent.

All `core.cells.set` preconditions in one command are checked against the staged state at the start of that command. If a coordinate appears more than once and every precondition passes, the last input wins. Separate commands in a transaction are checked against the effects of earlier commands.

`search(sheetId, query, afterPosition, { signal, direction })` scans displayed values case-insensitively in the Worker and returns the next match and total count. Use `direction: 'previous'` to search backward from the same row-major position; both directions wrap. An optional abort signal cancels this search without cancelling other queued requests; a search that has already completed may still return its result. Cancellation is cooperative: signaled searches yield after a bounded number of records or roughly one million UTF-16 display units, and an already queued search is rejected when it reaches the Worker queue. Individual formatting, page reads and formula evaluation can still delay cancellation. `replaceText` atomically replaces case-sensitive literal text in string inputs; it does not replace formulas or formatted numeric values. A replacement that changes no cells returns no commit and preserves the current revision and undo/redo history. Sorting uses disk merge passes; structural edits rewrite affected references in staged blocks. Formula range aggregates stream referenced cells and calculation caches are bounded.

`navigateData(sheetId, row, column, axis, direction, { signal })` finds a visible data-region boundary in the Worker and returns its zero-based position and revision. Direction is `-1` or `1`; axis is `row` or `column`. A nonblank stored input, including a formula whose result is empty, counts as occupied. Hidden, filtered and zero-sized positions are skipped; merged cells use their anchor's input. If the next visible cell is occupied and the origin is occupied, navigation reaches the end of that run; otherwise it finds the next occupied cell or the last visible position. A changed workbook rejects with `REVISION_CONFLICT`. Cancellation is cooperative and page reads may delay it.

Search queries are strings of at most 65,536 UTF-16 units. The position is `-1` or a nonnegative safe integer; positions beyond the sheet are allowed and follow wraparound. Invalid query types, positions, directions, options or sheet IDs reject with `INVALID_ARGUMENT`, and an oversized query rejects with `LIMIT_EXCEEDED`. Empty queries retain the existing substring-search behavior over recorded cells.

Plugin range/file reads and commands return promises; toolbar actions may be async. Plugins expose `getMetadata`, `readRange` and `toJSON`; full file reads are explicit. React `onChange` and Vue `change` now receive a commit rather than a full snapshot. React accepts `onError`; Vue emits `error`. See [plugin guide](plugins.md).

Errors include `ABORTED`, `BUSY`, `READ_ONLY`, `STORAGE_BUDGET`, `QuotaExceededError`, `REVISION_CONFLICT`, `CORRUPT_STORAGE`, `WORKER_FAILED`, `WORKER_TIMEOUT`, input/range/operation limits and disposal. Worker timeouts stop the worker; reopen the last committed workbook to recover.

## Formula metadata

`@opensheetjs/formula` exports immutable `functionDefinitions` and `supportedFunctions`, plus `functionDefinition(name)` for canonical uppercase lookup. The same registry validates arities in both evaluator paths and generates the [function reference](formula-functions.md). See [semantic limits and compatibility changes](formulas.md).

### Stable multi-column sorting

```ts
await sheet.sort('A2:D600', [
  { column: 0, direction: 'asc' },
  { column: 3, direction: 'desc' },
]);
// Existing single-column calls remain supported.
await sheet.sort('A2:D600', 0, 'asc');
```

Column indexes are zero-based absolute sheet positions inside the selected range. Keys are evaluated in priority order; equal rows retain source order. Supply 1–64 distinct columns. Formula keys use the same pinned pre-sort stage as scalar keys. Headers are excluded by choosing a range below them; the editor’s Custom sort dialog provides an explicit header checkbox. Commands accept either `{ sheetId, range, keys }` or the legacy `{ sheetId, range, column, direction }`, never both. See the formula sort design for mixed-type ordering and resource limits.

### Typed multi-column filters

```ts
await sheet.setFilter({
  mode: 'all', // 'any' combines conditions with OR
  headerRows: 1,
  criteria: [
    { column: 0, operator: 'eq', value: 'active' },
    { column: 2, operator: 'gte', value: 10 },
    { column: 2, operator: 'lt', value: 100 },
  ],
});
await sheet.setFilter(null); // clear all conditions
```

Use 1–64 conditions on absolute zero-based columns. Repeated columns allow bounded intervals. Conditions apply to calculated values across the whole sheet after `headerRows`; they do not remove data or overwrite manual hidden flags. `eq`/`neq` accept text, finite numbers or booleans with strict types and case-insensitive text equality. Null and errors match neither equality nor inequality. `gt/gte/lt/lte` require numeric values; numeric text never passes these operators. Dates stored as serial numbers use the workbook's date system. Text-only `contains/notContains/startsWith/endsWith` ignore case and interpret text literally, without wildcards. `isBlank/isNotBlank` treat null and empty text as blank; `isError/isNotError` inspect ordinary formula errors. Calculation resource exhaustion aborts atomically, rather than being treated as an ordinary error. Conditions short-circuit in supplied order.

The legacy `{ column, query }` filter remains a case-insensitive search of formatted display text with one header row. Custom filters use raw calculated values instead. Both reject the staged operation if a filter value exceeds the calculation work limit. The playground provides editing, all/any matching, explicit value types, header count and clear/reset through Custom filter. Applying uses the dialog's opening revision.

Inserted columns shift condition indexes; deleting a referenced column removes its conditions, clearing the filter when none remain. Row edits within the header prefix adjust its size. Visibility is recalculated after edits, sort, undo/redo and reopen. Native JSON preserves criteria; XLSX filter export is still unsupported and reported as a compatibility loss.

`workbook.onProgress(listener)` subscribes to Worker progress and returns an unsubscribe function. Sorting reports `sort-runs`, `sort-merge` and `sort-output` with completed/total rows and `elapsedMs` for each finished phase. These phases describe staging work, not a committed revision; the command promise is still the durable completion signal. Import progress also reports `import-source-stage` (source handling and scratch staging wall time) and `import-scratch-write` (time awaiting scratch writes within that stage); neither is a durability acknowledgement. Dispose clears subscribers. Observer errors are isolated from command execution.

Native import reports cumulative `elapsedMs` summaries for `import-scratch-read`, `import-decode`, `import-pages` and `import-scratch-delete` after all staged pages are processed. `import-pages` includes page construction, encoding and storage, not just database I/O; `import-scratch-read` includes asynchronous cursor waits. These are diagnostic wall-clock totals, not CPU samples, and precede final metadata publication/cleanup. Existing `import-parse` and `import-store` counters still report completed pages.

### Cell notes and links

```ts
await sheet.range('B2').setNote('Review this assumption\nSource verified');
await sheet.range('B2').setLink({ target: 'https://example.com/source', tooltip: 'Source' });
const notes = await sheet.range('B2:B4').getNotes();
const links = await sheet.range('B2:B4').getLinks();
await sheet.range('B2').setNote(null); // leaves the link and value unchanged
```

`setNote` and `setLink` apply one value across a range and participate in commands/history. `getNotes` and `getLinks` return matrices with null for absent metadata. `core.cells.annotation` accepts `{ sheetId, range, note?, link? }`, requires at least one metadata field, preserves omitted fields and deletes explicit null fields. Supply both fields for one atomic update. Notes allow 32,768 characters; link targets allow 4,096 and descriptions 1,000. Covered merged cells reject the whole command; annotate the merge anchor instead. Clearing values preserves annotations; `clear({ all: true })` removes them. Copy, fill, sorting and structural edits move them with their cell data.

`isSafeLinkTarget` implements the shared link profile: absolute HTTP/HTTPS URLs with hosts, or nonempty mailto targets, without raw ASCII whitespace/control characters or backslashes. Encode spaces in URLs. Relative paths, protocol-relative URLs, file/data/script schemes and malformed addresses are rejected. Native metadata validation enforces the profile; XLSX import omits unsupported/malformed links and reports the loss. Links open only through an explicit user action in the note/link dialog, in a new browsing context with `noopener noreferrer`.

The playground's **Cell note and link…** action edits the first selected cell, including merged anchors. Save and removal are revision-bound atomic commands. Read-only mode permits viewing/copying notes and explicitly opening approved links, while disabling mutation. Notes are plain text. The grid paints a note marker, applies link color/underline without changing base styles, and adds note/link presence to the active cell's accessible description. This automated behavior is not a substitute for native screen-reader acceptance.

### Row and column actions

`sheet.insertRows(index, count)` and `sheet.insertColumns(index, count)` insert before the zero-based index. The playground's sheet actions insert as many rows or columns as the selection spans. `sheet.setRowsHidden(index, count, hidden)` and `sheet.setColumnsHidden(index, count, hidden)` change visibility as a single undoable command; each range must be nonempty and within the sheet. `sheet.setAllRowsHidden(hidden)` and `sheet.setAllColumnsHidden(hidden)` target the full axis at command execution time, including axes added by an earlier queued operation. These commands preserve sizing and existing axis identities, and revealing visible axes does not create metadata for them. `sheet.setRowsHeight(index, count, size)` and `sheet.setColumnsWidth(index, count, size)` set sizes from 20 to 2,000 pixels across a range in one undoable command, preserving visibility. The playground offers both adjustments for the current selection, prefilled from its first row or column. The playground can hide or show selected rows or columns and show all rows or columns. Selection-sensitive immediate actions ask the user to retry after pending writes finish, so they do not use stale coordinates; full-axis show can queue. Changes survive native save and reopen; ordinary storage and undo budgets apply. A standalone visibility command that makes no change creates no history entry.

### Frozen panes

`sheet.setFreeze(rows, columns)` freezes the specified number of leading rows and columns in one undoable command. In the playground, **Freeze rows and columns before selection** uses the selection's top-left cell as the first scrolling row and column; the first-row and first-column actions preserve the other axis, and **Unfreeze rows and columns** clears both. Frozen panes survive native save and reopen. XLSX export preserves simple panes when the generated archive passes the checks and size limit described in the [format guide](04-sheetjs-and-formats.md); otherwise it reports a loss. With raw workbook parts, XLSX import preserves simple frozen panes with an unscrolled default view and verified worksheet relationships; split panes, unsupported view states and non-default selections are reported as losses.

### Numeric fill series

`range.fillSeries({ start, step = 1, axis = 'row' })` replaces values in the selected rectangle with an arithmetic numeric series. `axis: 'row'` repeats the same vertical series in every column; `axis: 'column'` repeats the horizontal series in every row. Negative and zero steps are supported. Inputs must be finite numbers. Terms are calculated from their decimal representations rather than by repeated floating-point addition; final cell values remain JavaScript numbers. Overflow rejects the entire command.

The command preserves each target cell's style, note and link, replaces formulas, recalculates dependents, and supports undo/redo and persistence. Any merged-cell intersection is rejected. Series and style commands stage pages and each accept at most 200,000 cells; the storage and 32 MiB undo budgets still apply. The equivalent validated command is `core.cells.fillSeries` with `{ sheetId, range, start, step?, axis?, calendar? }`. `calendar: 'month'` or `'year'` interprets the integer start as a date serial in the workbook date system and requires an integer step; each result keeps the original day where possible and clamps it in shorter months. The API preserves existing formatting; callers can apply a date format in the same transaction. The playground exposes **Sheet actions → Fill series…**, with revision conflict and read-only handling. Its date mode fills integer day, month or year steps and applies `yyyy-mm-dd` in the same undoable transaction while preserving other style properties and annotations. It supports 1900 and 1904 date systems, including the 1900 system’s compatibility date 1900-02-29; ranges must remain from 1900-01-01 or 1904-01-01 respectively through 9999-12-31. Automatic pattern inference and drag handles are not implemented.

### Text clipboard safety

Grid cut captures the source range and workbook revision. Source values are cleared only after `navigator.clipboard.writeText` resolves; permission failures preserve cells and history. Plain and structured copy reject a calculation `#LIMIT!` with `LIMIT_EXCEEDED` rather than copying an incomplete result; literal text with those characters remains copyable. If any workbook edit occurs while copying, the revision conflict preserves the newer data. If the workbook or grid is read-only when the clipboard write completes, the clear is rejected and the source is retained. Disposing the grid before the clipboard write completes also prevents the clear and suppresses late error delivery to its host. Moving the selection does not change the captured source. Successful clear is one undoable command and retains cell formatting/annotations. Clipboard contents and IndexedDB cannot be committed as one transaction: a later clear failure can leave copied text on the clipboard while preserving source cells. Cut currently writes plain text; moving formulas and their dependents with structured cut semantics is not implemented.

An accepted grid paste uses the selection captured at invocation. When both selection dimensions are exact multiples of the clipboard rectangle, the clipboard is tiled across that selection; otherwise the source rectangle is pasted once at the selection’s top-left. Hidden rows and columns are included. Structured formula references relocate for each destination, and a merged-cell intersection rejects the entire paste. Tiled pastes use a compact Worker command and stage pages atomically, up to 200,000 cells and a 32 Mi UTF-16-unit logical expansion budget; transformed formula growth also counts. Sparse non-tiled pastes have a separate two-million-comparison merge-check budget. Storage and undo budgets still apply. If the user changes selection before acknowledgement, completion leaves the newer selection in place. A disposed grid does not send late selection or error callbacks; disposal does not undo a paste already committed to the workbook.

TSV parsing supports quoted tabs, CR/LF, doubled quotes and ragged rows (blank padding). A field beginning with a quote must close before a separator or the end of input; text following its closing quote is rejected rather than silently altered. Literal quotes in a quoted field must be doubled. Copy quotes carriage returns as well as tabs/newlines. Parsing checks the 2 MiB character limit, 32,768-character cells, 100,000 rows, 1,000 columns and 200,000-cell rectangular limit before matrix expansion; the core's smaller working-set limit may additionally reject a paste atomically. The grid also provides HTML table and selective-paste support described below.

### Selective copy and atomic paste

`sourceRange.copyTo(targetRange, { mode })` supports `all` (the existing default), `values`, `formulas` and `formats` within the same sheet. Source and destination must have equal dimensions and no merged-cell intersections. Hidden rows are included by physical range position.

Outside a transaction, `copyTo` in `all` mode and `fillDown` reject with `BUSY` while another write is awaiting acknowledgement; retry after the pending write finishes and refresh the selection. Inside a transaction, their commands apply in transaction order to the staged state.

- `values` copies calculated scalar results (including ordinary errors and blanks), retaining destination formatting and annotations. Incomplete `#LIMIT!` calculations abort the operation.
- `formulas` copies inputs, including literal cells and blanks, and adjusts relative/mixed formula references. Destination formatting and annotations remain intact.
- `formats` replaces destination formatting with source formatting, including number formats; an unformatted source clears destination formatting. Values, formulas, notes and links are preserved.

Selective modes capture the workbook revision and source bounds when invoked, read a bounded source snapshot and submit one revision-bound `core.cells.paste` command. Overlap therefore uses original source cells; intervening writes cause `REVISION_CONFLICT`, including writes before the source read completes. Existing read and command working-set limits apply, and one undo restores the operation.

`core.cells.paste` accepts `{ sheetId, cells: [{ row, column, input?, style?, note?, link? }] }`, with at least one of `input`, `style`, `note` or `link` per cell. Optional `note` and `link` accept null to clear the corresponding annotation. Omitted fields are preserved, `style: {}` clears formatting, and input/style changes are committed together. This command rejects any target intersecting a merge. It is the core primitive for structured clipboard support; the HTML clipboard parser and browser paste-special UI use this command.

### HTML clipboard tables

`CanvasGrid.pasteHTML(html, fallbackText?)` and the grid paste event accept one HTML table. HTML takes precedence over plain text when a table is present; non-table HTML uses the plain-text fallback. Text includes entities, tabs and line breaks. Ragged rows are blank-padded. Pasted values/formulas and supported formats commit in one command with undo/redo. Destination notes and links are preserved; absent source formatting clears destination formatting.

The parser uses an inert template that is never inserted into the page. Scripts, styles, embedded documents and templates do not contribute cell text; event handlers, links and media are not imported. Supported formatting is cell-level inline bold/italic/underline, alignment, wrapping, 8–72 px font size (pt converted), and six-/three-digit hex, integer RGB/opaque RGBA plus common named foreground/background colors. Stylesheets, inherited/nested rich-text formatting, borders, number formats, images and links are outside this initial HTML profile.

Multiple/nested tables and source row/column spans are rejected rather than silently flattened. Existing paste text/dimension/working-set limits apply. Read-only grids ignore the operation. External application-generated fixtures remain separate acceptance work.

### Paste-special interface

The playground exposes **Sheet actions → Paste special…**. Paste into the focused capture box, choose a mode, then apply. Capture and cancel do not change cells. The target and workbook revision are captured when opening the dialog; changes made while it is open cause a conflict instead of overwriting newer data. Read-only mode disables capture and apply. A bounded plain-text preview never renders clipboard markup.

The renderer exports `pasteClipboard(book, selection, { text, html? }, { mode?, baseRevision? })`, returning the pasted selection after its commit. Modes are `all` (default), `values`, `formulas` and `formats`. Values mode interprets scalar text but keeps leading-`=` text literal and preserves a leading apostrophe in external text and HTML; it cannot evaluate external formulas to recover their cached result. Formula mode copies supplied inputs verbatim, without source-relative relocation because external text/HTML carries no trusted source coordinates. Values/formula modes preserve destination formatting. Formats mode requires an HTML table, replaces supported formatting, and preserves destination inputs and annotations. Normal HTML parsing limitations apply. Internal reference-aware source copying remains available through `Range.copyTo(..., { mode })`.

Image/file-only clipboard events are ignored by the grid. The Paste special dialog reports unsupported data and discards any previously captured payload, preventing an empty or stale paste.

HTML text extraction preserves explicit BR line breaks and separates paragraph/div/list/preformatted/heading block boundaries, without adding a trailing synthetic newline. Embedded SVG, MathML, canvas, audio and video subtrees are excluded from text along with scripts/styles/frames/templates. Traversal is iterative, with no recursive call stack or sibling-sized queue. The 32,768-character text limit counts the resulting cell text; a removed trailing block separator does not consume that limit.

### Structured clipboard copies

`copyClipboard(range)` from the renderer, also exposed as `CanvasGrid.copyClipboard()`, returns `{ text, html }` from one bounded cell snapshot. Text is formatted TSV for other applications; HTML contains escaped display text/basic CSS plus a versioned OpenSheet cell payload with validated input, captured scalar value, complete supported style and source coordinates. The normal grid copy event publishes both MIME representations through promise-backed ClipboardItems when available; platforms without rich clipboard writing fall back to text. Cut currently remains the guarded plain-text path.

`CanvasGrid.selectionVersion` is a read-only, per-grid counter for guarding asynchronous selection updates. It advances when a selection is set, including setting the same coordinates again, and when a structural change transforms it. Compare both the grid identity and the counter before applying a delayed selection; it is not a workbook revision or a persistent identifier.

On OpenSheet paste, all/formula modes relocate relative and mixed references from copied coordinates. Values mode uses captured scalar values, including errors/blanks, even if the source has since changed. Formats mode preserves destination input and uses the copied supported style, including number formats. Clipboard input is untrusted: schema validation, origin bounds, rectangular shape and version checks precede the atomic command. Unknown fields in structured cells or their input/value objects, and malformed payloads, are rejected. Normal structured paste copies source notes/links, including explicit clearing when absent at the source; selective values/formula/formats modes retain destination notes/links. Merges are not copied. Older version-1 payloads with omitted annotation fields preserve the corresponding target annotation. Copy rejects intersecting merges, incomplete `#LIMIT!` calculations and HTML over the 2 MiB character limit, and inherits the bounded range-read limit.

External applications may strip custom data attributes; the remaining HTML then follows the external-table text/style profile, without source-relative formula reconstruction. Native cross-application preservation has not yet been established. The core now re-exports `offsetFormula(expression, rowDelta, columnDelta)` for reference-aware integrations without a renderer-to-formula package dependency. For repeated relocation of one expression, `@opensheetjs/formula` also exports `createFormulaOffsetter(expression)`, which parses once and returns a `(rowDelta, columnDelta) => string` function with the same reference-boundary behavior.

Structured annotations remain plain data: note text is never rendered as markup, and links use the same HTTP/HTTPS/mailto target validation as native cells. Metadata is validated before applying any cell, so an unsafe target or oversized note rejects the complete clipboard matrix. An ordinary external HTML table still does not import arbitrary links/notes from its DOM.

`isSupportedNumberFormat(code)` identifies the conservative numeric/date/time display profile. Unsupported codes can still be stored and exported, but their rendered display may be approximate. XLSX imports report those cases; the profile is not full Excel SSF compatibility.
