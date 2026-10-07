# Data, commands and consistency

> Design notes, not a claim of implemented functionality. See [implementation status](implementation-status.md) and the [current API](api.md) for supported behavior.

## Model and identity

A WorkbookFile is serializable, validated data, without DOM nodes, functions, plugin instances, credentials or network objects. Runtime indexes use typed arrays and sequence/interval trees; native JSON uses homogeneous page columns, axis runs and sparse metadata objects. Model schema, command protocol and package versions are distinct.

Rows and columns have stable sheet-local uint32 IDs and independently ordered positions. A1 addresses map to current positions; numeric ranges use zero-based, exclusive ends. Sparse storage retains nonempty cells and supported metadata. Blank, empty string, zero and false are distinct. Numbers must be finite; long identifiers and unsafe integers should be text. Spreadsheet dates preserve date-system semantics rather than silently becoming UTC timestamps.

Formula text is separate from derived calculation results. Imported cached results have provenance and expire after edits. Viewport and transient selection belong to UI session state rather than the business snapshot. Snapshot reads are defensive or frozen; internal mutation is prohibited.

## Commands and transactions

All mutations enter through validated commands. A transaction stages commands against one starting state, validates invariants, then atomically commits one revision, one history entry and one event. Asynchronous callbacks receive an isolated transaction facade; nested transactions and concurrent writes through the outer facade are rejected. Failure must preserve the previous workbook, even if an inner error is caught.

An envelope may provide workbookId, commandId, baseRevision and protocolVersion. Current retry deduplication is instance-local; durable idempotency and remote authorization are host responsibilities. Unknown command types and incompatible schemas must fail explicitly.

The intended processing boundary is syntax validation, resource budgeting, host authorization where applicable, revision checks, staged model changes, invariant validation, commit, history and notifications. Listener failures do not reverse a completed transaction.

## References and structural changes

A1 ranges are case-insensitive. Stable identities do not imply complete Excel structural semantics. Merge rejects overlapping ranges and covered nonempty values unless discarding is explicit and undoable. Edits cutting through merged ranges fail. Formula references must be transformed consistently with structural changes; unsafe rewrites must block the edit.

Sorting moves data/styles with copy-relative formulas while row IDs and heights stay fixed. Merged ranges and formula sort keys are unsupported. Filters are visibility projections, not data deletion. Deleting reference endpoints can conservatively yield #REF!.

## History, events and errors

Current history stores persistent block-level deltas and retains at most 100 transactions or an estimated 32 MiB. Oversized mutations fail atomically. New edits clear redo history. Undo/redo creates new revisions; calculation caches do not become separate business history entries.

The current event contract is documented in [API](api.md). Commit events are not complete network replication records. Standard error codes include INVALID_ARGUMENT, INVALID_RANGE, READ_ONLY, LIMIT_EXCEEDED, UNSUPPORTED_FEATURE, REVISION_CONFLICT, PLUGIN_CONFLICT and DISPOSED. Dispose is idempotent and invalidates old handles.

## Persistence and migration

IndexedDB holds the editor source of truth; native JSON is a separate compact interchange document. Validate size, schema and unsafe object keys before loading. The sole public JSON document is WorkbookFile, schemaVersion=3. It encodes occupied 64 × 32 physical-ID pages, named homogeneous value columns, numeric style references, sparse ID-keyed metadata and fixed axis runs. Binary pages are the persistent representation; binary native files carry equivalent lossless data. There is no public expanded snapshot export. See [format](storage-format.md). The default database is opensheet; old files and databases are not migrated. Unknown newer schemas must not be accepted as writable data. Remote durability requires an explicit host acknowledgement before showing a saved state.
