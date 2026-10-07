# Architecture

The main document owns controls, selection, layout and Canvas painting. It retains immutable sheet metadata and a bounded cache of loaded cell blocks. Unloaded cells show a loading state. Painting never opens IndexedDB and never exports an entire workbook. Viewport and frozen-pane blocks are prefetched; responses from older revisions are discarded.

A module Worker owns IndexedDB access, command validation, formulas, search, filtering, structural changes and disk sorting. Requests are serialized; cancellation is delivered immediately outside that queue. Public I/O methods are asynchronous. Transactions collect commands through a scoped facade so unrelated UI work cannot join a batch.

IndexedDB contains `heads`, `records`, `sizes` and `scratch`. A head references immutable metadata, block records, a revision and undo/redo entries. Blocks encode 64 × 32 cell positions with compact type tags and shared row/column identities. Block dependency records persist parsed formula information; calculated values remain in a bounded cache and are invalidated at every committed revision.

Mutations allocate immutable staged records. Only the final head update makes them visible; its IndexedDB transaction is the commit point. Failed operations delete their staged records; interrupted initial creation is cleaned even when no head exists. Failed post-commit reclamation keeps the durable head and is retried on reopen. A byte ledger is committed atomically with each data record. On writer reopen, recovery scans this workbook's byte ledger without loading cell payloads and removes records unreachable from its head and history. Ordinary edits touch data blocks containing the changed cells, rather than copying every cell; the manifest/metadata and history pointers remain resident.

Undo/redo retains block pointer deltas and metadata versions, constrained to 100 steps and 32 MiB of logical deltas. Structural changes stage relocated blocks and rewritten formulas. Sorting creates bounded sorted runs and merges two disk runs at a time, retaining two rows rather than all run heads. Search stores only the next/first match and total count. Filters store row visibility metadata computed against staged data before publication.

Web Locks coordinates one writer per database/workbook. Other tabs open read-only; BroadcastChannel revision notices invalidate their caches and reload metadata. Worker errors reject outstanding requests, and request deadlines terminate an unresponsive worker. Reopening recovers the last committed head.

Worker block/calculation caches together reserve at most 64 MiB; the main viewport cache reserves 8 MiB. Temporary operation buffers and metadata are separate. See [API](api.md), [performance protocol](performance.md) and [limitations](implementation-status.md).
