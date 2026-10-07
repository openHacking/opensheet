# Implementation status

The default engine is a Worker-backed IndexedDB workbook with 64 × 32 blocks, compact disk records, a 64 MiB Worker cache reservation and an 8 MiB viewport cache reservation. The former 200,000 total-cell limit and whole-snapshot node-count limit are removed. Row, column, sheet, text, read-batch and operation resource limits remain.

All playground scenes, the editor, formula bar, toolbar, plugins, React/Vue wrappers, file adapters and examples use asynchronous workbook I/O. Changes persist locally and scene identifiers restore the corresponding workbook on refresh. The Performance Lab uses a separate database and deterministic mixed data; see [its protocol](performance.md).

Supported behaviors include typed values, cross-sheet formulas, formats, merges, frozen panes, search, atomic replacement, filters, disk-backed sorting, structural reference updates, atomic command batches, persistent undo/redo and multi-tab single-writer coordination. Revisions are published only after storage commits; staged failures and cancellation preserve the previous revision. Recovery removes orphan records left by interrupted workers.

Full snapshot imports and `toJSON()` are explicit materialization boundaries. JSON and table output can be streamed by input batches; Blob download output still consumes memory. SheetJS XLSX conversion is materialized and reports unsupported features. Complex styles, charts, macros, external/named references, array formulas and full Excel compatibility remain outside supported behavior. Sorting directly on formula results is still rejected.

Ordinary mutations use a 32 MiB traversal planning budget (512 reserved bytes per visited cell), a 32 MiB working-set/undo delta budget, and single read batches are limited to 16,384 cells. Large structural operations and sorting stage on disk but remain subject to the undo and total storage budgets. Formula evaluation has a 128-dependency-depth and 10,000,000-reference work budget. Metadata for up to 100,000 rows remains resident; caches are bounded, not a guarantee that total browser heap never exceeds the cache budgets.

Browser Workers, IndexedDB and Web Locks are required for an editor workbook. Pure snapshot readers, schemas, formulas and file adapters can be used in Node; custom runtimes need an injected Worker and an IndexedDB implementation. No old persisted-engine migration or synchronous workbook API compatibility is provided.
