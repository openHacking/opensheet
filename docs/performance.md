# Performance Lab protocol

Open `#performance` in the playground. Test one size, automatically find capacity, stop, restore, clear the test workbook or download structured results. Tests use a dedicated `opensheet-performance` database and do not clear other playground data.

## Workload and budget

The deterministic workload uses 100 columns, 60% finite numbers, 30% 32-byte ASCII strings and 10% booleans. Zero and false count as populated. A stateless fixed-seed generator permits full verification without retaining an extra fixture. Formulas, per-cell styles and XLSX conversion are separate workloads.

Default application budget: 256 MiB. IndexedDB stores binary cell pages and tagged UTF-8 JSON metadata/indices, so the ledger counts the byte lengths of the actual stored application payloads. IndexedDB key/engine filesystem overhead belongs to the separately reported browser usage estimate. Budget checks include encoded data, metadata, history, manifest and temporary records. The UI also reports `navigator.storage.estimate()` origin usage and quota, which are estimates shared with other same-origin storage. Before starting, the application budget may not exceed 80% of estimated remaining quota. If estimates are unavailable, the UI says unknown and still enforces the application budget. A real quota error stops probing immediately.

Pages use typed value columns, offset runs/bitmaps, string arenas/dictionaries and bit-packed booleans; page coordinates implicitly reconstruct stable uint32 row/column identities. Background gzip applies only when it saves at least 10%. Empty unformatted cells have no record. Worker caches reserve at most 64 MiB and the viewport cache at most 8 MiB using page buffer/position/property/dependency reservations. These reservations do not claim to be total process heap. Browser heap measurements cover the main document where available; transient import/export working sets and Worker runtime overhead are separate.

## Measurement and search

A size must pass three rounds. Each round closes and reopens storage, waits for the first fully loaded editable viewport, verifies every populated value in batches, exercises edits across blocks with undo/redo checks, and scrolls to dispersed positions.

After five edit warm-ups, record 100 edits from command start through persisted acknowledgement and completed viewport painting. After 30 scroll warm-ups, record 300 scroll samples through viewport data readiness and painting. Thresholds: edit P95 ≤100 ms, complete scroll frame P95 ≤33 ms, restore ≤10 seconds. Full verification time is reported separately. A failing round can end the candidate early because it can no longer pass all three rounds.

Automatic probing starts at 10,000 cells and doubles until a failure or the 10,000,000-cell dimension boundary. It then refines the passing/failing bracket to 10,000-cell precision. One candidate is retained at a time. Cancellation never counts as a capacity failure or pass. A quota error stops immediately; an application-budget failure can be refined. Maximum means the largest observed passing size in this session, not a universal or mathematically proven maximum. Results retain all tested failures, environment information, thresholds, workload and storage figures.

## Reproducible diagnostics

`pnpm benchmark:browser --write-baseline` records fixed-scale diagnostics against the actual Worker/IndexedDB engine; `--sizes=10000,100000,200000,1000000` changes scales. `--compare` requires the same browser, platform, architecture and fixture protocol. The fixture differs from the former all-memory benchmark, so old timings are not directly comparable.

`pnpm benchmark:capacity` drives the complete three-round UI search against the running playground and saves candidate results as it proceeds. Use `--url=http://127.0.0.1:4173/opensheet/` for a production preview, `--budget-mib=256` to set the budget, and `--timeout-ms=3600000` for the run deadline. It keeps the UI thresholds and sample counts unchanged.

The fixed-scale script is a diagnostic, not the three-round capacity acceptance run. The capacity UI's downloaded JSON records each candidate and every completed round. CI verifies correctness and probing behavior in Chromium and WebKit; device-dependent capacity/latency is recorded, not used as a universal CI pass condition.

XLSX is materialized through SheetJS and is excluded from capacity conclusions. XLSX/CSV import limits remain 20 MiB input and 200 MiB expanded XLSX. Native JSON/binary Blob/stream imports are page staged and have no fixed 20 MiB file limit; configured storage and 32 MiB page/metadata limits apply. Full toJSON and SheetJS conversions materialize output explicitly. Editor storage and generated performance workbooks do not retain a complete scalar-object snapshot. Use --cells=10000000 --output=benchmarks/capacity-binary-10m.json for a single-candidate three-round run.

## Binary-engine observations — October 7, 2026

The [codec run](../benchmarks/storage-binary.json) used the deterministic mixed fixture. At 1,000,000 cells, canonical JSON was 22,772,634 bytes, raw pages 16,280,720 bytes and native binary with selective gzip 12,220,001 bytes. Encoding plus compression and verification took about 3.54 seconds in Node 22. Empty maximum-dimension JSON was 385 bytes. Payload sizes exclude IndexedDB overhead/history and should not be confused with engine resident memory.

[Browser diagnostics](../benchmarks/browser-binary.json), macOS arm64, Chromium 153, recorded 1m/10m edit P95 19/19 ms, restore P95 85/234 ms, stored application payloads 14.19/150.24 MiB and main-document retained heap 10/11 MiB. Worker cache reservation remained 11.69 MiB and viewport 6.94 MiB. Cold scrolling P95 was 36.1/36.3 ms, above the 33 ms target. This diagnostic is not an acceptance pass: it repeatedly edits one cell and does not perform the three complete verification rounds. Main-document heap excludes Worker/IndexedDB runtime heap. Historical capacity results apply to the previous engine. There is no matched competitor benchmark supporting a “strongest spreadsheet” claim.

The stricter single-candidate runs ([1m](../benchmarks/capacity-binary-1m.json), [10m](../benchmarks/capacity-binary-10m.json)) perform a complete value verification before dispersed edits with undo/redo. Both correctly stop after the first failed round: at 1m, edit P95 65.2 ms, scroll P95 35.3 ms, restore 133 ms, verification 3.55 s, payload 11.85 MiB; at 10m, edit P95 197 ms, scroll P95 52.8 ms, restore 217 ms, verification 38.34 s, payload 117.88 MiB. Neither passed the three-round acceptance protocol. These failures retain the unchanged 100/33 ms thresholds. Cold page decoding/cache churn, history/viewport costs and broader end-to-end profiling remain optimization targets; successful storage/round-trip validation alone does not establish interactive capacity.

To reproduce native I/O validation, serve the built playground with `pnpm preview --port 4174`, then run `node scripts/native-io-benchmark.mjs --url=http://127.0.0.1:4174/ --cells=1000000`. The output [native-io.json](../benchmarks/native-io.json) records complete JSON, streaming output bytes, Blob imports and all-cell checksums. Run benchmarks alone, without concurrent browser tests or rebuilding their assets.
