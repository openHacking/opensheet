# Performance Lab protocol

Open `#performance` in the playground. Test one size, automatically find capacity, stop, restore, clear the test workbook or download structured results. Tests use a dedicated `opensheet-performance-v4` database and do not clear other playground data.

## Workload and budget

The deterministic workload uses 100 columns, 60% finite numbers, 30% 32-byte ASCII strings and 10% booleans. Zero and false count as populated. A stateless fixed-seed generator permits full verification without retaining an extra fixture. Formulas, per-cell styles and XLSX conversion are separate workloads.

Default application budget: 256 MiB. IndexedDB stores UTF-8 JSON byte arrays, so the ledger counts the byte lengths of the actual stored application payloads. IndexedDB key/engine filesystem overhead belongs to the separately reported browser usage estimate. Budget checks include encoded data, metadata, history, manifest and temporary records. The UI also reports `navigator.storage.estimate()` origin usage and quota, which are estimates shared with other same-origin storage. Before starting, the application budget may not exceed 80% of estimated remaining quota. If estimates are unavailable, the UI says unknown and still enforces the application budget. A real quota error stops probing immediately.

Blocks store local offsets and compact type tags; row and column identities are stored once per block. Empty unformatted cells have no record. Worker caches reserve at most 64 MiB and the viewport cache at most 8 MiB using conservative encoded-size estimates. These reservations do not claim to be total process heap. Browser heap measurements cover the main document where available; transient import/export working sets and Worker runtime overhead are separate.

## Measurement and search

A size must pass three rounds. Each round closes and reopens storage, waits for the first fully loaded editable viewport, verifies every populated value in batches, exercises edits across blocks with undo/redo checks, and scrolls to dispersed positions.

After five edit warm-ups, record 100 edits from command start through persisted acknowledgement and completed viewport painting. After 30 scroll warm-ups, record 300 scroll samples through viewport data readiness and painting. Thresholds: edit P95 ≤100 ms, complete scroll frame P95 ≤33 ms, restore ≤10 seconds. Full verification time is reported separately. A failing round can end the candidate early because it can no longer pass all three rounds.

Automatic probing starts at 10,000 cells and doubles until a failure or the 10,000,000-cell dimension boundary. It then refines the passing/failing bracket to 10,000-cell precision. One candidate is retained at a time. Cancellation never counts as a capacity failure or pass. A quota error stops immediately; an application-budget failure can be refined. Maximum means the largest observed passing size in this session, not a universal or mathematically proven maximum. Results retain all tested failures, environment information, thresholds, workload and storage figures.

## Reproducible diagnostics

`pnpm benchmark:browser --write-baseline` records fixed-scale diagnostics against the actual Worker/IndexedDB engine; `--sizes=10000,100000,200000,1000000` changes scales. `--compare` requires the same browser, platform, architecture and fixture protocol. The fixture differs from the former all-memory benchmark, so old timings are not directly comparable.

`pnpm benchmark:capacity` drives the complete three-round UI search against the running playground and saves candidate results as it proceeds. Use `--url=http://127.0.0.1:4173/opensheet/` for a production preview, `--budget-mib=256` to set the budget, and `--timeout-ms=3600000` for the run deadline. It keeps the UI thresholds and sample counts unchanged.

The fixed-scale script is a diagnostic, not the three-round capacity acceptance run. The capacity UI's downloaded JSON records each candidate and every completed round. CI verifies correctness and probing behavior in Chromium and WebKit; device-dependent capacity/latency is recorded, not used as a universal CI pass condition.

XLSX is materialized through SheetJS and is excluded from capacity conclusions. Current file import limits remain 20 MiB compressed input and 200 MiB expanded XLSX. Native snapshot adapters also materialize their input; editor storage and generated performance workbooks do not retain full snapshots.
