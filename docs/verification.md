# Verification

## Historical baseline — October 5, 2026

The original local validation used macOS x64, Node.js 24.11.1 and pnpm 10.22.0. CI targets Node.js 22. Historical results are not evidence that a later revision or remote workflow passed.

| Check                            | Recorded result                                           |
| -------------------------------- | --------------------------------------------------------- |
| Formatting and strict TypeScript | Passed                                                    |
| Vitest                           | 56 tests across five files, including property tests      |
| Playwright                       | 11 tests each in Chromium and WebKit                      |
| Production build                 | Ten ESM packages, declarations and demo                   |
| Package consumption              | Headless imports and NodeNext declarations                |
| Package archives                 | Main package includes JS, types, CSS, README and license  |
| Dependency audit                 | No reported issues in that scan; not a security guarantee |

Coverage included initialization, formulas, edits and history, sheets, read-only mode, rectangular paste, synthetic composition events, real Worker XLSX imports, downloads, hostile JSON rejection, snapshot restoration, plugin cleanup, accessible table view, narrow layouts and distant-cell navigation.

Historical fixes included workspace dependency bundling, portable React return types and NodeNext declaration imports. Core regressions covered selection metadata, formulas without caches, sheet invariants, unsafe keys, copy/fill operations and partial merged-range export rejection.

The [headless benchmark](../benchmarks/baseline.json) records one diagnostic run. It does not establish p95 browser responsiveness or a guaranteed scale.

## Current validation — October 7, 2026

The paged-engine rewrite passed `pnpm check`: formatting, strict TypeScript, 81 unit/property/IndexedDB contract tests, four architecture rule tests, module boundaries, ten package builds, declarations, NodeNext consumer checks, pure snapshot smoke checks and the Vite production Worker bundle. The `/opensheet/` production-path check passed all five scene reloads, assets and a real XLSX Worker import. All 70 browser tests passed across Chromium and WebKit, including the complete performance example and real Worker fault injection. The TestWorker used by unit tests invokes the exact worker dispatcher with fake IndexedDB; it does not pretend to measure browser thread or storage performance.

[Fixed-scale diagnostics](../benchmarks/browser-baseline.json) cover 10,000, 100,000, 200,000 and 1,000,000 mixed populated cells. The 1,000,000-cell case recorded edit P95 34 ms, complete scroll-frame P95 23.8 ms and restore P95 50 ms, with Worker cache reservation 57.72 MiB and viewport reservation 7.38 MiB. Main-document retained heap was 9 MiB; this excludes Worker runtime heap. These are observations on macOS arm64 and headless Chromium 153, not guarantees for another device. The fixed-scale diagnostic edits one cell repeatedly; the capacity example separately exercises dispersed blocks and three full rounds.

The [complete capacity run](../benchmarks/capacity-baseline.json) used the production build, default 256 MiB budget and all three-round acceptance checks. It converged to **6,870,000 passed / 6,880,000 failed**, at 10,000-cell precision. The largest passing case recorded edit P95 99.8 ms, complete scroll-frame P95 23.5 ms, restore 99 ms and 196.66 MiB of actual encoded payloads (about 30.02 bytes per cell). The next tested scale failed the editing latency threshold. The 10,000,000-cell candidate separately hit the application budget boundary. These are this session's observations; the timings are close to the threshold, so the bracket is not a reproducible guarantee for every run or device. All candidates and round metrics are retained.

## Adversarial review evidence

| Failure challenged                                     | Verification                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| ------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Unloaded cells mistaken for blanks                     | `storage.test.ts` distinguishes loading, empty, zero and false after eviction/reopen, and evaluates cross-block/cross-sheet formulas, lazy branches and cycles. Raw cell/input/formula-text reads do not trigger calculation.                                                                                                                                                                                                                              |
| Cancelled/interrupted operations leave partial commits | Contract tests cancel staged generation, preserve the old head and reclaim orphan records. Browser tests kill the actual storage Worker mid-generation, reopen the old committed value, and inject a native quota exception. Post-commit reclamation failure also keeps a durable commit successful.                                                                                                                                                       |
| Cache or formula dependencies grow without bounds      | ByteCache oversize/eviction tests, bounded formula cache, persistent per-block dependency records, fixed-scale cache reservations, and rapid browser scrolling with data rereads. A spy proves a one-cell edit above the former 200,000-cell limit reads at most four blocks.                                                                                                                                                                              |
| Staging exceeds the application budget                 | Synchronous ledger-write exceptions explicitly abort already queued data writes. Exact ledger accounting includes encoded keys/data, byte-ledger entries, history, metadata and head; budget exhaustion during generation preserves the original revision. Sort scratch space and replacement allocations are checked before writes. A large blank-range style command is rejected before block reads or cell allocation by the traversal planning budget. |
| Timing omits loading or painting                       | Performance samples await `CanvasGrid.ready()` after persisted edits and dispersed scrolls. Restore includes opening and painting the first loaded viewport; full streamed validation has a separate timer. Browser tests download round-level measurements and verify refresh restoration.                                                                                                                                                                |

Automatic search has tests for doubling, bracketing, 10,000-cell refinement, quota termination, cancellation exclusion and dimension boundaries. Browser test output is retained in `test-results/`; capacity results are downloadable JSON, with environment, workload, byte accounting and all rounds.

## Remaining acceptance

Real screen-reader/native IME testing, Firefox, framework host lifecycle E2E, real spreadsheet-reader acceptance, leak profiling and cross-device performance evaluation remain outstanding. The demo's HTML preview is not TeX rendering, and exported XLSX files need independent reader validation.

## Module boundaries

Run `pnpm test:architecture` to validate the module graph and its rule fixtures. `pnpm check` includes this check, TypeScript, unit tests, package declarations, NodeNext consumer checks, headless package imports and the production build. `pnpm test:e2e` exercises Chromium and WebKit, including repeated workbook replacement and disposal, stale import cancellation, IME, read-only input, merged cells and frozen panes.

Toolbar assertions account for actions moved into the accessible More tools menu. Workbench reset assertions use the implemented default split of 67%.
