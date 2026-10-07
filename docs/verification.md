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

## Current validation

Run the commands in [Contributing](../CONTRIBUTING.md) against the current revision. Local validation on October 7, 2026 passed formatting, strict TypeScript, 58 unit/property tests, ten package builds and headless consumer checks. Chromium and WebKit passed 48 browser tests, including editor overflow, merged selection, frozen geometry and icons. The production-path check passed all four scene reloads, assets and a real XLSX Worker import under `/opensheet/`. Remote CI and deployment results are reported separately after execution.

The [browser benchmark](../benchmarks/browser-baseline.json) records three workbook loads plus repeated edits, selections and scrolls for 10,000, 100,000 and 200,000 populated cells. It reports p50/p95 timings, long tasks, serialized fixture size and Chromium heap before and after forced garbage collection. Run `pnpm benchmark:browser -- --compare` after `pnpm build` to compare against the saved baseline on the same browser and machine. These measurements do not establish a cross-device responsiveness guarantee or an IndexedDB storage capacity.

## Remaining acceptance

Real screen-reader/native IME testing, Firefox, framework host lifecycle E2E, real spreadsheet-reader acceptance, leak profiling and cross-device performance evaluation remain outstanding. The demo's HTML preview is not TeX rendering, and exported XLSX files need independent reader validation.

## Module boundaries

Run `pnpm test:architecture` to validate the module graph and its rule fixtures. `pnpm check` includes this check, TypeScript, unit tests, package declarations, NodeNext consumer checks, headless package imports and the production build. `pnpm test:e2e` exercises Chromium and WebKit, including repeated workbook replacement and disposal, stale import cancellation, IME, read-only input, merged cells and frozen panes.

Toolbar assertions account for actions moved into the accessible More tools menu. Workbench reset assertions use the implemented default split of 67%.
