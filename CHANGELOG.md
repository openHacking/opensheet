# Changelog

## 0.2.0-dev.0 — Unreleased development candidate

This candidate is not a stable release and has not been published to npm. Review the [capability inventory](docs/capabilities.md), [known limits](docs/implementation-status.md) and [roadmap gates](ROADMAP.md) before deploying it.

### Added and changed

- Expand spreadsheet workflows with multi-key filtering and sorting, fill and paste options, find/replace, annotations, and row, column and freeze controls.
- Expand formula functions and documented coercion, error, date, lookup and reference behavior; retain explicit limits and documented differences from other spreadsheet applications.
- Add native file verification, binary and JSON streaming, safer import boundaries, compatibility reporting and additional XLSX metadata handling.
- Add internal validation and conditional-formatting prototypes; they are not persisted, enforced or available as end-user features. Add bounded inspection and revision-aware mutation APIs.
- Improve Worker recovery, transaction fault handling, cancellation, browser support and keyboard navigation; add installed-package React, Vue, Vite and esbuild integration fixtures.
- Add row-batched snapshot reads and Node JSON reading examples. Batch size bounds the returned matrix, not the complete source file, formula cache or process memory.

### Upgrade boundaries

- The Vue adapter now requires Vue 3.5.42 or later in the 3.x line; upgrade older Vue hosts before installing this candidate. Core now uses Zod 4.6.6.
- This is a preview API and behavior can change before a stable release. Review the [API guide](docs/api.md), [storage version policy](docs/adr/0001-native-format-versions.md) and compatibility reports before upgrading.
- Native schema 3 remains the supported file format; pre-0.1.1 persisted engine layouts and other native schema versions are unsupported. General migration and downgrade are not guaranteed. Retain and verify independent backups.
- Strict cross-device performance, independent spreadsheet compatibility, native IME and assistive-technology gates remain open; local passing tests are not a v1 or industry-leading certification.

## 0.1.1 — 2026-10-08

- Replace editor workbooks with asynchronous Worker/IndexedDB block storage and bounded caches; persist revisions, undo/redo and multi-tab writer coordination.
- Add Performance Lab with deterministic mixed data, configurable storage budgets, cancellation, full recovery verification and capacity bracketing.
- Stream JSON/table exports, stage structural edits/replacements and sort through disk merge passes. Migrate plugins, wrappers, examples and browser benchmarks to async I/O.
- Remove the 200,000 total-cell cap. This breaks synchronous workbook APIs and does not migrate old persisted formats.

## 0.1.0 — Development preview

### Added

- Headless workbook model, command validation, atomic transactions and undo/redo.
- Formula parser and evaluator with reference transforms and resource limits.
- Canvas grid, DOM editor, keyboard and clipboard interactions, frozen panes and zoom.
- SheetJS adapters, Worker imports, compatibility diagnostics and XLSX export.
- LaTeX, Markdown, HTML, CSV, TSV and native JSON workflows.
- Plugin SDK, React/Vue wrappers, schemas and consumer build checks.
- English documentation, OpenSheet branding and four interactive demo scenes.
- GitHub Pages deployment workflow.

### Changed

- Scoped packages now use `@opensheetjs/*`. The main package remains `opensheet` and its stylesheet remains `opensheet/style.css`.

The ten packages are published to npm as version 0.1.0. The API and supported behavior may change before v1.0.
