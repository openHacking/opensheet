# Changelog

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
