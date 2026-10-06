# 0.1.0 implementation status

## Available behavior

| Layer            | Current implementation                                                                              |
| ---------------- | --------------------------------------------------------------------------------------------------- |
| Workbook         | Sparse cells, stable row/column IDs, validation, frozen reads, snapshots, multiple sheets           |
| Commands         | Validated reducer, synchronous atomic transactions, revisions, limited instance-local deduplication |
| History          | Immutable forward/inverse patches; 100 transactions / estimated 32 MiB                              |
| Formulas         | Basic functions, references and transforms, cycles, errors and resource limits                      |
| Grid             | Visible Canvas rendering, DOM editor, IME events, selection, keyboard, paste, freeze and zoom       |
| Editing          | Values, formulas, styles, merges, sorting/filtering, copy and fill-down                             |
| Interoperability | Sparse/dense SheetJS, date systems, formulas/caches, notes and safe links                           |
| Demo             | Worker imports, file protection, JSON/XLSX download, four interactive scenes                        |
| Formats          | CSV/TSV, Markdown, escaped HTML, LaTeX tables and supported merges                                  |
| Ecosystem        | Plugins, React/Vue lifecycle wrappers, testing helpers, schemas and Agent navigation                |

## Boundaries

- Core currently depends on the formula package. Calculation is synchronous and revision-cached, without incremental dependency scheduling, a calculation Worker or asynchronous custom functions.
- Merges use positional rectangles updated by structural commands. Partial merge operations fail. Deleted reference endpoints become #REF!; not all Excel structural rules are reproduced.
- Layout rebuilds after commits. There is no interval-tree viewport index, drag-to-resize headers or drag fill handle. Row/column sizes and fill-down have APIs.
- Filters support single-column text inclusion. Formula sort keys are rejected. Data validation, conditional formatting, built-in charts and pivot tables are unavailable. Demo charts/timelines are application components.
- Plugin interfaces cover only documented capabilities. Custom editors, formula functions, shortcut registries, storage providers and plugin scaffolding remain future work.
- There is no IndexedDB autosave. Download JSON or XLSX before leaving; refreshing restores samples. No uploads or telemetry are present.
- Date display and currency/percentage/decimal formats are simplified subsets. Date imports can report approximations.
- SheetJS CE conversion does not preserve all styles, charts, images, macros, named ranges or unknown file parts. Array-formula imports are protected read-only. Strict conversion rejects only detected incompatibilities.
- XLSX export omits potentially stale formula caches and requires reader recalculation. Native JSON preserves supported OpenSheet state.
- LaTeX exports do not compile TeX/PDF, import source, expose arbitrary raw math input or implement full fixed-width layouts. Markdown does not support merged cells.
- Keyboard, ARIA active cells and a separate DOM table view exist. Real VoiceOver/NVDA, native IME and complex touch acceptance are incomplete.
- Only schemaVersion=1 is supported; there are no published historical snapshots requiring migration. Durable idempotency, authentication and remote replication are host concerns.

## Resource protection

Limits are 20 sheets, 200,000 stored cells, 100,000 rows and 1,000 columns per sheet, 200,000 cells per range and 32,768 characters per text value. These protect resources; they do not prove smooth interaction at maximum scale. Mutations exceeding history budgets fail atomically.

Formula parsing limits expression length, tokens and depth; evaluation limits range visits and recursion. Browser parsing uses a Worker with 20 MiB file limit, declared ZIP-size prechecks and a 20-second timeout. These cannot guarantee browser memory isolation. Server hosts need actual CPU/memory limits and process isolation.

## Verification and next steps

See [verification](verification.md) and [Contributing](../CONTRIBUTING.md) for checks. Firefox, native input/screen-reader acceptance, real Excel/WPS/LibreOffice fidelity, full framework-host E2E, leak profiles and repeated browser performance matrices remain incomplete.

Prioritize calculation isolation and incremental indexing, real input/accessibility checks and host tests before expanding persistence or plugin interfaces. Stable v1 requires API diffs, backward-compatibility fixtures, performance baselines and actual integration evidence. Version 0.1.0 is an npm development preview.
