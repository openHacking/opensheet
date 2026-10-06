# OpenSheet 0.1 API

This guide describes current behavior. Numbered architecture documents describe long-term goals. When they differ, follow this guide, the type declarations and [implementation status](implementation-status.md).

## Instances and ownership

`createOpenSheet({ container, mode: 'edit' | 'read', toolbar?, onError? })` creates a browser host. The container must exist and have an explicit height. Call `createWorkbook({ sheets? })` or `await load(snapshot)` afterward. Each instance owns one workbook; calling `createWorkbook` twice is rejected.

`createWorkbook` is synchronous. `load` validates a snapshot, replaces the old workbook and reinstalls active plugins. Old workbook and sheet handles become disposed. `ready()` resolves immediately in 0.1. `dispose()` is idempotent and releases UI and listeners.

The browser facade exposes `getWorkbook`, `setMode`, `getGrid`, `selection.get/set/onChange`, `commands.execute/describe/canExecute`, `plugins.get/has/remove`, `use`, `notify` and `on`. `canExecute` checks mode and schema quickly; `execute` performs full semantic validation.

`on` supports `workbook:committed`, `selection:changed`, `plugin:error` and `lifecycle:disposed`, and returns an unsubscribe function. A commit includes commandId, previousRevision, revision, source, commands and changedRanges. In 0.1 changedRanges is empty and layout rebuilds on every commit. Undo/redo events have no commands; events are not a complete replication protocol.

## Workbooks, ranges and transactions

- `getSheets()` returns ordered sheet handles; `getSheetById` and `getSheetByName` return undefined when absent.
- `addSheet(name, { rows?, columns? })`, `removeSheet(id)` and `renameSheet(id, name)` manage sheets.
- `toJSON()` returns a defensive copy. `new Workbook(snapshot)` validates and restores it. `sheetData(id)` is frozen read-only data.
- `execute(command | envelope)` accepts local commands or an envelope with protocolVersion=1, workbookId, commandId and baseRevision. The last 1,000 envelope requests are cached per instance for retry deduplication, not durable idempotency.
- `transaction({ label? }, callback)` produces one revision, event and history entry. Nested transactions, promises and asynchronous I/O are rejected. Any command failure aborts the transaction, even when caught inside the callback.
- `undo()` and `redo()` return success; `canUndo/canRedo` query availability. History retains at most 100 transactions and an estimated 32 MiB of patches; oversized individual mutations fail atomically.
- `setReadOnly(true)` prevents writes and undo at the core layer. Protected array-formula imports cannot be forced back to edit mode.

```ts
import { createWorkbook } from '@opensheetjs/core';

const book = createWorkbook();
const sheet = book.getSheets()[0];
book.transaction({ label: 'Quote' }, () => {
  sheet.range('A1:B1').setValues([[100, 3]]);
  sheet.range('C1').setFormulas([['A1*B1']]);
});
console.log(sheet.range('C1').getValues()); // [[300]]
book.undo();
book.dispose();
```

Numeric coordinates are zero-based; Rect endRow/endColumn are exclusive. A1 addresses are case-insensitive. Selections can be passed as ranges. `setValues` requires a complete rectangle; null is blank, distinct from empty text, zero and false. Literal `=1+1` remains text in `setValues`; `setFormulas` explicitly sets formulas with or without a leading equals sign.

`setInput(text)` applies single-cell UI parsing: equals starts a formula, apostrophe forces text, TRUE/FALSE are booleans, finite numbers become numeric values, and leading-zero or long identifiers remain text. Paste uses the same rules. `getValues` returns scalars or error objects, `getFormulas` returns expressions or null, and `getDisplayValues` returns formatted strings.

```ts
sheet.range('A1:B2').setValues([
  [1, 2],
  [3, 4],
]);
sheet.range('C1').setFormulas([['SUM(A1:B2)']]);
sheet.range('A1:B2').setStyle({ bold: true, background: '#e1efe7' });
sheet.range('C1:C10').fillDown();
sheet.range('A1:B2').copyTo('D1:E2');
```

`clear()` keeps formatting; `clear({ all: true })` removes records. Merge rejects covered nonempty values unless `discardCoveredValues: true` is explicit. Covered cells and partial merged ranges cannot be edited or cleared. `unmerge` removes intersecting merges. Copy/fill operations do not handle merged ranges in 0.1.

## Structure and view

Sheet methods include `insertRows/deleteRows/insertColumns/deleteColumns(index, count=1)`, `setRowHeight`, `setColumnWidth`, `setRowHidden`, `setColumnHidden`, `setFreeze(rows, columns=0)`, `setFilter({ column, query } | null)` and `sort(range, column, 'asc' | 'desc')`.

Sorting is stable: data/styles move and relative formulas use copy semantics; row identities and heights stay fixed. Formula sort keys and merged ranges are rejected. Filtering uses case-insensitive text inclusion, always keeping the first row as a header. Structural edits intersecting merges require unmerging first. Unsupported formula rewrites block structural edits. Deleted reference endpoints conservatively become #REF! rather than reproducing all Excel range shrinking rules.

Styles include bold, italic, underline, six-digit hex color/background, fontSize, align, wrap, border and numberFormat. Wrapping needs sufficient row height; it does not automatically expand rows. Fonts and number formats support a limited subset.

Formulas support common aggregate, logical, conditional, rounding and absolute-value functions, A1/$A$1 and cross-sheet references, ranges, comparisons, concatenation, percentages and arithmetic. There is no arbitrary code, network access or complete Excel coercion. Evaluation is synchronous and cached by revision. `await book.calculation.calculate({ revision? })` validates the revision and calculates; it is not Worker-based.

## SheetJS interoperability

```ts
import * as XLSX from 'xlsx';
import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs';

const imported = fromSheetJS(XLSX.read(buffer, { type: 'array', cellNF: true }));
await app.load(imported.snapshot);
console.log(imported.report);
const exported = toSheetJS(app.getWorkbook().toJSON());
console.log(exported.report); // Show compatibility warnings before downloading.
XLSX.writeFile(exported.workbook, 'result.xlsx');
```

The adapter has no SheetJS runtime dependency and accepts structurally compatible objects. The workspace pins SheetJS CE 0.20.3 from its official archive; the lockfile records integrity. `{ unsupported: 'strict' }` rejects detected approximations, dropped features and blockers; it cannot prove lossless handling of parts not exposed by the parser.

XLSX export does not write potentially stale formula results and reports external recalculation requirements. Preserve native JSON alongside exports. Imported caches display with a dagger at the original revision and expire on changes. Array formulas are read-only. See implementation status for styles, charts and macros.

## Format exports

```ts
import { exportRange, parseDelimited } from '@opensheetjs/formats';
const result = exportRange(book.toJSON(), {
  sheetId: sheet.id,
  range: 'A1:D10',
  format: 'latex',
  options: { booktabs: true, header: true },
});
console.log(result.text, result.report, result.requiredPackages);
```

Formats are latex, markdown, html, csv and tsv. CSV/TSV prefix dangerous formula text by default; `options.safe: false` preserves raw text. Numeric negative values are not rewritten. HTML escapes text and allows limited styles; LaTeX escapes content and reports required packages. Partial merged-range exports are rejected. `parseDelimited(text, separator=',')` handles quoting, multiline fields, BOM and CRLF, returning rectangular string rows.

## Framework wrappers

```tsx
import { OpenSheetView } from '@opensheetjs/react';
import 'opensheet/style.css';
<OpenSheetView
  initialSnapshot={snapshot}
  readOnly={false}
  onReady={(app) => console.log(app)}
  onChange={(snapshot) => save(snapshot)}
/>;
```

```vue
<script setup lang="ts">
import { OpenSheetView } from '@opensheetjs/vue';
import 'opensheet/style.css';
</script>
<template><OpenSheetView @ready="(app) => console.log(app)" /></template>
```

initialSnapshot is read at mount only; later replacements use `app.load`. Set a height through React style or external Vue styling. Wrappers dispose at unmount. Types and built artifacts are checked; full framework-host lifecycle acceptance remains an integration responsibility.
