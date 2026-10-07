# SheetJS and table formats

> Design notes, not a claim of implemented functionality. See [implementation status](implementation-status.md) and the [current API](api.md) for supported behavior.

## Ownership and supported fidelity

SheetJS reads/writes file containers; OpenSheet owns the editable model and UI. The adapter accepts structurally compatible SheetJS workbooks without depending on the SheetJS runtime. The demo pins CE 0.20.3 from its official archive, including lockfile integrity.

Import supports sparse/dense cells, sheet order, scalar values, formulas, imported caches, date systems, dimensions, hidden rows/columns, merges, notes and safe links within explicit limits. Complex styles, charts, images, macros, named ranges and unknown file parts are not guaranteed to survive conversion. Array formulas force read-only behavior.

The IndexedDB workbook is the source of truth after editing; native OpenSheet snapshots are interchange data. Keeping an original file is a separate host responsibility. Do not merge arbitrary original workbook parts and claim lossless XLSX patching.

## Compatibility reports

Conversions return reports of supported entries, approximations, dropped features and blockers. Show meaningful summaries and detailed locations; aggregate repeated issues to avoid unbounded report memory. Strict mode rejects detected incompatibilities but cannot guarantee preservation of unexposed file parts.

Imported formula caches are marked and expire after edits. Export avoids potentially stale formula results and requests external recalculation where appropriate. Real Excel, WPS and LibreOffice acceptance needs independent testing.

## Browser import pipeline

The demo checks file size and declared ZIP expansion, transfers bytes to a Worker, parses, adapts, validates and atomically loads the snapshot. The current file limit is 20 MiB and timeout is 20 seconds. On errors, cancellation or timeout the previous workbook must remain intact.

ZIP declarations can be dishonest; Worker timeout and prechecks do not guarantee browser memory isolation. Server integrations need process isolation and actual CPU/memory limits. Formula length, depth, range visits and workbook dimensions also have resource budgets.

## Dates and numeric values

Preserve 1900/1904 systems. Test boundary serials, fractions and timezone semantics rather than treating spreadsheet serials as JavaScript epoch timestamps. Long IDs and integers beyond safe precision should remain text. Current display formatting is a subset, not the full Excel SSF language.

## Text exports

`exportRange(snapshot, { sheetId, range, format, options })` returns text, diagnostics and required LaTeX packages where applicable. It reads model data rather than Canvas pixels.

- LaTeX supports tabular, booktabs and supported merged cells; escape ordinary text and report package requirements. There is no TeX/PDF compilation or source import.
- Markdown uses the first row as a header, escapes pipes and line breaks, and reports unsupported merges.
- HTML uses semantic tables, escaped text and a style allowlist rather than arbitrary scripts.
- CSV/TSV supports quotes, multiline text, BOM and CRLF. Safe mode prefixes dangerous formula text; explicit raw mode preserves values.

`streamExport(book, config)` reads bounded asynchronous batches without materializing the entire workbook; `exportRange` remains the explicit snapshot convenience. XLSX remains a SheetJS materialization boundary and does not establish engine capacity. Exports cutting through a merged range are rejected. HTML grid previews are not actual TeX rendering. Future formats should remain separate adapters instead of expanding core responsibilities.

## References

[SheetJS cell model](https://docs.sheetjs.com/docs/csf/cell/), [formulas](https://docs.sheetjs.com/docs/csf/features/formulae/), [dates](https://docs.sheetjs.com/docs/csf/features/dates/) and [installation](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/) describe the upstream format contracts.
