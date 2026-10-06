<p align="center"><img src="docs/assets/logo.svg" width="280" alt="OpenSheet" /></p>
<p align="center"><strong>An open spreadsheet. A simple API.</strong></p>
<p align="center">Embed an editable spreadsheet in your app. Work with formulas, extend it with plugins, and turn cells into code.</p>
<p align="center">
<a href="https://github.com/openHacking/opensheet/actions/workflows/ci.yml"><img src="https://github.com/openHacking/opensheet/actions/workflows/ci.yml/badge.svg" alt="CI" /></a>
<a href="LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-26a777" alt="Apache 2.0" /></a>
<img src="https://img.shields.io/badge/status-development_preview-64756e" alt="Development preview" />
</p>
<p align="center"><a href="https://openhacking.github.io/opensheet/">Live demos</a> · <a href="docs/api.md">API</a> · <a href="docs/plugins.md">Plugins</a> · <a href="CONTRIBUTING.md">Contributing</a></p>

![OpenSheet — an open spreadsheet, a simple API](docs/assets/social-preview.png)

## Try it

Explore four interactive examples, powered by the same OpenSheet API:

| Demo                                                                | What to try                                                                |
| ------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| [Budget Simulator](https://openhacking.github.io/opensheet/#budget) | Change quantities and prices; watch totals and spending indicators update. |
| [Sales Dashboard](https://openhacking.github.io/opensheet/#sales)   | Edit monthly revenue; see live metrics and a revenue chart.                |
| [Project Planner](https://openhacking.github.io/opensheet/#planner) | Adjust task dates and progress; update a project timeline.                 |
| [Table to Code](https://openhacking.github.io/opensheet/#code)      | Edit cells and export LaTeX, Markdown, HTML, CSV or TSV.                   |

Files are processed in your browser. No account, uploads or telemetry. Download JSON to keep your edits; refreshing restores the sample. Charts and timelines are demo application components.

## Features

- **Embeddable editor** — a virtualized Canvas grid, DOM cell editor, keyboard navigation, selection, clipboard, frozen panes and zoom.
- **Workbook engine** — multiple sheets, typed values, synchronous atomic transactions, undo/redo and a headless TypeScript API.
- **Formulas and formatting** — arithmetic, relative/absolute and cross-sheet references, common functions, cell styles and merged ranges.
- **File interoperability** — SheetJS adapters, browser Worker imports, compatibility reports and XLSX export.
- **Cells to code** — LaTeX, Markdown, HTML, CSV and TSV generation, plus native JSON snapshots.
- **Extensible by design** — a plugin SDK, React and Vue wrappers, testing helpers and JSON schemas.

## Usage

**Version 0.1.0 is a development preview. Packages are not published to npm yet.** The example below applies when consuming the workspace or built package archives; see [Contributing](CONTRIBUTING.md) for setup.

```ts
import { createOpenSheet } from 'opensheet';
import 'opensheet/style.css';

// The container must exist and have an explicit height.
const app = createOpenSheet({ container: '#sheet' });
const book = app.createWorkbook({ sheets: [{ name: 'Sales' }] });
const sheet = book.getSheetByName('Sales')!;

sheet.range('A1:B3').setValues([
  ['Month', 'Revenue'],
  ['January', 1200],
  ['February', 1800],
]);
sheet.range('B4').setFormulas([['SUM(B2:B3)']]);

// Release the instance when your component unmounts.
app.dispose();
```

For headless workbooks, import `createWorkbook` from `@opensheetjs/core`. Framework wrappers use `@opensheetjs/react` and `@opensheetjs/vue`. [Read the API guide](docs/api.md).

## Packages

| Package                                   | Purpose                                                          |
| ----------------------------------------- | ---------------------------------------------------------------- |
| `opensheet`                               | Browser facade, toolbar, formula bar, sheet tabs and plugin host |
| `@opensheetjs/core`                       | Workbook model, commands, transactions, history and schemas      |
| `@opensheetjs/formula`                    | Formula parsing, evaluation and reference transforms             |
| `@opensheetjs/renderer`                   | Canvas grid, DOM editor, input and viewport                      |
| `@opensheetjs/plugin-sdk`                 | Plugin contracts, capabilities, dependencies and cleanup         |
| `@opensheetjs/adapter-sheetjs`            | SheetJS conversion and compatibility reports                     |
| `@opensheetjs/formats`                    | Delimited parsing and table exports                              |
| `@opensheetjs/react` / `@opensheetjs/vue` | Framework lifecycle wrappers                                     |
| `@opensheetjs/testing`                    | Headless workbook and snapshot helpers                           |

## Documentation

- [API guide](docs/api.md) and [plugin guide](docs/plugins.md)
- [Implementation status and limitations](docs/implementation-status.md)
- [Architecture](docs/01-architecture.md) and [data model](docs/02-data-and-commands.md)
- [Verification](docs/verification.md), [changelog](CHANGELOG.md) and [Agent navigation](llms.txt)
- [Workbook schema](schemas/workbook.schema.json) and [command schema](schemas/commands.schema.json)

## Current limitations

OpenSheet does not promise lossless Excel compatibility. Complex styles, charts, macros, array formulas and other unsupported file features have explicit limits. Keep native JSON for supported OpenSheet data. Formula evaluation and number formatting cover a subset of Excel behavior. The project has not reached a stable v1 API or completed full performance and assistive-technology validation. See [implementation status](docs/implementation-status.md).

## Contributing

Bug reports, documentation improvements, tests and independent plugins are welcome. Start with the [contribution guide](CONTRIBUTING.md). Report vulnerabilities through the process in [Security](SECURITY.md), and follow the [Code of Conduct](CODE_OF_CONDUCT.md).

## License

[Apache-2.0](LICENSE). See [third-party notices](THIRD-PARTY-NOTICES.md) for dependency attribution.
