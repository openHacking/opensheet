# @opensheetjs/adapter-sheetjs

Part of [OpenSheet](https://github.com/openHacking/opensheet), an embeddable TypeScript spreadsheet.

Version 0.1.1 is a development preview published to npm; this workspace is the unreleased 0.2.0-dev.0 candidate. This package provides ESM and TypeScript declarations. The main browser entry is `opensheet`; scoped packages use `@opensheetjs/*`.

See the [API guide](https://github.com/openHacking/opensheet/blob/main/docs/api.md) and [implementation status](https://github.com/openHacking/opensheet/blob/main/docs/implementation-status.md) for supported behavior. Local setup, builds and tests are documented in [Contributing](https://github.com/openHacking/opensheet/blob/main/CONTRIBUTING.md).

Licensed under [Apache-2.0](LICENSE).
When reading XLSX with SheetJS, enable `bookFiles: true` to report data validation and conditional formatting omitted from worksheet parts declared in the XLSX content-type manifest (with conventional `xl/worksheets/*.xml` paths as a fallback). Strict conversion rejects these losses. Parsed workbook objects without raw parts cannot provide this check. This reports rules; it does not import or preserve them.

XLSX import and export report a potential semantic difference for detected `MATCH` and `VLOOKUP` calls. This conservative warning does not inspect the current lookup data: errors, blanks and wildcard matches can behave differently across applications. Strict conversion rejects the warning. The independent reader cases and their limits are documented in [formula reader findings](../../docs/formula-reader-findings.md). Other lookup functions and syntaxes are not fully audited.
