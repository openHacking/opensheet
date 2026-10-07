# Third-party notices

OpenSheet original code is licensed under Apache-2.0. Packages externalize their runtime dependencies; distributed applications must retain the licenses/notices of bundled dependencies.

| Dependency                                                                | Use                                                                               | Declared license                    |
| ------------------------------------------------------------------------- | --------------------------------------------------------------------------------- | ----------------------------------- |
| Immer 11.1.21                                                             | Immutable transactions and history patches                                        | MIT                                 |
| Zod 4.6.5                                                                 | Runtime schemas                                                                   | MIT                                 |
| SheetJS CE 0.20.3                                                         | Playground file parsing/writing; adapter itself has no SheetJS runtime dependency | Apache-2.0                          |
| React 19.2.0 / Vue 3.5.24                                                 | Optional framework adapters and development typechecking                          | MIT                                 |
| Lucide 1.52.0                                                             | Editor and playground icons                                                       | ISC / MIT for Feather-derived icons |
| TypeScript / Vite / Vitest / esbuild / Playwright / fast-check / Prettier | Development, builds and tests                                                     | See installed package LICENSE files |

SheetJS distribution source: https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz. The lockfile pins the archive integrity. SheetJS copyright: Copyright (C) 2012-present SheetJS LLC. Its full license is copied into the built playground third-party notices alongside other bundled dependency licenses.

No competitor source code, graphics, private workbook data or logos were incorporated. The spreadsheet interface and table-generator workflow are independently implemented.
