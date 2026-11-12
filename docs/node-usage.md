# Node and headless use

The pure file and formula utilities can run in Node without a DOM, browser Worker or IndexedDB. A full editable `Workbook` is different: its default implementation uses a browser module Worker, IndexedDB, Web Locks and BroadcastChannel. Node does not provide that complete browser environment by default, and OpenSheet does not currently provide or support a Node persistence adapter. Importing the package in Node does not imply that `createWorkbook()` will work there.

With `@opensheetjs/core` installed in a Node 22 or newer project, copy [node-file-reader.mjs](../examples/node-file-reader.mjs) into that project and run:

```sh
node node-file-reader.mjs ./workbook.json 'Sheet1' 'A1:C10'
```

The input is an OpenSheet native **JSON** workbook, not an XLSX file or `.opensheet` binary. The example validates the file, selects a sheet by name, evaluates the requested range and prints JSON. It rejects ranges over 10,000 cells as an example-specific limit. Formula results follow OpenSheet's supported formula semantics and may contain error objects. `FileReader` is read-only and `dispose()` clears its calculation and page caches.

This example reads and parses the entire JSON file before validation. The reader's bounded page cache is not a bound on the input object, metadata, formula work or process memory. Use only files appropriate for the available memory and separately bound untrusted inputs. For native binary metadata and page access, see `BinaryReader`; do not pass binary files to this JSON example.

For larger rectangles, use the synchronous `range.batches({ maxCells: 4096 })` iterator and consume each batch before requesting the next. Each batch is at most the configured cell count, but formula caches, the already materialized JSON and batches retained by the caller remain outside that limit. See the [API guide](api.md) for bounds and display options.

The existing [headless.ts](../examples/headless.ts) demonstrates the editable asynchronous workbook API in a browser-capable environment. A custom host may supply `workerFactory`, but must implement the expected Worker messaging and the required storage and coordination facilities. Merely supplying a Node `worker_threads.Worker` is not a supported adapter, and the project does not claim full Node-engine compatibility or persistence durability for arbitrary shims.

The [isolated consumer check](../scripts/node-consumer-acceptance.mjs) packs the built core and formula packages, installs them in a temporary project using the offline pnpm store, and runs the example through the installed package. It checks the formula output, batched reads, two error cases, package resolution and the missing browser-runtime error. Reproduce from the repository after `pnpm build` with `node scripts/node-consumer-acceptance.mjs benchmarks/NEW-REPORT.json`; its offline install requires the external dependencies to be cached. The temporary project is removed, while the report and its consumer lockfile are retained. One local Node run is recorded in [the report](../benchmarks/node-consumer-2026-11-12-batches.json); it is not a test of every Node version or a Node storage adapter.
