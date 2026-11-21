# Edit, reopen and back up a workbook

This guide targets the **unreleased 0.2.0-dev.0 candidate**, not the published 0.1.1 preview. The [local candidate example](../examples/local-candidate/) is a small Vite host that installs local package archives. Its JavaScript is the runnable example for this guide; use the [published-version README](../README.md#usage) only with the matching published API.

## Run the example

You need Node 22 or later, pnpm 10.22.0, a browser with module Workers, IndexedDB, Web Locks and BroadcastChannel, and the local `opensheet-0.2.0-dev.0-attachment-fix` candidate archives. These locally staged archives are a specific build, not a guarantee that they contain later workspace changes. Follow the copy, install and preview commands in the [example README](../examples/local-candidate/README.md). Open the `/embedded/sheets/` URL printed by Vite. The host uses a non-root path intentionally; it needs a server origin such as localhost, not a `file:` URL. Installation may download third-party packages.

The example imports the editor and its stylesheet. On first launch it creates a 10-row workbook, writes `7` and `5` to A1 and B1, and writes a formula to C1; the displayed result is `12`. Later launches reopen the saved workbook. Its container has an explicit 500 px height. Click a cell and edit it; a successful write is committed to browser storage. The complete, tested implementation is [main.js](../examples/local-candidate/main.js).

## Keep the workbook and reopen it

A workbook has a local `book.id`. Store that ID in your application and use `openWorkbook(id)` to reopen the same local workbook after closing or disposing the previous instance. `await book.close()` releases the writer lock; `app.dispose()` also releases the editor and its workbook. Neither is a backup or deletes the saved workbook. The example stores its ID in localStorage and reopens it after a page reload. If reopening fails, it displays the error and retains the ID instead of silently creating another workbook. A production host should decide how to manage workbook IDs and offer an explicit recovery path. See [creation, attachment and lifecycle](api.md#create-restore-and-close) for the available APIs.

Only one instance is the writer at a time. Another instance may open read-only until the writer closes and the reader acquires the lock. Wait for writes to finish and surface `book.persistenceState` or `onPersistence` in a real host; `saved` means an acknowledgement was received, not that browser storage can never be evicted. See [storage failure and recovery](api.md) for error codes and the retry boundaries.

## Save a portable copy

Choose **Download backup** to save the complete native workbook as JSON outside the local database. To restore, choose **Restore backup** and select that file. A successful restore creates a new local workbook and saves its ID for the next reload; the prior workbook remains in local storage. A malformed file leaves the current workbook in place. If the workbook was restored but its ID could not be saved, the example reports that distinction and asks you to export a backup before refreshing.

The example uses `book.toJSON()` and `app.load(file)`, so it materializes the complete workbook. Its download and restore controls each limit the file to 64 MiB; larger workbooks need an application designed for the [streaming export APIs](api.md#native-json-and-binary-io) and an appropriate import workflow. The underlying API and the example limit are separate.

This native backup is distinct from XLSX. When exchanging spreadsheets with other applications, inspect the [file compatibility and loss report](04-sheetjs-and-formats.md); OpenSheet does not promise lossless Excel compatibility.

## If opening fails

- `WORKER_UNAVAILABLE` or `WORKER_FAILED`: verify the Worker asset URL and the site's security policy, then reopen the same ID. The message identifies possible causes, not a diagnosis.
- `STORAGE_UNAVAILABLE`, `WEB_LOCKS_UNAVAILABLE` or `BROADCAST_CHANNEL_UNAVAILABLE`: check the browser capabilities and permissions, and use a supported origin. Private browsing and storage policies can restrict them.
- A quota or budget failure is not a successful save. Keep any unsaved input, free space or adjust the configured budget, and retry only after checking the reported state. An exported copy is separate from local storage.

See the [API guide](api.md), [runtime capability matrix](runtime-capabilities.md), and [implementation limitations](implementation-status.md) before relying on a specific deployment or file feature. The example and its automated verification do not establish strict CSP, server rendering, native accessibility, or support for every bundler and browser version.
