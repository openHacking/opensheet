# Local candidate host

This Vite example installs the unpublished `0.2.0-dev.0` archives directly. It uses pnpm overrides for internal packages because these candidate versions are not available from the registry. Third-party dependencies such as Vite and Lucide may still be downloaded.

From the repository root, copy the example and five archives into a new directory outside the workspace:

```sh
cp -R examples/local-candidate /path/to/new-host
mkdir /path/to/new-host/archives
cp artifacts/opensheet-0.2.0-dev.0-attachment-fix/{opensheet,core,formula,plugin-sdk,renderer}.tgz /path/to/new-host/archives/
cd /path/to/new-host
pnpm install
pnpm build
pnpm preview --host 127.0.0.1
```

Open the `/embedded/sheets/` path shown by Vite. `package.json` is a standalone example: an existing host should merge its dependencies and pnpm overrides rather than replace its own manifest. File paths in overrides are relative to the host project. The generated `pnpm-lock.yaml` should be retained by the host; use `pnpm install --frozen-lockfile` after it exists.

This example does not configure strict CSP or server rendering. See the [API guide](../../docs/api.md) and [candidate packaging notes](../../docs/candidate-packaging.md) for those boundaries and the candidate's verification status.

For the edit, reopen and backup workflow, see [Getting started](../../docs/getting-started.md). The example stores its workbook ID in localStorage and reopens it on reload. If opening fails, it reports the error and keeps the ID for recovery.

Use **Download backup** and **Restore backup** for native JSON files up to 64 MiB in this example. Restore creates a new local workbook; a successful restore saves its new ID.
