# @opensheetjs/plugin-sdk

Part of [OpenSheet](https://github.com/openHacking/opensheet), an embeddable TypeScript spreadsheet.

Version 0.1.1 is a development preview published to npm; this workspace is the unreleased 0.2.0-dev.0 candidate. This package provides ESM and TypeScript declarations. The main browser entry is `opensheet`; scoped packages use `@opensheetjs/*`.

Plugin contexts are valid only while their installation is active. Removing a plugin or rolling back a failed setup revokes its context and removes registered resources; retained toolbar actions no longer run. Operations already started by a plugin are not automatically cancelled. Capabilities describe access through the plugin context and do not sandbox arbitrary JavaScript.

Registry installation, removal and disposal are synchronous transactions. A setup or cleanup callback cannot start another registry mutation; schedule it after the current operation returns. Dependency and capability declarations are copied when installation starts.

`setup` must return synchronously; a returned promise is rejected and registered resources are rolled back. The registry captures the returned `dispose` function at installation and invokes it with the returned object as its receiver, even if that object's `dispose` property is later changed. Cleanup callbacks should also finish synchronously: promise rejections are reported, but removal does not wait for asynchronous cleanup. A late disposer returned by an unsupported asynchronous setup is invoked when available, without delaying the failed installation.

The registry retains a copy of each installed definition for workbook replacement, including plugins installed directly through `editor.plugins.use`. `registry.definitions()` returns detached declarations in dependency installation order. Editing declarations on the original definition or a returned declaration does not change the installed declarations; uninstall and install a new definition to change them.

Class instances may implement `Plugin` with prototype methods and getters. Their `setup` method runs with the original instance as its receiver, including after workbook replacement, so private instance state remains available. Identity, dependencies and capabilities are still captured at installation; other mutable instance state remains owned by the plugin.

`checkPluginConformance(() => [myPlugin], testHost)` runs synchronous registry probes for dependency ordering, missing or cyclic dependencies, incompatible API versions, duplicates, rollback and disposal. It also installs the supplied bundle, disposes it, and reinstalls its retained declarations, recording synchronous errors and calls to toolbar and commit-listener cleanup functions. Supply all of the plugin's dependencies and a **disposable test host/workbook**: the helper runs plugin code and may modify that workbook. The result contains `passed` and named `checks` with failure details. This is a preview diagnostic, not proof of leak freedom: it cannot inspect plugin-owned resources, verify that host cleanup implementations actually remove resources, await asynchronous cleanup, simulate editor workbook replacement, or establish compatibility with independent hosts.

If a plugin fails to restore after a workbook is replaced or imported, the workbook operation remains successful. The editor reports the plugin error through `plugin:error`, skips plugins whose dependencies failed, and continues restoring independent plugins. Failed plugins must be installed again explicitly.

See the [API guide](https://github.com/openHacking/opensheet/blob/main/docs/api.md) and [implementation status](https://github.com/openHacking/opensheet/blob/main/docs/implementation-status.md) for supported behavior. Local setup, builds and tests are documented in [Contributing](https://github.com/openHacking/opensheet/blob/main/CONTRIBUTING.md).

Licensed under [Apache-2.0](LICENSE).
