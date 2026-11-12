# Installed-package framework acceptance

From the repository root, run `pnpm build`, install Playwright's Chromium,
WebKit and Firefox browsers, then run:

```sh
node scripts/framework-acceptance.mjs --output=test-results/framework-acceptance.json
```

The runner packs seven local packages, installs their tarballs in a temporary
consumer outside the workspace, builds it, and tests both Vite development and
production servers. There are no OpenSheet source aliases. A Zod configuration
shim is explicitly aliased for CSP compatibility. The output path must be new;
the report and generated consumer lock preserve the package hashes and exact
dependency resolution. Installation needs registry access or a populated cache.
Ports 5191 and 5192 must be available.

The fixture uses React/react-dom 19.2.0, Vue 3.5.42, Vite 7.3.7 and Zod 4.6.6. React Strict
Mode is enabled in both builds; effect replay is asserted in development, where
React implements that check. This is not certification of every supported
framework version or a heap benchmark.

Copy the `optimizeDeps.exclude` configuration from `vite.config.js` into a Vite
consumer. It preserves the installed core module's relative Worker URL and
avoids partial prebundling of its re-exporting packages. With Vite's defaults,
this fixture reproduced a Worker request under `.vite/deps/engine.worker.js`
and failed initialization. Excluding only core also reproduced a missing named
re-export. The fixture's production build bundles the Worker normally. See
[Vite 7 dependency optimization](https://v7.vite.dev/config/dep-optimization-options).

`initialFile` is read at mount. Changing that prop does not replace an edited
workbook; call the ready instance's `load(file)` for deliberate replacement.
`readOnly` and event callbacks update while mounted. Unmount disposes the editor;
late initialization and action errors must not call the old host. Errors while
mounted must still be delivered.

The harness covers edit callbacks, replacement, writer-lock release on reopen,
pending initialization, ten rapid navigation cycles, invalid initialization,
and action failures before and after unmount.

## Deployment and server rendering

The Vite host runs under `/embedded/sheets/`. Production checks assert that all
observed Worker URLs and the single CSS link use that asset prefix. The preview
server supplies an enforced CSP with same-origin scripts, Workers and styles;
neither `unsafe-inline` nor `unsafe-eval` is enabled. Development mode uses Vite's
normal HMR behavior and is not certified under that production policy.

Copy `csp-zod.js` together with the exact `zod` alias in `vite.config.js` when
using this strict CSP recipe. It imports `zod/v4`, sets `jitless: true`, and
re-exports it before dependent schemas are constructed. Simply calling config
after a static OpenSheet import is too late. The alias also applies to Vite's
Worker build. The host needs a direct, compatible Zod dependency. This setting
affects other consumers of the same Zod instance and disables its dynamic
compilation optimization. Prior benchmarks without this recipe do not establish
its performance. See [Zod's CSP guidance](https://zod.dev/compile).

Without the shim, the Chromium fixture worked but recorded an `eval` CSP
violation during Zod's capability probe. That failure is retained. The runner
asserts zero document `securitypolicyviolation` events with the shim, checks the
actual response policy and exercises Worker-backed edits and reopen. This does
not instrument every Worker security-policy event or certify every host policy.

`ssr.mjs` runs from the installed consumer using Node without DOM, Worker or
IndexedDB globals. It imports both wrappers, renders their empty host divs with
the React and Vue server renderers, and verifies that no ready callback runs.
It also reads an empty cell through the pure FileReader API. The runner embeds
this server-rendered markup in actual `react.html` and `vue.html` build entries.
The production hydration cases load those responses and use React `hydrateRoot`
or Vue `createSSRApp`, assert preservation of the server host element, no
recoverable React errors or console errors, and then exercise the same full
editing/replacement/unmount flow. All 18 cases (three browsers × two frameworks
× development, production mount and production hydration) use installed packages.
This is not a Next/Nuxt routing or React Server Components integration.

Both wrappers use the shared `os-framework-host` CSS class for the default 500px
height, avoiding inline style attributes in server HTML. Import
`opensheet/style.css` once in the client entry. Under a strict no-inline-style
policy, customize server-rendered sizing through a stylesheet class rather than
the React `style` prop or Vue style attribute. Inline styles explicitly supplied
by the host remain its CSP responsibility. Keep Worker-backed workbook creation
out of server render paths. Broader framework versions and bundle-size acceptance
remain separate work.
