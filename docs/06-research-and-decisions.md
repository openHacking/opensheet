# Research and architecture decisions

> Design notes, not a claim of implemented functionality. See [implementation status](implementation-status.md) and the [current API](api.md) for supported behavior.

## Evidence policy

External documentation informs design; it does not prove OpenSheet implements a feature. Verify dependencies against their official documentation and pinned versions. Distinguish prototype observations, design proposals and measured behavior. Never infer compatibility from file-extension support alone.

## Decisions retained for the preview

| Decision | Rationale | Revisit when |
| --- | --- | --- |
| Headless TypeScript core | Reuse data behavior across browser and Node hosts | Runtime constraints require a different boundary |
| Canvas grid plus DOM editor | Virtualized viewport with native text entry | Accessibility or interaction evidence warrants change |
| Validated commands and immutable history | Atomic failure and predictable undo | Measured memory or throughput requires new internals |
| Separate formula package | Deterministic, bounded parsing and evaluation | A provider interface is justified |
| SheetJS interoperability adapter | Keep file handling separate from the model | Fidelity tests identify missing mappings |
| Independent format exporters | Reuse tables without renderer dependencies | Additional format demand is demonstrated |
| Native JSON snapshots | Preserve supported OpenSheet state | A real schema upgrade requires migration |
| Trusted plugins with explicit cleanup | Small understandable extension boundary | Untrusted extensions require isolation |

## Reference material

- [SheetJS documentation](https://docs.sheetjs.com/) for workbook interoperability.
- [GitHub Pages custom workflows](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages) for static demo deployment.
- [TypeScript documentation](https://www.typescriptlang.org/docs/) for public declarations and module resolution.

## Unverified assumptions

The current design prioritizes desktop browser embedding while keeping the demo responsive. Team capacity, stable API requirements and representative workload sizes need evidence from actual integrations. Full Excel compatibility, arbitrary plugins and advanced collaboration are not assumed.

Evaluate demos through correctness of editing and exported data, successful host integration, and clear reports of unsupported features. Avoid treating stars, downloads or visual polish as proof of engine performance or compatibility.
