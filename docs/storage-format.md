# Native format and storage (schema 3)

`WorkbookFile` is the single public, lossless JSON interchange shape. `toJSON()` returns the complete occupied-cell document; `streamJSON()` emits that same document in page-sized chunks. There is no public `toSnapshot()` or expanded workbook snapshot conversion. The binary `.opensheet` container carries equivalent data and is the default save format. JSON is `.opensheet.json`. Older files and persisted engine formats are rejected; no migration or fallback is provided. `schemaVersion: 3` identifies this format, not a compatibility layer. Database `opensheet` and scene key `opensheet.scene.<id>` have no version suffix.

The generated [JSON schema](../schemas/workbook.schema.json) describes shape; `validateWorkbookFile()` also verifies semantic constraints such as duplicate identities, offset collisions, references and dimension bounds. Hosts and agents should use [the API](api.md) for ordinary reads and writes.

## Identity and axes

Each sheet has stable uint32 row and column IDs, allocated monotonically from `nextId`. IDs are sheet-local; they are not current display positions. Insert, delete, reorder and undo preserve surviving identities. A1 and command rectangles address current positions; rectangles use zero-based, exclusive ends.

`rows` and `columns` contain `{ nextId, order, meta? }`. Each order run is exactly `[startId, count, step]`, with `step` equal to 1 or -1. Concatenating runs produces display order; every ID must be unique and smaller than `nextId`. A singleton is `[id, 1, 1]`. Empty axes are invalid. Metadata is a sparse object keyed by stable numeric ID, for example `"7": { "size": 36 }`. JSON object keys are strings, even when they represent integers. Sheets appear in display order directly; there is no duplicate `sheetOrder` array.

The runtime uses chunked sequence trees and sparse typed reverse indexes for identity/position lookup. Axis metadata remains resident as arrays for UI and command processing; structural changes can still copy these arrays. Row/column limits are 100,000/1,000 per sheet, at most 20 sheets; stable IDs can outlive these dimensions through repeated insert/delete.

## JSON pages

Only occupied physical pages appear in `sheet.blocks`. Each block is `{ row, column, data }`. Pages cover 64 stable row IDs by 32 stable column IDs. A local offset is `rowId % 64 * 32 + columnId % 32`; reconstruct IDs as `block.row * 64 + floor(offset / 32)` and `block.column * 32 + offset % 32`. Consult axis order to locate a cell on screen. A block is not a rectangular display-order slice after reordering.

`data` uses named, homogeneous columns:

```json
{
  "row": 0,
  "column": 0,
  "data": {
    "numbers": { "positions": [1, 33], "values": [18000, 23500] },
    "strings": { "positions": [0, 32], "values": ["Month", "January"] },
    "formulas": { "positions": [34], "values": ["B2*2"] },
    "styles": [[0, 2, 0]]
  }
}
```

Supported input columns are `numbers`, `strings`, `booleans`, `formulas` and `errors`. `positions` and `values` have equal length; offsets are unique across input columns. Formula text excludes the leading `=`. An empty string, zero, false and literal formula-looking text remain distinct. `blanks` lists explicit blank records; an absent cell has no record. A blank with a style, note or link must remain stored.

Optional property columns `notes`, `links`, `formats`, `cached` use the same `{ positions, values }` shape and must refer to existing inputs or explicit blanks. Link values are `{ target, tooltip? }`; cached values are tagged scalars such as `{ type: "number", value: 42 }`. `styles` contains fixed triples `[startOffset, count, styleIndex]`, without overlaps. The global `styles` array contains property objects; its numeric indexes are file-local. Workbook exports omit unused styles. Runtime style identity is internal, not a stable file contract.

All numbers must be finite. JSON would otherwise lose negative zero: corresponding `numbers.values` contain 0 and `negativeZero` lists their offsets. `cachedNegativeZero` does the same for numeric cached values. Arbitrary UTF-16 strings, including unpaired surrogates, survive round trips. Imported cached formula results carry revision provenance and are dropped from exports after any edit. Derived calculation and viewport caches are not stored in the file.

## Binary pages and container

All multibyte integers and numbers use little-endian byte order. An OSP3 page starts with bytes `79,83,80,3`, then uint32 metadata length. JSON metadata describes sparse properties, local styles and sections. The header and each section align to eight bytes. Positions choose the smallest of uint16 offsets, uint16 `[start,count]` runs or a 2,048-bit bitmap. Numeric values use lossless signed int8/int16/int32 when possible and float64 otherwise; negative zero uses float64. Booleans use bitmaps. Strings use offsets into UTF-8 arenas or page-local dictionaries, whichever is smaller. Pages containing unpaired surrogates use lossless UTF-16LE arenas. Pages never retain a full scalar-object map in the decoded cache.

The binary file starts with `79,83,66,3`, uint32 metadata byte length and uint32 CRC32. Its UTF-8 metadata is the JSON document with empty blocks. Each page frame has seven uint32 fields: sheet index, physical row, physical column, codec, encoded length, decoded length and encoded CRC32, followed by the payload. Codec 0 is raw OSP3; codec 1 is gzip. Compression is chosen only for at least 10% savings. A directory marker consists of uint32 `0xffffffff` and uint32 directory length. The UTF-8 JSON directory contains `{sheet,row,column,offset,length}` entries; offset points to the frame and length includes its 28-byte header. The 24-byte footer is ASCII OSBE, uint64 directory offset, uint64 directory length and uint32 directory CRC32. `BinaryReader.open()` reads metadata/directory ranges; `page(index)` retrieves a page independently. Sequential import validates the directory against frames already read.

IndexedDB records have a one-byte discriminator: 0 for UTF-8 JSON `[key,value]`, 1 for binary payload. Uncompressed pages use OSP3; background compression stores OSG3 followed by uint32 decoded length, uint32 decoded CRC32 and gzip bytes. Background compression replaces one record and its size ledger atomically without changing business revision. Internal metadata carries a style-identity side table so page style references survive reopen; public files still contain only numeric style indexes. The directory is divided into 256 hash buckets; commits replace touched buckets, metadata only when changed, and touched pages. History retains references to page versions, up to 100 entries or a conservative 32 MiB delta budget. Reopen reconstructs references and reclaims orphan staging.

## Import, bounds and cost

JSON object, string, Blob and transferable byte stream imports are supported. Blob/stream parsing emits complete pages without `JSON.parse` of the entire document, permits metadata after blocks, and stages pages on disk. Binary import also processes one page at a time. Both validate metadata, IDs, styles, duplicates, unsafe keys and resource costs before publishing a single head. Staging, old committed data and history all count toward the 256 MiB default budget; replacing a large workbook can require room for both versions. Cancellation or failure before publication preserves the prior committed workbook. Cancellation cannot reverse an already committed import.

Metadata, JSON page parsing and decoded binary pages have 32 MiB per-section working limits; individual text values have a 32,768 UTF-16-unit limit. Whole native files have no arbitrary 20 MiB limit but remain subject to the storage budget. Decoded Worker cache reservation defaults to 64 MiB, viewport cache to 8 MiB. These are cache reservations, not total process heap limits. Full `toJSON()` deliberately retains the complete JSON object; `JSON.stringify` creates an additional complete string. Streams bound page working data only when their sink consumes incrementally; collecting into a Blob retains all output bytes. XLSX conversion through SheetJS materializes its document separately.

`tests/file-codec.test.ts`, `tests/indices.test.ts`, storage tests and browser download/reimport tests cover lossless data, corruption, indexing, atomicity, cancellation, history and cache invalidation. [Codec results](../benchmarks/storage-binary.json) measure payload sizes; [browser results](../benchmarks/browser-binary.json) additionally include real Worker, IndexedDB, history, rendering and caches. Neither proves superiority over other spreadsheet products.
