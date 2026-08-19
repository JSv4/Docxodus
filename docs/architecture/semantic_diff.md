# Semantic diff

`SemanticDiff` is the verification-facing comparison API for answering “what
meaningfully changed between these DOCX packages?” It is additive to the redline
surface: `DocxDiff.Compare`, `GetRevisions`, `GetEditScriptJson`, and the legacy
`DocxSession.GetDiff` retain their existing contracts.

The public result is a stable, versioned JSON document. Version 1 is named
`docxodus.semantic-changes` and reports each change with its owning package part,
structural path, left/right anchors and scopes when available, change family,
operation, and typed before/after values. The machine-readable contract is
[`semantic-changes-v1.schema.json`](../schemas/semantic-changes-v1.schema.json).

## Public APIs

```csharp
var changes = SemanticDiff.Compare(left, right);
var canonical = changes.ToCanonicalJson();

// Convenience aliases on the existing comparison facade.
var sameChanges = DocxDiff.GetSemanticChanges(left, right);
var json = DocxDiff.GetSemanticChangesJson(left, right);

// A session compares its opening package with its current logical checkpoint.
using var session = new DocxSession(bytes);
// ...mutate the session...
var sessionChanges = session.GetSemanticChanges();
```

`SemanticDiffOptions` carries the existing `DocxDiffSettings`, enables the package
supplement by default, and exposes its shared `PackageManifestOptions` as
`PackageOptions`. The semantic defaults are 10,000 ZIP entries, 64 MiB of
decompressed data per entry and per parsed XML part, 256 MiB of aggregate
decompressed data, a 1,000:1 maximum compression ratio, and 2,048 decoded
characters per canonical package URI. XML nesting is capped at 256 elements before
recursive normalization. Integer and byte bounds must be positive; the
ratio must also be finite. Entry
names reject backslashes, control characters, malformed escaping, traversal, and
ASCII-case-insensitive duplicates after canonical package-URI decoding. Relationship parts reject duplicate
relationship ids. Declared ZIP sizes provide early rejection, while per-entry and
aggregate byte budgets are also enforced against the actual decompressed streams.
The raw-byte overloads and byte-oriented bridges run manifest preflight before
constructing a `WmlDocument`, revision pre-acceptance, or Open XML SDK/IR parsing,
even when `IncludePackageChanges` is `false`; disabling the supplement does not
disable validation. The `WmlDocument` overload also preflights before semantic IR
parsing, but the caller has necessarily opened the bytes while constructing that object.

The session baseline is controlled by
`DocxSessionSettings.CaptureInitialProjection` (default `true`). Enabling it now
retains both the initial projection and an exact copy of the opening package, so
callers that disable it avoid both costs but cannot call `GetDiff` or
`GetSemanticChanges`. Current bytes are made from an isolated logical checkpoint;
the live session is not saved or mutated merely to inspect changes.

Equivalent surfaces are available as:

| Surface | Stateless | Session |
|---|---|---|
| WASM bridge | `DocxDiffBridge.GetSemanticChangesJson` | `DocxSessionBridge.GetSemanticChanges` |
| npm | `docxDiffGetSemanticChanges(left, right, settings?)` | `session.getSemanticChanges()` |
| npm Web Worker | `worker.getSemanticChanges(left, right, settings?)` | `workerSession.getSemanticChanges()` |
| Python | `docx_diff_get_semantic_changes(left, right, settings=None)` | `session.get_semantic_changes()` |
| MCP | — | `docxodus_get_content` with `format: "semantic_changes"` |

The byte-oriented transports accept the existing `DocxDiffSettings` comparison controls but keep
the package supplement enabled and enforce the documented default preflight policy. Call the .NET
API with `SemanticDiffOptions` when an application needs custom package limits or explicitly wants
to omit the package-level supplement.

The MCP form is document-wide and rejects `anchorId` instead of silently
returning a partial schema.

## Version 1 schema

The compact canonical form has a fixed field order. This example abbreviates the
typed values but shows every change-level field:

```json
{"schema":"docxodus.semantic-changes","schemaVersion":1,"changeCount":1,"changes":[{"id":"chg-000001","operation":"modify","family":"run_formatting","partUri":"/word/document.xml","path":"paragraph.run[0].format","leftAnchor":"p:body:abc","rightAnchor":"p:body:abc","leftScope":"body","rightScope":"body","moveId":null,"before":{"kind":"object","value":{"bold":{"kind":"boolean","value":false}}},"after":{"kind":"object","value":{"bold":{"kind":"boolean","value":true}}}}]}
```

Operations are `insert`, `delete`, `move`, and `modify`. A move carries a stable
`moveId`; inserts and deletes use the explicit `absent` value on the missing side.
The v1 family vocabulary is:

- content and layout: `text`, `block_structure`, `run_formatting`,
  `paragraph_formatting`, `style`, `numbering`, `list`, `section`, and
  `page_setup`;
- tables: `table`, `table_row`, `table_cell`, `table_span`, `table_width`, and
  `table_style`;
- stories and review data: `header`, `footer`, `field`, `footnote`, `endnote`,
  `comment`, `revision`, and `annotation`;
- package-linked content: `hyperlink`, `bookmark`, `content_control`, `image`,
  `media`, `relationship`, and `opaque_package_part`.

Before/after data is a closed union: `absent`, `string`, `boolean`, `integer`,
`digest`, `object`, or `array`. Object members are ordinally sorted; arrays keep
their semantic order. Version 1 integers are restricted to the inclusive
ECMAScript-safe range −9,007,199,254,740,991 through 9,007,199,254,740,991 so a
canonical value crosses .NET, Python, and JavaScript JSON clients without rounding.
Several OOXML attributes — `wp:extent/@cx`, `w:gridCol/@w`, `w:bookmarkStart/@w:colFirst`,
and declared ZIP entry sizes — parse as unbounded 64-bit values, so a crafted package can
exceed that range. Such a value is projected as a lossless decimal `string` rather than an
`integer`: the range stays a hard guarantee for the `integer` kind, one value's kind degrades
instead of the comparison failing, and two distinct out-of-range values stay distinguishable.
Modeled state that is already 32-bit typed keeps the range check as an assertion.
Digest values separate the cryptographic algorithm from
the normalization domain, for example:

```json
{"kind":"digest","algorithm":"SHA-256","profile":"docxodus-ir-content-v1","value":"...lowercase hex..."}
```

Consumers must branch on `schema` and `schemaVersion`. New schema versions may
append families, but v1 names will not be repurposed.

## Determinism and suppression

The existing IR edit script remains the alignment and move-detection authority.
A projection layer turns it into semantic changes and compares typed IR
registries for formats, styles, numbering, tables, stories, fields, links,
content controls, images, notes, and comments. A package-level supplement covers
facts not represented by the IR: relationships, media bytes, bookmarks, native
revision/annotation records, and otherwise unknown parts.

`SemanticChangeSet` canonicalizes caller order with an ordinal sort over location,
family, path, operation, and typed values, then assigns `chg-000001`,
`chg-000002`, and so on. `ToCanonicalJson()` and `ToCanonicalUtf8Bytes()` are the
forms to hash or sign. `ToJson(indented: true)` is display-oriented.

The package supplement deliberately suppresses representation-only differences:

- ZIP order, timestamps, compression, XML declaration/BOM, namespace prefixes,
  attribute order, and insignificant whitespace in known Word-owned metadata
  parts do not create changes;
- relationship ids are not identity, and owner-relative and package-absolute
  internal relationship targets resolve to one canonical part URI. Relationship
  references are also compared at their owning XML locations: a coordinated `rId`
  rewrite is suppressed, but swapping the semantic targets of two references is
  reported even when the relationship inventory is otherwise unchanged;
- unknown XML is never dropped. Its expanded-name fingerprint preserves text and
  whitespace nodes, comments, processing instructions, and top-level nodes because
  vendor/custom XML may assign meaning to any of them;
- styles, numbering, and theme fonts have typed records for the commonly modeled
  values plus normalized full-part residual records, so defaults, conditional
  styles, restart/suffix/legal numbering, picture bullets, run properties, theme
  colors, and non-Latin theme fonts cannot disappear through registry projection;
- binary unknown parts and media are represented by bounded size/content-type
  facts and SHA-256 digests rather than embedded payloads.

The package supplement consumes the same bounded pass as
`PackageManifestGenerator`: shared content-type resolution, relationship
enumeration, `ChangeLocation` diagnostics, raw digests, and
`XmlSemanticNormalizer` results. Its internal inspection view retains detached
parsed XML trees, but no archive handles or duplicate raw payloads, so there is no
second ZIP/XML reader with a divergent safety policy. Any error-severity manifest
finding fails closed before the SDK. Well-formed unknown/vendor content remains
visible as opaque data when it is safe to inspect.

## Performance guard

`SemanticDiffTests.Thousand_paragraph_single_edit_has_bounded_time_and_output`
builds two 1,000-paragraph documents with one text edit. The complementary
`Dense_table_document_single_edit_has_bounded_time_and_output` test builds 200
tables (400 rows) with one cell edit. Both exercise the default package supplement,
require completion under 30 seconds, and cap canonical output at 1,000,000 bytes.
On the 2026-08-15 final integration run, the paragraph guard reported 132.4 ms,
one change, and 589 canonical bytes; the dense-table guard reported 205.0 ms,
eight changes, and 6,663 canonical bytes.
The package supplement bounds entry count, URI length, compression ratio, and both
per-entry and aggregate actual decompressed bytes; it does not impose a whole-result
truncation policy. Callers comparing extremely change-dense documents should stream
or persist the returned canonical bytes according to their own application limits.

## Compatibility

Semantic diff is a new API and JSON namespace. It does not change the existing
edit-script JSON, tracked-revision list, produced redline, or projection-based
`DocxSession.GetDiff` shapes. This separation lets verification consumers adopt a
durable audit schema without turning the renderer's internal edit script into a
permanent cross-version contract.
