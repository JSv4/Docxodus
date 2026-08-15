# Deterministic package manifests

> **Status:** Implemented in `Docxodus/Verification`. Schema identifier:
> `https://docxodus.dev/schemas/verification/package-manifest/v1`.

`PackageManifestGenerator` inspects DOCX/OPC bytes without opening a mutable
`WordprocessingDocument`. It inventories the complete ZIP package, resolves OPC metadata,
computes three deliberately different package identities, extracts high-signal document facts,
and reports malformed or unsupported structures as stable findings. It is intended for
verification, provenance, caching, and deciding whether two files differ only in packaging or XML
serialization.

This is not an XML-signature implementation and `isValid` is not a claim that Microsoft Word will
render the file. It means that the structural and safety checks documented here produced no
error-severity finding.

## Entry points

| Surface | Stateless bytes | Current live session |
|---|---|---|
| .NET | `PackageManifestGenerator.Generate(bytes)` / `GenerateJson(bytes)` | `session.GetPackageManifest()` |
| WASM export | `DocumentConverter.GeneratePackageManifest(bytes)` | `DocxSessionBridge.GetPackageManifest(handle)` |
| npm | `await generatePackageManifest(fileOrBytes)` | `session.getPackageManifest()` |
| npm worker | `await worker.generatePackageManifest(fileOrBytes)` | — |
| Python | `generate_package_manifest(bytes)` | `session.get_package_manifest()` |
| MCP | — | `docxodus_get_content({ sessionId, format: "manifest" })` |

The shared `VerificationOps` facade owns canonical JSON for all transports. MCP manifests always
cover the complete package and reject `anchorId`; a package hash scoped to one paragraph would
have a misleading meaning.

The stateless operation hashes the exact supplied bytes. The session operation first creates the
same isolated logical checkpoint used by transaction snapshots: it clones the package and overlays
dirty cached XML on the clone. Unsaved edits are therefore represented, while the live package,
cache objects, undo/redo history, and document version are unchanged.

**All three session digests describe the checkpoint, not the opened file.** Serializing a package
through the SDK rewrites XML and repacks the ZIP, so a session manifest taken immediately after
opening a document normally differs from the stateless manifest of the same bytes in the raw,
ordered *and* normalized-semantic digest. Compare session manifests with session manifests and
stateless manifests with stateless manifests; comparing across the two entry points measures the
round-trip, not the edits.

## Schema-v1 envelope

The top-level fields are emitted in this fixed order:

1. `schema`, `schemaVersion`, `packageKind`, `isValid`
2. `rawPackageBytesDigest`, `orderedOpcContentDigest`, `normalizedSemanticDigest`
3. `entries`, `contentTypes`, `relationships`, `facts`, `findings`

Every digest is `{ "algorithm": "SHA-256", "value": "<lower-case hex>" }`. A digest is `null`
when safety limits, encryption, or unreadable content prevent computing it; absence is never
represented by an empty digest. In particular, breaching the entry-count or total-expansion limit
suppresses both content digests: an inspection that stopped early has not seen the whole package,
and two packages differing only past the cut must not compare equal.

`packageKind` is one of `opc`, `zip`, `zip-encrypted`, `ole-encrypted`, `ole`, or `malformed`.
Password-encrypted OOXML commonly appears as an OLE compound file containing `EncryptedPackage`
and `EncryptionInfo`; it is identified without attempting decryption. Traditional ZIP encryption
is detected from central-directory flags. Both are unsupported and produce error findings.

### Entries and content types

`entries` contains one record for every inspected physical central-directory entry, including
duplicate names. Each record contains:

- a canonical leading-slash URI and a stable duplicate `occurrence`;
- resolved content type plus source (`override`, `default`, `implicit`, or `unresolved`);
- declared uncompressed and compressed sizes;
- the exact uncompressed-byte digest and, for readable XML, normalized XML digest;
- `isXml` and tri-state `isEncrypted` (`null` when the central-directory flags cannot be
  established authoritatively).

`[Content_Types].xml` and relationship parts receive their package-defined implicit MIME types.
A normal part without a matching Default or Override retains `contentType: null` and produces
`missing_content_type`; the manifest does not invent a MIME type. `missing_content_type` is only
emitted when `[Content_Types].xml` was actually parsed — when the declaration file itself is
absent, malformed, oversize, or unreadable, the manifest says so once
(`missing_content_types` / `malformed_content_types` / `content_types_unreadable`) instead of
blaming every part in the package. `contentTypes` preserves every Default/Override declaration,
including duplicate occurrences, while separately reporting duplicate keys, conflicting values,
and Override targets that do not exist.

**Directory-only ZIP entries.** OPC packages should not carry them, but 7-Zip, Windows'
*Send to → Compressed folder*, several Java/PHP zip writers, and some Word templates do, and Word
opens those files. They are inventoried with a trailing-slash URI (`/word/`), reported as a
`directory_entry` **warning** rather than an error, exempt from content-type resolution, and
excluded from both content digests — adding or dropping folder entries is repackaging, not a
document change.

### Relationships

`relationships` includes every readable package-level and part-level Relationship with owner URI,
Id, type, raw target, normalized target mode, resolved internal target URI, and target presence.
`/` is the package owner. Targets may be relative to the owning part or package-absolute
(`/word/document.xml`, the form the Open XML SDK writes); neither can escape the package root.
A target is external only when it carries an RFC 3986 scheme, never merely because it starts with
a slash. External targets are retained but never dereferenced.

The generator distinguishes:

- `missing_target`: a declared internal Relationship resolves to a package part that is absent;
- `dangling_relationship`: an XML `r:id`/`r:embed`/`r:link`-style reference has no Relationship in
  its owning part;
- `relationship_part_unreadable`: the `.rels` part exists but was never parsed, so its
  relationships are unknown. `dangling_relationship` is suppressed for that owner — every
  reference would otherwise look dangling because of one unread file;
- duplicate or conflicting Relationship IDs within one owner;
- a relationship part whose owner is absent or whose URI cannot identify an owner.

### Facts

`facts` records the main-document URI, Strict/macro flags, document-property presence, structural
counts, header/footer/note/style/numbering/theme/media/custom-XML counts, drawing/altChunk/field
counts, tracked-revision-family counts, and Word comment/threading plus Docxodus-annotation counts.
Revision `total` is exactly the sum of insertions, deletions, move-from, move-to, property,
structural (`cellIns`/`cellDel`/`cellMerge`), and other (custom-XML revision-range) changes.
These are deterministic package facts, not layout estimates. Paragraph and table counts span the
readable WordprocessingML stories in the package; no page count is inferred.

## The three package identities

The three root digests answer different questions:

| Digest | Input | Ignores | Typical use |
|---|---|---|---|
| `rawPackageBytesDigest` | Exact caller-supplied byte array | Nothing | Byte provenance, transfer integrity |
| `orderedOpcContentDigest` | URI-ordered identities plus each exact uncompressed-entry SHA-256 | ZIP entry order, timestamps, compression, central-directory layout | Detect a pure ZIP repack |
| `normalizedSemanticDigest` | URI, occurrence, content type, normalized XML digest or raw binary digest | The XML serialization choices below, plus ZIP packaging | Detect a serialization-only XML rewrite |

All three use SHA-256. For the ordered OPC digest, each entry contributes canonical URI,
occurrence, actual decompressed size, and the lower-case hexadecimal exact-byte digest. Strings in
the aggregate streams are UTF-8 prefixed by a little-endian 32-bit byte length; occurrences use
little-endian 32-bit integers and entry sizes use little-endian 64-bit integers. Entries sort by
canonical URI using ordinal comparison, then by occurrence. Using the already-computed per-entry
digest avoids a second decompression pass over untrusted input.
Duplicates receive occurrences by raw digest, size, then original archive index, so archive order
does not perturb otherwise distinguishable duplicates. Directory-only entries contribute to
neither digest.

Each semantic-digest entry is tagged with the identity it contributed: `X` a normalized XML
digest, `B` opaque binary bytes, `U` an entry whose declared content type says XML but which
could not be parsed. A `U` entry falls back to its exact bytes, so one unparsable part costs that
part its serialization independence rather than costing the whole package its identity. The tag is
part of the hashed stream, so a part cannot silently move between the three states.

Interpret a comparison as follows:

| Raw | Ordered OPC | Normalized semantic | Interpretation |
|---|---|---|---|
| same | same | same | Byte-identical package |
| different | same | same | Repacked ZIP only |
| different | different | same | XML serialization-only change covered by schema v1 |
| different | different | different | Semantic XML, binary, URI, or content-type change |

A null lower-level digest means the comparison is unavailable, not unequal.

## Normative XML normalization

Schema v1 parses with DTD processing prohibited, no external resolver, preserved comments,
processing instructions, and whitespace, and a configured character ceiling. Its digest token
stream applies these rules:

- XML declarations, BOM/encoding choice, namespace-prefix spelling *in element and attribute
  names*, and namespace-declaration placement are ignored. A prefix appearing inside an attribute
  *value* — `mc:Ignorable="w14"`, `xsi:type` and other QName-valued attributes — is hashed as
  written, so renaming such a prefix reports a semantic difference. Schema v1 does not interpret
  attribute values, and treating an unknown value as a QName would be a guess; this is a known
  conservative limitation, not a claim that the two documents differ.
- Element and attribute names use expanded `{namespace URI, local name}` identity. Non-namespace
  attributes sort ordinally by namespace URI, local name, then value.
- Attribute order, quote style, entity spelling, empty-element spelling, and CDATA-versus-text
  spelling are ignored. Adjacent text nodes are coalesced. XML-mandated CR/CRLF normalization is
  inherited from the parser.
- Element order, attribute values, text characters, comments, and processing instructions remain
  semantic by default.
- For an explicit schema-v1 allowlist of known OOXML/OPC MIME types only, whitespace-only text
  between child elements is treated as indentation and ignored. Vendor-looking MIME prefixes are
  not sufficient. `xml:space="preserve"` disables that rule for its subtree;
  `xml:space="default"` restores the default behavior.
- Application/opaque XML, including unknown extensions and `customXml` data with generic
  `application/xml`, preserves whitespace-only text. This includes an unrecognized
  `application/vnd.openxmlformats-*+xml` or `application/vnd.ms-*+xml` value. Docxodus does not
  guess an unknown schema's element-content model.
- At the document root of `[Content_Types].xml`, Default/Override elements sort by their complete
  declaration identity. At the root of `.rels` parts, Relationship elements sort by
  Id/Type/Target/TargetMode. This sorting occurs only when the other root children are formatting
  whitespace; comments, processing instructions, or non-whitespace text retain the original node
  order and therefore remain semantic.

Strict OOXML namespace URIs are deliberately **not** rewritten to Transitional namespace URIs.
Both forms parse and contribute the same categories of facts, but conversion between conformance
classes changes the normalized semantic digest. Schema v1 therefore never asserts that a Strict
package and a Transitional package are equivalent merely because their visible text matches.

The normalized digest is intentionally conservative. It does not collapse WordprocessingML run
boundaries, style cascades, field results, markup-compatibility choices, image encodings, or other
constructs that could render alike. Such equivalence belongs to a future renderer-aware artifact,
not this package manifest.

## Deterministic ordering and JSON

`PackageManifest.ToJson()` is the canonical compact representation. In addition to fixed property
order, collections sort as follows:

- entries: URI, occurrence;
- content-type declarations: kind, key, content type, occurrence;
- relationships: owner, Id, type, target, target mode;
- findings: severity (`error`, `warning`, `info`), code, location fields, message.

All string comparisons are ordinal unless a package identifier is explicitly case-insensitive in
OPC resolution. `ToJson(indented: true)` is for display and is not the canonical byte envelope.

## Safety limits and findings

`PackageManifestOptions` defaults are:

| Limit | Default |
|---|---:|
| central-directory entries | 10,000 |
| total declared uncompressed bytes | 1 GiB |
| XML part bytes parsed | 32 MiB |
| per-entry expansion ratio | 1,000:1 |
| decoded package URI characters | 2,048 |

Raw OPC segment validation rejects malformed escapes, encoded slash/backslash or unreserved
characters, empty/dot/trailing-dot segments, and paths that escape the package root before any
percent decoding can hide them. Both declared and actual decompressed bytes are bounded: every
entry has a saturating `compressedSize × MaxCompressionRatio` ceiling and all reads share the
package budget. ZIP traversal/absolute/backslash paths, count/size/ratio breaches, DTDs, encrypted entries,
unreadable payloads, duplicate/conflicting metadata, and relationship faults are returned in
`findings`. Each finding has a stable snake-case code, `info`/`warning`/`error` severity, a human
message, and a reusable `ChangeLocation` (entry, owner, relationship, target, or property path).
Consumers should branch on `code`, not message text.

Malformed and encrypted inputs return a manifest with their exact raw digest and structured
findings instead of throwing for normal validation failures. Programmer errors such as null input
or invalid option values still throw. The generator never extracts files to disk and never follows
external Relationship targets.

Encryption flags are read from classic or ZIP64 central-directory metadata. If that metadata
cannot be parsed authoritatively, the manifest emits `zip_encryption_detection_unavailable`, sets
each affected entry's `isEncrypted` to `null`, and does not read or hash those entry payloads.
This detection is deliberately stricter than `System.IO.Compression.ZipArchive`: it requires the
end-of-central-directory record to terminate the file exactly and its entry count to match, so a
ZIP carrying a prepended stub, an appended signature, or a miscounted comment fails closed. Such a
package still yields its raw digest, entry inventory, and findings, but no content digest and no
resolved content types — deciding that an unreadable central directory contains no encrypted
entries is a claim the manifest will not make.

The entry-count limit truncates inspection at `MaxEntryCount`. Declared expansion is nonetheless
summed over the whole central directory, so a package cannot dodge the size budget by also
breaching the entry-count limit.

## Examples

```csharp
var manifest = PackageManifestGenerator.Generate(docxBytes);
if (!manifest.IsValid)
    foreach (var finding in manifest.Findings)
        Console.WriteLine($"{finding.Code}: {finding.Message}");

using var session = new DocxSession(docxBytes);
var afterEdits = session.GetPackageManifest();
```

```typescript
const before = await generatePackageManifest(bytes);
const current = session.getPackageManifest();
console.log(before.normalizedSemanticDigest?.value);
```

```python
before = generate_package_manifest(docx_bytes)
with open_session(docx_bytes) as session:
    current = session.get_package_manifest()
```

## Known limits of schema v1

Recorded so a reader does not mistake an accepted trade-off for an oversight.

- **`PackageManifestOptions` is a .NET-only surface.** WASM, npm, the stdio Python host, and MCP
  call the parameterless overloads and therefore always use the defaults above. A browser or
  Python caller cannot raise `MaxXmlPartBytes` for a legitimately large package; it receives the
  degraded manifest instead. Plumbing the options object through the four transports is a
  separate change.
- **Parsed XML is retained for the whole generation.** Every readable XML part stays materialized
  as an `XDocument` until the manifest is returned, so the documented limits bound bytes read but
  not the larger in-memory DOM. The defaults are sized for documents, not for adversarial archives
  of maximum-size XML parts; lower `MaxTotalUncompressedBytes` and `MaxXmlPartBytes` when
  inspecting untrusted input in a memory-constrained host.
- **The WASM export accepts any buffer size.** `DocumentConverter.GeneratePackageManifest` skips
  the 100 MB `MaxDocumentSizeBytes` guard its sibling exports apply, because triaging an
  oversized or malformed package is the point of the operation. The manifest's own limits still
  apply, but a browser caller is responsible for not marshaling a buffer its tab cannot hold.
- **Manifest JSON grows with the package.** It is roughly 10–40 KB for a typical document and
  scales with entry, relationship and finding counts. `docxodus_get_content(format: "manifest")`
  returns the whole artifact with no summary mode, so an agent budgeting context should read
  `isValid` and `findings` rather than the full envelope.
