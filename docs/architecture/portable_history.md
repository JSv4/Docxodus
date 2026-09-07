# Portable document history

The host owns everyday persistence. Users open a document, save checkpoints, browse
versions, compare any pair, restore, and download files without managing store directories.

Two explicit exports prevent accidental draft disclosure:

- `.docx`: one exact selected snapshot; no external history is embedded.
- `.docxhistory`: one captured document head and its complete referenced history.

The archive is a versioned ZIP, not a dump of the host's storage directory. Typed
references select snapshots, versions, package effects, retry receipts, original
publication states, and backend decisions/proposals. Unrelated documents and
unreferenced blobs are excluded. IDs and snapshot bytes survive unchanged.

Opening an archive is read-only and validates the bounded graph. Import validates
before publication, writes immutable blobs, then atomically initializes an absent
head with its **original revision**. Existing different histories are never
overwritten. Identical imports are idempotent. Import requires an additive
absent-only initialization capability; ordinary head CAS semantics do not change.

Implementation/review groups:

1. Bounded typed graph traversal and validation.
2. Stream archive export and read-only opening; strict format and hostile-file tests.
3. Safe writable import and document-scoped native API.
4. Client bindings and a concise frontend controls guide.
5. Real DOCX/file lifecycle and failure/recovery verification.

Each group gets a separate stacked PR and adversarial 5.6-sol review/revision.
Tests must exercise real legal documents, actual archive and snapshot files,
reopening with fresh stores, arbitrary comparison, restore, continued saves,
receipt retries, unrelated-data exclusion, malformed input, concurrent changes,
and failures before/after publication. Test artifacts and reproduction commands
will be documented; transient unit-test files alone are not the completion gate.

No GUI, transport, background recorder, implicit checkpoint, distributed filesystem
guarantee, or automatic retention policy is added. Archive integrity is not
authentication: the host still owns access control and trust/provenance policy.

## Graph validation

The internal graph reader accepts a pinned head, not a mutable head store. It verifies
every typed record once, checks each incoming edge, and streams payload hash checks.
Radix subtree receipt counts and maximum revisions prove receipt retention without
re-expanding shared trees per publication. Referenced publication heads must be on the
captured chain; a restore target may retain another same-document version branch.

Effects are checked against exact endpoint entry names/digests, without materializing
changed payloads or constructing a ZIP from untrusted effects. Original text proposals
and accepted maps are checked separately. Snapshot buffers are transient; graph metadata,
blobs, edges, graph-blob validation reads, and conservative package-expansion work have distinct
limits. Expansion work is charged with conservative multipliers, so the allowance is
not a promise that an archive of that size will open. Package-level ZIP/XML limits also
apply. This proves bounded structural/content consistency, not authorship or trust in
host timestamps and metadata.

Archive validation is **not a backend decision-policy audit**. It retains acceptance
statuses, conflict reasons, and chosen text-map offsets; it does not re-adjudicate
those choices by rerunning reconciliation. A verified contribution proves its recorded
before/after content, not that accepting that contribution was the right response to
an operation proposal. The authoritative backend owns that decision. Opening a history
file must not reinterpret recorded outcomes using a newer conflict policy.

## Archive v1

The container is a narrow [ZIP/ZIP64 profile](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT):
one disk, stored/deflate entries, no encryption, directories, links, or ZIP comments.
Entry names are exactly `history.json` and `blobs/<lowercase-sha256>`; extra fields are
bounded. ZIP metadata uses these container/count/manifest limits, not the graph-blob
`MaxValidationBytes` allowance. The central directory is checked before allocating ZIP entry objects. No entry
is extracted to a filesystem path. All inventory entries must be reachable, and all
reachable entries must be present.

`history.json` has five required fields and rejects unknown/duplicate properties:

```json
{
  "schema": "https://docxodus.dev/schemas/history/archive/v1",
  "schemaVersion": 1,
  "documentId": "host-owned-stable-id",
  "head": { "revision": "8", "state": { "sha256": "<64 lowercase hex>", "length": 1234 } },
  "blobs": [{ "sha256": "<64 lowercase hex>", "length": 1234 }]
}
```

Inventory is sorted by digest, with no duplicates. Revision is a canonical positive
Int64 decimal string; lengths are nonnegative Int32 numbers. Blob contents use their
existing versioned codecs, copied verbatim. Outer ZIP compression/serialization is not
an identity guarantee; raw blob bytes, snapshot bytes, version IDs, and the captured
head are. Reopening does not need the original store or any sidecar files.

## Writable import and document facade

`DocxVersionHistory.Document(id)` binds the existing read surface plus identified
checkpoint/restore methods. It owns no storage or editor and performs no implicit writes.
`ImportHistoryArchiveAsync` first validates the isolated archive with the destination's
snapshot/record/package limits. Unsupported head adapters fail before validation/writes;
a different existing head fails before blob copying. Every blob is then put through the
immutable store contract, including exact-head retries, so missing blobs are repaired
and corrupt collisions cannot masquerade as successful import.

The archive reader and entry leases are closed before publication; caller-owned input
stays open, while owned input is closed. One `IHistoryHeadInitializer` call either
inserts the captured positive revision into an absent slot or returns the existing head.
It shares exclusion with ordinary CAS; FileHistoryHeadStore uses its existing exclusive
lock and atomic rename, with replacement disabled for initialization. Ordinary CAS still
increments by one. Different heads always conflict; exact equality is idempotent. No
post-commit reads/cancellation can obscure this call's success. I/O can lose an acknowledgement:
retry the same archive. If another writer has since advanced it, the retry reports conflict
rather than rolling history back. Document IDs cannot be remapped because records bind them.

## Client boundary

The existing generated-JSON `HistoryClientOps` owns readonly capability checks and archive
calls for WASM, Python and MCP. Standalone archive handles own their isolated reader; scoped
documents borrow their writable client's lifetime. Closing releases handles, never host data.
Byte-oriented clients cap files at 64 MiB, use copied/base64 buffers, and leave larger-file
streaming to native hosts. Existing browser adapters remain valid; only adapters with optional
`initializeHead` opt into exact-head import. The memory adapter shares the same synchronous
critical section with ordinary CAS. WASM verifies length/hash before passing copied blobs to
the host callback. No storage callbacks are installed for readonly archive opening.

MCP uses its existing session capability and canonical location. The embedded archive ID
must match before destination reads/writes, so moving a history to a differently named MCP
document is not supported. Import does not change an open session or source file. Shared
byte operations have no new transport; host permissions, durability and retention remain
outside the library. [Frontend controls](../history-controls.md) intentionally document the
small application-facing API rather than restating the underlying storage graph.
