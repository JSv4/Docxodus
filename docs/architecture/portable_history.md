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
blobs, edges, raw validation reads, and conservative package-expansion work have distinct
limits. Expansion work is charged with conservative multipliers, so the allowance is
not a promise that an archive of that size will open. Package-level ZIP/XML limits also
apply. This proves bounded structural/content consistency, not authorship or trust in
host timestamps and metadata.
