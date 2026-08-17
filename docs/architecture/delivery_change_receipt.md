# Delivery change receipt

Issue #458 composes evidence that already belongs to the document transaction, package,
semantic-diff, rendering, and verification layers. It does not create a second mutation
engine or reinterpret those components' schemas.

## Contract and trust boundary

A delivery receipt connects these immutable inputs:

- the source and delivered DOCX package identities from the package manifest;
- one or more committed mutation-batch results and their normalized requests;
- exactly one typed source-to-delivered semantic change set and one for every state-changing
  transaction;
- exactly one available clean DOCX whose bytes are the delivered package;
- optional validation and redline-reversibility results;
- optional review DOCX, HTML, PDF, image, and report artifacts;
- optional page citations produced for a specific document version and render fingerprint.

The exact writer contract is published as
[delivery-change-receipt-v1.schema.json](../schemas/delivery-change-receipt-v1.schema.json),
with schema identifier
`https://docxodus.dev/schemas/verification/delivery-change-receipt/v1`. The schema is the
closed projection emitted by the v1 writer: nullable members are present as JSON `null` and
collections are present as arrays, including when empty. A receipt is an envelope containing a
canonical `payload` and a `receiptDigest`. The digest is SHA-256 over the canonical UTF-8 JSON
bytes of `payload`; it never includes itself. Every available artifact record carries its own
SHA-256 and byte length. Verification first checks the receipt digest, then independently hashes
each available artifact.

The receipt proves what the supplied evidence says and whether referenced bytes still
match. It is not a signature or an assertion about who controlled a transaction ID.
Callers that need authorship or non-repudiation sign the canonical receipt bytes outside
this schema.

## Deterministic payload

Canonical serialization uses UTF-8 without a byte-order mark, camel-case property names, JSON
strings for enums, and no insignificant whitespace. String escaping is exactly
`Utf8JsonWriter` with `JavaScriptEncoder.Default`, not `UnsafeRelaxedJsonEscaping`: JSON quote,
backslash, and control-character escapes use the writer's fixed spellings; HTML-sensitive Basic
Latin characters are escaped; and non-Basic-Latin scalar values are emitted as uppercase
`\uXXXX` UTF-16 code-unit escapes, using a surrogate pair for a supplementary scalar. `/` is not
escaped. These escape spellings, including escape-letter and hexadecimal case, are hash-significant.

Every JSON object's properties are sorted with `string.CompareOrdinal`. This is a comparison of
UTF-16 code units with no culture, case folding, or Unicode normalization; it is not UTF-8 byte or
Unicode scalar-value order. Duplicate property names are rejected. Map-shaped inputs are emitted
as arrays sorted by their stable key. Sets, findings, package parts, relationships, artifacts, and
warnings have defined ordinal sort keys. Operation, transaction, and lineage-event arrays retain
their meaningful execution order. Transactions and undo/redo events also carry one contiguous
shared `sequence`, so chronology remains unambiguous when an edit follows an undo.

Core receipt numbers use the exact invariant token emitted by the typed `System.Text.Json` writer.
When raw JSON is verified, every known numeric field must have the same raw token as that typed
projection: numerically equivalent alternatives such as `1.0` or `1e0` do not match core `1`.
Conversely, a numeric value in an unknown `fullEvidence` extension is grammar-validated and then
written with its original `JsonElement.GetRawText()` token; it is never rounded through `double`.
Different extension lexemes therefore produce different receipt digests even when they denote the
same mathematical value. All typed non-negative 64-bit receipt integers (versions, sequences,
artifact lengths, and citation versions) are restricted to the JavaScript-portable interval
0 through 9007199254740991 inclusive. Typed 32-bit counts and indexes retain their narrower CLR
range. The portable-integer rule does not reinterpret arbitrary extension-number tokens.

Portable verification enforces the builder's order, not merely the enclosing payload digest.
Rehashing a payload after reordering or duplicating artifacts, package changes, citations,
evidence, warnings, authored changes, object changes, affected-anchor sets, transactions, or
lineage events therefore remains invalid.

Wall-clock receipt-generation time, local paths, ZIP timestamps, random identifiers, and
renderer process details are excluded. A caller-provided transaction ID is retained and included
in the receipt entry identity; entries without one use their deterministic receipt sequence. The
remaining identity inputs are the request fingerprint, base/result document versions, and
before/after package identities. Replaying an already committed identified transaction contributes
the same entry rather than a second edit, while distinct identified or anonymous no-op
transactions remain distinct entries.

## Transaction evidence

Every `DocxSession.ExecuteBatch` result that belongs to the delivery must contribute one entry.
The entry records:

- transaction ID and canonical request fingerprint when the transport supplied them;
- atomic or best-effort mode, base version, result version, and outcome;
- normalized operations in step order, including tool, action, and normalized arguments;
- resolved anchors, scopes, created/modified/removed objects, and structured errors;
- revision, comment, and annotation changes with their recorded authorship; revision evidence also
  retains the owner record's `dateUtc`, family, part URI, scope, primary anchor, resolution status,
  privacy-aware diagnostic, constituent IDs, and constituent keys;
- before/after package identities and the semantic changes attributable to the entry;
- the batch result's optional reported package-content digest as evidence distinct from the
  independently manifested before/after package identities;
- warnings emitted by the transaction and evidence collectors.

Operation arguments are data, never embedded JSON fragments. Normalization rejects
duplicate properties and uses the same operation vocabulary as the executing surface.
Transport-only fields such as storage paths and preview controls are not document
operations and are omitted.

An atomic failure contributes diagnostic evidence but is not labeled committed. Its complete
normalized request remains authoritative even though the core result is sparse: a preflight
failure marks every other operation `notExecuted`, while an execution failure marks the successful
prefix `succeededRolledBack`, the failing operation `failedRolledBack`, and the unvisited suffix
`notExecuted`. Best-effort results cover every request and never use rolled-back states. The
verifier recomputes operation success from every nested result and checks execution/rollback state
against the transaction outcome. A best-effort transaction is partially committed only when it
retained at least one successful step and at least one failed step. Isolated previews may be
recorded as predictions, but they are distinct from committed entries. A prediction must start at
the current receipt state. Its after identity may equal its before identity, or it may report a
version exactly one greater when at least one operation succeeded. Either way, prediction output is
not registered as reachable, does not become current, does not enter the applied stack, and does
not clear or populate the redo stack; the next real transition must still start from the
pre-prediction current identity. Because `MutationBatchStepResult.Success` is the conjunction of
its results, an executed step may legitimately return an empty result list: it is retained as a
successful no-op operation rather than being rewritten as an internal failure.

Undo and redo do not masquerade as new user-requested changes. They are ordered lineage
events that reference the affected committed entry and record their before/after document
versions and package identities. The builder and verifier run the same deterministic state
machine: applied and redo histories are LIFO, every state-changing transaction or lineage event
advances the document version by exactly one, and a new state-changing transaction clears redo.
A committed no-op neither enters applied history nor clears redo. Undo followed by redo therefore
preserves one transaction entry plus two reversible state transitions. Repeated undo, redo before
undo, and undo of an older entry while a newer entry is applied all fail closed. A retry with an
identical transaction ID, request fingerprint, and result evidence references the original entry;
a conflicting fingerprint or result fails before receipt composition.

## Change attribution

Receipt package-change records use one of three dispositions:

- `userRequested`: directly covered by a normalized requested operation;
- `derived`: required fallout whose derivation is known, such as a relationship or content
  type added for requested media;
- `unexpected`: observed in the package delta but not covered by requested or declared
  derived-change rules.

Attribution never hides evidence. The complete semantic change set remains referenced and every
changed manifest entry remains recorded even when each package change is expected. Attribution
rules apply to the package delta; semantic changes retain their owner-defined #457 classification
instead of receiving a second receipt-level disposition. Unknown parts and relationships default
to `unexpected`. Receipt success can be configured to fail on any unexpected package change
without altering the underlying evidence.

Every package-change record that names a transaction must name a committed or partially committed
transaction that remains on the final applied stack. If it also carries an operation index,
including for `derived`, that operation must have succeeded and must not have been rolled back. An
undone transaction cannot explain delivered bytes; a subsequent valid redo makes it applied again.
This is especially important for partially committed best-effort batches: a failed step remains
visible as evidence, but it cannot be used to explain bytes retained by a different successful
step. `derived` changes may reference their originating transaction and operation with an explicit
derivation.

## Privacy profiles

The default `hashAndSummary` profile includes stable locations, change kinds, counts, digests, and
bounded structural summaries, but excludes full paragraph, comment, and operation text. `hashOnly`
retains identities, dispositions, counts, and hashes with `DeliveryTextEvidence.summary` and
`.value` both null. `hashAndSummary` requires a non-textual summary and a null value;
`fullEvidence` requires both the summary and exact value. `characterCount` is .NET `string.Length`,
so it counts UTF-16 code units rather than Unicode scalar values or grapheme clusters; `digest` is
SHA-256 over the exact UTF-8 encoding of the value.

Revision diagnostic messages use `DeliveryTextEvidence`. Free-form derivations,
artifact-unavailable reasons, and supplied evidence summaries use the same profile boundary.
`hashOnly` stores only `sha256:<64 lower-case hex>`; `hashAndSummary` stores
`<field label>; <UTF-16 count> characters; sha256:<64 lower-case hex>`; and `fullEvidence` stores
the original text. Operation argument summaries are structural property counts and full argument,
result, authored, and package-record values occur only in `fullEvidence`. Semantic evidence remains
an exact, separately hashed artifact owned by the #457 schema; the receipt records only its root
schema/version/count and transition binding.

The selected profile is part of the canonical payload. Redaction is performed before
canonicalization and hashing; removing text from an already generated receipt invalidates its
digest. Artifact hashes remain available in every profile, so a verifier with authorized access to
the bytes can validate them independently. The JSON verifier authenticates the raw payload before
deserialization and then requires every known field, array position, value, and core numeric lexeme
to equal the typed projection. Unknown optional object properties may be retained only for
`fullEvidence`. They are rejected recursively for `hashOnly` and `hashAndSummary`, because an
otherwise ignored field could smuggle free text past redaction. The published schema describes the
closed output of the current writer; this `fullEvidence` reader allowance is the additive
compatibility rule, not permission for the v1 writer to invent fields.

## Page citations and render evidence

A page citation is accepted only with all of:

- a document version and raw package digest reachable through the validated transaction and
  undo/redo history;
- the anchor and scope resolved at that version;
- page/section/fragment coordinates;
- the complete renderer fingerprint and page-map fingerprint;
- the independently hashed raw PageMap artifact bytes that supplied those coordinates;
- the hash of the PDF or render report that carries the cited pagination.

Receipt construction rejects a citation whose version, package identity, renderer
fingerprint, or artifact reference does not match the supplied delivery evidence. A
continuous HTML projection is not promoted to a page citation. A `PageImage` artifact is also
excluded: the citation's render artifact must have role `Pdf` or `RenderReport`, even when an image
represents one physical page. Both construction and portable verification strictly parse the
referenced PageMap bytes, reject malformed UTF-8 and duplicate JSON properties, apply the same
portable geometry/order/story contract used by `DocxSession.RegisterPageMap`, and project the
citation again. The PageMap artifact digest remains over the renderer's original bytes; JSON
canonicalization is only an input-validity gate.

## Artifact records

Each artifact has a stable role, media type, byte length, SHA-256, availability status,
and optional document-version/render binding. Paths are optional display hints and must be
relative; identity never depends on them. Interior slash and backslash separators are normalized
to `/`, while POSIX roots, UNC/backslash roots, drive-qualified paths, empty segments, and dot or
parent segments are rejected identically on every host OS. Every receipt requires exactly one available
`CleanDocx`: its artifact digest, package digest, document version, and actual bytes must exactly
match `DeliveredDocument`. It also requires a `SemanticDiff` artifact containing the exact
`SemanticChangeSet.ToCanonicalUtf8Bytes()` output for the aggregate comparison and one typed
binding per state-changing transaction. Other roles include review DOCX, HTML, PDF, page map,
package manifest, validation result, reversibility proof, and render reports. An unavailable
artifact has an explicit reason and no digest or byte length; it cannot masquerade as an output
that was produced and verified.

The caller's manifest is not accepted as proof of the clean output. Every document identity carries
the package-manifest schema and the manifest's `MainDocumentUri`, in addition to its raw and
optional content/semantic digests. Construction and verification both run the corrected bounded
package-manifest generator over the exact supplied clean bytes, require a valid OPC package with a
WordprocessingML main document, and compare the recomputed package kind, schema, main-document URI,
raw digest, ordered OPC-content digest, and normalized-semantic digest with `DeliveredDocument`.
The recomputed clean manifest's entries and relationships—not a same-identity caller inventory—are
also the sole source for delivered package-change observations. Semantic JSON receives
the analogous treatment: the complete closed #457 type graph is reconstructed and accepted only
when serializing it again produces byte-for-byte identical canonical evidence. Whitespace,
property-order variants, partial objects, null nested records, and merely rehashed substitutes fail.

The source side has a narrower trust boundary in v1: receipt construction receives a source
manifest but not the exact source DOCX bytes, so it cannot independently reparse that inventory.
The adapter does require the supported manifest schema/version, `IsValid`, package kind `opc`, and
a nonblank Word main-document URI before creating `SourceDocument`; those are caller-attested
manifest facts, not a reparse of source bytes. Source entries and relationships are therefore
trusted caller input and affect the reported deltas even when the manifest carries a genuine
raw-package digest identity. Consumers must not treat that digest or main-document URI alone as
authentication of the supplied source inventory. A future exact-source artifact registration can
make source and delivered inventory verification symmetric without changing this v1 behavior.

The receipt envelope is stored beside those outputs and is verified by hashing its canonical
payload. Artifact records are covered by that digest, and their bytes are covered by their
individual digests. Thus changing either a record or a referenced artifact is detectable.

## .NET composition and verification

`DeliveryChangeReceiptBuilder` composes caller-supplied source/delivered manifests, every
`MutationBatchResult` that belongs to the delivery, normalized operation arguments, and any
artifacts or external evidence. It does not mutate or render a document; it does independently
re-inspect the clean DOCX and semantic artifacts at their trust boundaries. `Build()` validates
the document lineage and returns an immutable envelope:

```csharp
var builder = new DeliveryChangeReceiptBuilder(sourceManifest, sourceVersion)
    .SetDeliveredDocument(deliveredManifest, deliveredVersion);

string entryId = builder.AddTransaction(
    DeliveryTransactionContribution.FromMutationBatchResult(
        batchResult, beforeManifest, afterManifest, normalizedOperations, transactionIdentity));
builder.AddArtifact(DeliveryArtifactInput.Available(
    "clean-docx", DeliveryArtifactRole.CleanDocx, docxMediaType, deliveredBytes) with
{
    Document = DeliveryDocumentIdentity.FromManifest(deliveredManifest, deliveredVersion),
});
SemanticChangeSet sourceToDelivered = SemanticDiff.Compare(sourceDocument, deliveredDocument);
SemanticChangeSet transactionChanges = SemanticDiff.Compare(beforeDocument, afterDocument);
builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
    sourceToDelivered));
builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForTransaction(
    entryId,
    transactionChanges,
    "semantic-transaction-1"));

DeliveryChangeReceipt receipt = builder.Build();
File.WriteAllBytes("delivery/change-receipt.json", receipt.ToJsonBytes(indented: true));
```

The portable verifier accepts either that object or its JSON bytes plus an artifact-id-to-bytes
map. It verifies the raw receipt payload first and then hashes every artifact independently:

```csharp
DeliveryReceiptVerificationResult result = DeliveryChangeReceiptVerifier.VerifyJson(
    receiptBytes,
    new Dictionary<string, byte[]>
    {
        ["clean-docx"] = deliveredBytes,
        ["semantic-source-to-delivered"] = sourceToDelivered.ToCanonicalUtf8Bytes(),
        ["semantic-transaction-1"] = transactionChanges.ToCanonicalUtf8Bytes(),
    });
```

Construction and verification share `DeliveryReceiptLimits`. Defaults cap the complete receipt
JSON at 16 MiB, each semantic artifact and PageMap at 64 MiB, any artifact at 256 MiB, all supplied
artifacts at 512 MiB, JSON depth at 128, aggregate collection items at 100,000, transactions and
operations per transaction at 10,000, artifacts at 1,024, and strings at 1 MiB. Clean-DOCX ZIP/XML
inspection additionally uses `CleanDocxManifestOptions`, whose corrected #493 package limits are
authoritative. Verification accepts overrides through `DeliveryReceiptVerificationOptions` and
returns stable `*_resource_limit` findings rather than parsing, decoding, copying, or hashing past
the applicable boundary. Object-based construction and verification preflight aggregate collection,
escaped-string, semantic-value depth, and serialized-byte budgets before canonicalization. Receipt,
canonical JSON, and exact typed-semantic writers additionally write through hard-capped streams, so
the configured ceiling is enforced while bytes are produced rather than after an oversized buffer
already exists. Every public receipt-output path is bounded: the overloads without limits use
validated defaults; the overloads accepting `DeliveryReceiptLimits` cap payload and compact
envelope output at that instance's `MaxReceiptJsonBytes`; and an indented request re-emits the
already canonical envelope through a second stream capped at the same configured value. Indented
output may therefore be rejected even when its compact equivalent fits. JSON arrays are charged one
item at a time before recursion, and object properties are charged before entering the capped
sortable property list; neither path first materializes an untrusted collection merely to learn
that it exceeds the item budget.

## Schema evolution

The checked-in v1 schema is the exact, closed writer projection. Portable readers provide a narrow
additive rule only for a `fullEvidence` payload: after authenticating the raw payload, they may
ignore unknown optional object properties but still require the complete known projection with
exact values and array lengths. Redacted profiles reject unknown properties recursively. All
profiles reject unknown enum values, digest algorithms, required evidence kinds, and major
versions. The envelope itself has exactly `payload` and `receiptDigest`, preventing unprotected
top-level extensions. A new required field, changed canonicalization or escaping rule, changed
field meaning, or changed default attribution/privacy behavior requires v2 and a new schema
identifier.

Writers always emit one exact schema identifier and canonicalization profile. Readers may
migrate an older payload to an internal model for display, but verification uses the rules
declared by that payload's own version. A migrated reserialization is a new receipt with a
new digest, not the original evidence.

## Dependency boundary

The receipt consumes package manifests from #456, semantic change sets from #457, and the
policy-neutral package delta from #463. All package-manifest API assumptions are confined to
`DeliveryPackageManifestAdapter`, which uses the corrected #493 contract, rejects manifests whose
`IsValid` flag is false, preserves the distinction between unavailable optional content/semantic
digests and invalid packages, and adapts the shared delta instead of defining another entry or
relationship comparator. The sole #457 seam accepts a typed `SemanticChangeSet`, stores its exact
`ToCanonicalUtf8Bytes()` output, and records the public schema, version, and count. Portable
verification strictly reconstructs the full nested #457 change/value schema and requires its exact
canonical bytes; receipt fields are not used as a substitute for that owner-defined contract.
The receipt hash-addresses validation and reversibility artifacts from #463/#464 without flattening
or renaming their fields. Their generic references require the exact owner schema identifier. When
a reference names an artifact, the artifact must be available under the owner-specific role, its
digest must equal the reference digest, and both construction and portable verification require the
supplied bytes to be the owner's exact canonical serialization: the full closed owner type is
deserialized and reserialized, so whitespace, reordered properties, omitted/null substitutions,
unknown fields, and merely rehashed lookalikes fail. A reference without an artifact ID is only a
schema-and-digest assertion because no owner bytes are present to parse. In either case, the receipt
does not reinterpret the owner's pass/fail or success decision; a consumer that relies on that
decision applies the #463 or #464 semantics to the authenticated owner object. The delivery
operation in #465 assembles artifacts and invokes the receipt builder; receipt construction itself
performs no document mutation, rendering, ZIP inspection, or OOXML parsing outside those shared
components.
