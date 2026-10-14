# Backend operation reconciliation

`DocxVersionHistory.SubmitOperationAsync` is a storage-neutral backend producer over the SAME
blob/head adapters, exact versions, package effects, and durable request journal as named history.
It does not create a GUI, network service, subscription, optimistic client, or real-time editor.
Hosts authenticate actors, persist captured requests, supply storage, and decide when to submit.

```csharp
var history = new DocxVersionHistory(hostBlobs, hostHeads);
var initial = await history.CreateVersionAsync("contract", "initial-request", null,
    capturedDocx, capturedMetadata);
var request = new DocxOperationRequest
{
    RequestId = "replica-uuid:42", Base = initial.Head, Kind = "text",
    Metadata = capturedMetadata,
    Text = new DocxTextSplice("/word/document.xml", TextNode: 0,
        Offset: 8, DeleteCount: 0, Insert: " proposed"),
};
// Persist request once before submission. Retry exactly this input after an uncertain response.
var outcome = await history.SubmitOperationAsync("contract", request);
if (outcome.Operation.Record.Status == "conflict")
{
    byte[] contender = await history.ExportOperationProposalAsync("contract", outcome.Operation.Id);
}
```

## Intent, effects, and publication

The bounded, generated v1 input codec records document/request ID, original exact base head,
operation kind, copied host metadata, sorted/unique read dependencies, optional resolution link,
and the exact candidate-byte reference for package submissions. Its canonical manifest digest
is the input fingerprint. Input capture finishes before awaiting storage. A request ID is shared
with the version/restore namespace, not inferred from timestamp or content identity.

The decision records original input, exact preceding publication, previous decision, assigned
publication revision, preserved proposal snapshot, accepted snapshot/version, optional actual
content commit, status/reason, and the mapped text splice when applicable. One CAS publishes
state + decision tip + request receipt. Content-changing acceptances create an exact version and
replayable import effects using the existing version producer. Conflicts, discards, and accepted
no-ops publish an outcome but do NOT fabricate a version or content commit.

The same ID/input returns its original outcome, not the current head, after later publications,
restarts, competing identical calls, or a lost acknowledgement. Changed input under that ID
raises `RequestConflict`. On a different writer's successful CAS, the loser re-reads and reconciles
the ORIGINAL submitted intent against the new accepted tail; it never silently replaces its base.
Unknown storage failures propagate. Retry the original request to resolve a possibly committed
result. Cancellation after this call's own successful CAS cannot turn it into failure.

## Supported reconciliation rules

| Intent / concurrent change | Outcome |
|---|---|
| Text edits in disjoint ranges or different `w:t` nodes | Map UTF-16 positions through accepted text effects and retain both. |
| Inserts at the same gap | Keep both in assigned acceptance order. |
| Overlapping replacements/deletions, or an insertion inside a deleted range | Preserve the contender as `OverlappingText`; do not silently delete concurrent work. |
| A concurrent insert inside an incoming deletion's observed range | Conflict; a stale deletion never expands to swallow inserted text. |
| Unknown/package mutation of the addressed text part | `UnknownTextChange`, even if a convenient-looking ordinal still exists. |
| Package writes to disjoint entries | Merge complete entry payloads atomically; untouched entries retain their uncompressed bytes. |
| Package entry already equals the proposed value | Compatible no-op for that entry; intentional request still gets an outcome. |
| Other overlapping package writes | Preserve the entire proposal as `ChangedPart`; no partial application. |
| Changed declared read dependency or package topology | `ReadDependencyChanged`. |
| A restore since the submitted base | `EpochChanged`; preserve the proposal instead of replaying it over the reset. |

Text addressing is **relative to an immutable base**, not a persistent sidecar identity system.
`TextNode` is a zero-based document-order ordinal of `w:t` in the named canonical OPC part URI.
Recognized text operations preserve those elements and their order; they change only text and
`xml:space`. Boundaries cannot split surrogate pairs. Inserted text must be valid XML characters.
Tracked-change/comment topology and opaque parts are retained, but this primitive does not
author new Word tracked revisions or merge arbitrary paragraph/table/review structures.

`Kind = "package"` takes the original candidate DOCX in the separate `candidateDocx` argument.
Declare any semantic whole-part dependencies in `ReadParts`; expected absence is meaningful too.
All content-type and relationship entries are guarded automatically. The library cannot infer
every application-specific semantic read from two snapshots. Package merging is NOT an arbitrary
same-part XML merge, and passing an incomplete semantic read set is not a correctness guarantee.
For text, the intervening contribution endpoints, actual applied footprint, and mapped text
effects are verified before using a recorded position map. Unknown edits cannot masquerade as
safe ordinal-preserving splices merely by matching the latest visible text.

## Inspecting and resolving conflicts

`ReadOperationsSinceAsync(documentId, acceptedHead)` returns ascending, validated durable outcomes
and a captured current view. Null scans all retained decisions within the budget. Ordinary
versions/labels may create gaps in decision revision numbers; publication ancestry proves those
gaps rather than assuming every revision is an operation. Content-only consumers continue using
`ReadChangesSinceAsync`; decision-only publications have empty content tails.

`GetOperationAsync` reads a retained reference and validates its codec/input relationships; a
reference alone is neither authorization nor proof it was published. The ordered reader correlates
each decision with its exact original publication, inline receipt, preceding decision, and
version/content commit. It walks publication ancestry once, without restarting the radix receipt
lookup for every decision. Missing/corrupt records fail explicitly, not as an empty log.

Resolve a conflict with a NEW request, using a base that has observed the conflict and setting
`Resolves` to its decision reference. Submit an explicit text/package change or `Kind = "discard"`.
A successful resolution is durable even without a content change. Only one accepted resolution
can claim the original conflict; a competing later one records `AlreadyResolved`. A resolution
that itself conflicts does not mark the original resolved. All attempts and proposals remain
available. This API does not decide on behalf of the host which user's contested edit is preferred.

## Limits, retention, and current cost

`maxAttempts` defaults to 16 (1–1024), with typed `Contention` on exhaustion; retry the same request.
`maxEntriesToScan` defaults to 10,000 and bounds each content/publication ancestry scan and each
decision-publication walk independently, per attempt. Intervening labels count too. Resolution
currently scans back through retained decisions to prove the target and detect prior resolutions.
Raise limits explicitly, pass cancellation, and impose host request timeouts/aggregate quotas.
Input/decision manifests are bounded to 512 KiB; inserts to 64 Ki UTF-16 characters; declared read
sets to 256 canonical URIs. Existing snapshot, OPC expansion, and contribution limits still apply.

This is a correctness-first backend/package path: it loads snapshots, verifies full contributions,
and rewrites ZIP packages. It is NOT the planned incremental keystroke path. Performance scales
with package size and intervening history. No claim of low-latency typing, automatic client
outboxes, selective undo, or complete #671 collaboration is made.

Retain the reachable state/publication, decision/input, proposal, version/snapshot, effect/payload,
and receipt-index graph. Orphaned prepared blobs after failed CAS are host cleanup work; never
delete reachable receipts merely because a response was acknowledged. The reference file adapter
supports cooperating local processes, not distributed/network filesystem coordination. Process
crash recovery and filesystem power-loss durability are distinct guarantees.
