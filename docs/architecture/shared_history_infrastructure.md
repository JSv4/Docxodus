# Shared publication and backend reconciliation infrastructure

Scope: storage-neutral durable publication, version-history fault/model fuzzing, and a backend
operation-stream reconciler. No GUI, transport, subscriptions, presence, or real-time editor.

## Completion gates

1. Durable request identity: the same document/request ID and canonical input returns the original
   committed result after restart, later writes, concurrent retries, or a lost acknowledgement.
   Reusing an ID with different input fails. Byte deduplication is independent of request identity.
2. One atomic publication exposes state and its request receipt together. Failures before it
   leave the old state; failures after it are recoverable by retry. No process-local dedup cache
   is authoritative. Existing immutable version records remain readable.
3. Version creation/restoration use that machinery, preserving exact bytes, lineage, metadata,
   content sequence/epoch, stale-write guards, and current host-owned storage contracts.
4. A backend acceptance log records operation IDs, original bases, accepted order, resolved
   effects/outcomes, and recoverable conflicts. Reconciliation combines compatible work and
   preserves incompatible contenders for explicit resolution. Replay never reruns clocks or
   allocators. Multiple instances/restarts and delayed/repeated/reordered submissions are tested.
5. Reproducible model/fault fuzzing covers version create/label/repack/restore/list/export/replay,
   retries and ID collisions, competing writers, cancellation, corrupt/missing storage, and
   failures at publication boundaries. Backend conflict streams are checked against independent
   expected content/outcomes, including cross-part dependencies and preservation of opaque data.
6. Document the adapter requirements, retry/retention semantics, conflict policy, limits, and
   executed seed/operation counts. Run existing history and binding regressions. Review each PR.

## Atomic receipts without circular content references

Each publication stores the latest request identity inline, plus an immutable index of earlier
receipts. The latest result is the publication's own head. The next publication promotes that
receipt into the index, now that its exact head is known, and writes its own request inline.
Thus a receipt can identify the exact original result without a hash cycle between the head,
index, and receipt. An immutable compressed radix index bounds lookup by request-key hash width,
not total history length. Legacy publications carry the index forward even without a request ID.

Version request fingerprints cover the operation, document, exact expected head, input snapshot
bytes or restore target, and normalized metadata. Caller IDs are bounded opaque strings scoped
to a document; hosts authenticate callers and allocate replica/request namespaces. Timestamps
are recorded metadata, not deduplication keys or ordering authority.
