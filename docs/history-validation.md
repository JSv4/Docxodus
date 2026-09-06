# Shared history infrastructure validation

## Version model/fault fuzzing

Run `bash scripts/history-fuzz.sh` for the extended corpus. Default ordinary test discovery runs
16 seeds × 96 steps; the script runs 64 × 128. Override `DOCXODUS_HISTORY_FUZZ_SEEDS` (1–1024) and
`DOCXODUS_HISTORY_FUZZ_STEPS` (1–10000) explicitly for larger campaigns. A failure reports seed,
step count, storage kind, and command trace. TRX output is in `Docxodus.Tests/TestResults`.

The recorded run below passed on 2026-09-06 (.NET 10.0.301 SDK / 10.0.9 runtime, Linux).
Generated race cases rendezvous both contenders at the actual head CAS, release them in seeded
order, and assert two attempts. Counts come from `history-model-fuzz-gated-64x128.trx`.
The complete corpus was rerun after backend integration with identical counts: all 72 version
model/fault/index cases passed in 3.59 minutes (`history-fuzz.trx`).

| Observed coverage | Count |
|---|---:|
| Passing seeds | 64 |
| Seeds using real filesystem adapters | 13 |
| Scheduled steps | 8,192 |
| Accepted versions, each checked by exact exported bytes | 5,409 |
| Content sequences, each materialized and effect-replayed | 3,548 |
| Restore operations | 640 |
| Metadata/byte-preserving repacks | 667 |
| Legacy saves carrying earlier receipts forward | 644 |
| Old request retries during generated streams | 712 |
| Same-ID/different-input refusals | 689 |
| Stale writes refused, then intentionally corrected | 679 |
| Corrupt/missing snapshot refusals | 690 |
| Competing different-request races | 327 |
| Concurrent identical-request races | 349 |
| Recoveries from lost post-commit acknowledgements | 245 |

All 64 cases passed in 3.57 minutes. Each seed also retries every accepted identified request at
the end, checks immutable pagination against its independent version list, and resolves every
content-boundary timestamp against its own recorded-time model. The oracle compares retained
input bytes and uncompressed ZIP entry bytes; it does not use the production package digest or
replay implementation to calculate expected content. Fixtures include Unicode and opaque binary
custom parts. Timestamps move backwards and can tie; labels do not fabricate content sequences.

Separate fault-matrix tests enumerate every immutable write before/after its durable store call,
plus failures/cancellation before and after head CAS, for create AND restore on memory AND real
filesystem adapters. Before-commit failures retain the old head; after-commit acknowledgement
loss returns the original result on retry; cancellation after successful CAS cannot erase it.
Producers are discarded before recovery. All four matrix cases passed. The receipt-index tests
also insert/read 3,000 randomly ordered IDs, enforce the 257-node lookup bound, and check that
lookup writes nothing and corruption is not treated as an absent request.

These are deterministic model/fault campaigns, not a claim of exhaustive state-space coverage,
power-loss certification, or a distributed-filesystem guarantee. Actual process-kill and backend
reconciliation evidence are separate gates; in-process service recreation is not a process kill.

## Backend reconciliation and fault tests

On 2026-09-06, all 18 `DocxBackend*` tests passed (`history-backend-core.trx`, 20.43 seconds).
These execute the real shared history backend over DOCX packages, not a standalone string model:
same-gap/disjoint text, contested ranges with exported proposals, explicit one-winner resolutions,
disjoint header/footnote package changes, read dependencies, no-ops/discards, restore epochs,
unknown structural boundaries, delayed duplicate requests, and input identity conflicts.
Four race cases (memory/filesystem × seeded left/right winner) force both head CAS attempts;
different compatible intents require a third reconciled publication attempt, while identical
requests require exactly two attempts and return one original outcome.

Four additional matrix cases enumerate every immutable write before/after its durable call and
both sides of CAS/cancellation for accepted AND conflicting operations on memory AND filesystem
storage. Recovery validates exact original receipts, preserved contender text, visible text,
version/decision counts, and recorded-effect replay. Record negatives reject malformed/future
codecs, missing/corrupt metadata/proposals, altered fingerprints, and forged valid-range text maps.
The decision audit is forbidden from reading even a single radix-index root and has exact budget
boundary tests including intervening named versions. Untouched package parts (notes, headers,
comment topology, relationships, and an opaque binary custom part) retain their entry bytes;
saved/reopened packages introduce no OpenXML validation errors beyond the independent baseline.

## Seeded backend stream campaign

`bash scripts/history-backend-fuzz.sh` runs 32 seeds × 16 rounds, all backend fault/record tests,
and the subprocess recovery matrix. Ordinary discovery defaults to 8 × 8; override
`DOCXODUS_BACKEND_FUZZ_SEEDS` (1–256) / `DOCXODUS_BACKEND_FUZZ_ROUNDS` (1–1024) for larger runs.
The full campaign passed on 2026-09-06: 58 cases in 6.19 minutes (`history-backend-fuzz.trx`).

| Observed backend model coverage | Count |
|---|---:|
| Passing seeds / filesystem seeds | 32 / 8 |
| Rounds with original-base, reordered requests | 512 |
| Accepted content changes | 2,048 |
| Accepted no-ops | 512 |
| Preserved overlapping conflicts | 512 |
| Explicit conflict discards | 512 |
| Forced competing publication races | 266 |
| Lost-acknowledgement recoveries | 480 |
| Old-request retries within streams | 886 |
| Final retry of every published operation | 3,584 |
| Exact versions exported and effect-replayed | 2,560 |

The text oracle retains observed character identities and original gap anchors; it never uses
production mapped offsets, digests, or replay to calculate expected text. Every outcome checks
content/publication positions, proposal text, untouched package bytes, normalized non-target body
and review/reference topology, and OpenXML validation-delta. Delivery order and CAS winner order
are seeded; failures report seed, rounds, and trace. This deliberately bounded operation-family
campaign is not exhaustive arbitrary structural/XML-operation coverage or full #671 collaboration.

## Real subprocess crash recovery and generated serialization

`DocxHistoryProcessRecoveryTests` builds the test-only `tools/history-recovery-probe` through a
project reference, so normal CI test builds cannot silently omit it. All eight cases passed
standalone (20.67 seconds, `history-process-recovery.trx`) and again in the extended backend run.
Each uses separate seed, crash, and recovery processes over the same private file-backed root.
The producer is actually killed immediately before entering or immediately after a successful
head CAS; Linux asserts SIGKILL exit code 137, not an ordinary thrown exception or service reopen.

Version create, restore, accepted text, and conflicting text each run at both boundaries.
The fresh process verifies pre/post-commit visibility, exact original retry result, later-write
old retries, immutable version/decision counts, preserved proposals, restore epoch, and replayed
ZIP entries. Reflection-based JSON serialization is disabled in all three processes; actual
history operations plus generated V3/output graph round-trips must work without it.

These tests do not interrupt inside the file adapter's flush/rename internals and do not certify
filesystem power-loss durability. The host still owns storage guarantees, retention, transport,
authorization, and aggregate quotas. No GUI, real-time editor, or outbox is implemented.
