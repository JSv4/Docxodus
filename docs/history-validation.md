# Shared history infrastructure validation

## Version model/fault fuzzing

Run `bash scripts/history-fuzz.sh` for the extended corpus. Default ordinary test discovery runs
16 seeds × 96 steps; the script runs 64 × 128. Override `DOCXODUS_HISTORY_FUZZ_SEEDS` (1–1024) and
`DOCXODUS_HISTORY_FUZZ_STEPS` (1–10000) explicitly for larger campaigns. A failure reports seed,
step count, storage kind, and command trace. TRX output is in `Docxodus.Tests/TestResults`.

The recorded run below passed on 2026-09-06 (.NET 10.0.301 SDK / 10.0.9 runtime, Linux).
Generated race cases rendezvous both contenders at the actual head CAS, release them in seeded
order, and assert two attempts. Counts come from `history-model-fuzz-gated-64x128.trx`.

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

Actual process-kill recovery and a larger seeded backend stream campaign remain follow-up gates;
these results do not claim either one or full real-time/editor collaboration.
