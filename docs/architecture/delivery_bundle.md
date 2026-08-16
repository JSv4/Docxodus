# Verified delivery bundles

Issue #465 adds one transport-neutral operation for deriving a requested set of document,
render, and evidence artifacts from exact baseline and working DOCX snapshots. The operation
returns immutable bytes plus a canonical manifest, or publishes the same result as a fresh,
atomically committed directory.

## Contract and flow

`DeliveryBundleService.BuildAsync` accepts a `DeliveryBundleBuildRequest` containing:

- named baseline and working snapshots with explicit document versions;
- the name and version to assign to the policy-derived final snapshot;
- separate `preserve`, `accept`, or `reject` policies for revisions already present in the
  baseline and revisions generated between baseline and working;
- an explicit list of artifact IDs, kinds, requiredness, and render profiles; and
- optional authoritative #458 mutation evidence for a change receipt.

The service snapshots all caller-owned inputs before processing. It first derives the policy
baseline, review, and final states and proves that accepting the review document produces the
named final state while rejecting it produces the policy baseline. It then plans artifacts in a
stable order, creates document and evidence outputs, calls an injected renderer only for declared
render capabilities, composes an authoritative receipt when requested, and constructs the
manifest. Required unavailable artifacts fail the operation unless the caller explicitly requests
an incomplete byte-return result.

Every artifact is bound to its exact baseline, working, review, or final source through the
manifest identities and typed relationships. A `usesPageMap` relationship is emitted only when
the renderer's sidecar bytes, source package digest, and renderer fingerprint match the separately
requested PageMap artifact.

## Supported artifact vocabulary

Schema v1 has a closed artifact vocabulary:

- baseline, policy-baseline, working, native-review, and final DOCX;
- standalone HTML, final PDF, review PDF, PageMap, and render report;
- baseline and final package manifests;
- semantic and package deltas;
- comprehensive deliverable validation;
- redline reversibility proof; and
- deterministic delivery change receipt.

Receipt construction never guesses mutation history from a before/after comparison. A caller must
supply `DeliveryReceiptContext` entries containing the authoritative #458 transaction contribution
and its exact before/after snapshots. The context must form a continuous chain that reaches the
bundle's named final snapshot. Per-transaction semantic evidence added for receipt closure is
declared as an implicit required artifact.

## Rendering boundary

The delivery core owns artifact intent, source selection, metadata validation, relationships,
failure policy, and verification. It does not discover a renderer. An
`IDeliveryArtifactRenderer` declares its supported artifact kinds, review profiles, and comment
profiles and returns bytes with its fingerprint, page count, PageMap, report, and diagnostics.
Renderer diagnostics remain structured end to end: the canonical manifest retains their stable
code, severity, phase, message, remediation, owning part, anchor, and resource identity rather than
flattening them into display strings.

`DocxodusExportHostRenderer` is the production boundary to epic #434. It requires absolute Node
and built-host paths owned by the process, launches without a shell or PATH guessing, uses one
length-framed request/response envelope, and renders each exact source/review/comment cohort only
once. HTML, PDF, PageMap, and render-report results from that cohort are validated together before
they enter the bundle. Every render cohort implicitly includes its PageMap and report sidecars;
their requiredness follows the cohort's requested outputs. If the adapter is not configured, the
CLI and MCP surfaces retain truthful unavailability rather than substituting the LibreOffice
legal-evaluation harness or continuous preview HTML.

If the framed host returns a schema-valid failed render report, the adapter preserves that report
as an available evidence artifact while marking HTML, PDF, and PageMap unavailable. The report's
source/profile binding, failure identity, and structured warnings are checked before retention; a
failure that occurred before renderer fingerprinting remains truthful rather than inventing an
identity. Callers can retain this evidence through the explicit incomplete-bundle path.

## Manifest and independent verification

`bundle-manifest.json` is a canonical JSON envelope. Its digest covers the payload, which records:

- exact baseline, working, and final names, versions, byte sizes, and SHA-256 digests;
- the two-part revision policy;
- every requested or implicit artifact, including explicit unavailable entries;
- byte size, SHA-256 digest, MIME type, portable relative path, and render metadata including the
  exact source name, document version, package digest, and structured renderer warnings; and
- deterministic typed relationships between artifacts.

The formal schema is
[`delivery-bundle-manifest-v1.schema.json`](../schemas/delivery-bundle-manifest-v1.schema.json).
`DeliveryBundleVerifier.VerifyJson` reparses manifest bytes, recomputes the payload digest, checks
canonical ordering and path/resource limits, and independently matches every supplied artifact's
size and digest. The verifier takes artifact bytes separately, so it does not trust the in-memory
bundle that created the declarations.

The required validation artifact uses
[`delivery-bundle-validation-v1.schema.json`](../schemas/delivery-bundle-validation-v1.schema.json).
It retains one baseline-aware verification of the final DOCX plus one verification for every
review/comment render cohort. Each cohort is compared with its own exact source snapshot so source
package findings remain correctly classified as pre-existing, while malformed, mismatched, or
unbound companion outputs remain new delivery blockers. The aggregate decision fails when any
constituent verification fails and otherwise preserves `passedWithPreExistingFindings` or
`notEvaluated` semantics from the selected verification mode.

## Atomic directory publication

`DeliveryBundleDirectoryPublisher.Publish` only targets a new absolute directory. It creates a
private, marked sibling stage on the same filesystem, writes artifact files, writes the manifest
last, rereads and verifies the staged bytes, checks the stage again immediately before commit, and
renames the directory as the single commit point. Any failure removes only the owned stage. It
never replaces an existing target or returns a misleading partial directory.

Callers that explicitly need failure diagnostics can request an incomplete in-memory bundle.
Directory publication rejects incomplete and failed manifests by default; a separate explicit
diagnostic-publication option can retain them without representing them as a successful delivery.

## Surfaces

The .NET API is authoritative. `docxodus-deliver` is a thin command-line adapter over the same
service and publisher; its complete syntax and renderer limitations are in
[`tools/delivery/README.md`](../../tools/delivery/README.md). The MCP `docxodus_deliver` tool uses
the same service with a named baseline and the current session, returning canonical manifest bytes
and bounded base64 artifacts. Neither transport fabricates receipt history or render capability it
does not possess.

End-to-end coverage requests every schema-v1 artifact, drives the production #434 host in real
Chromium, publishes a fresh directory, reopens every file independently, verifies source bindings
and the canonical manifest, and retains an HTML evidence index plus checksums even when an
assertion fails. Fast unit coverage also uses a clearly identified deterministic test renderer.
