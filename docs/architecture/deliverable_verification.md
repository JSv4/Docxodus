# Deliverable verification

> **Status:** Implemented in `Docxodus/Verification`. Report schema:
> `https://docxodus.dev/schemas/verification/deliverable-verification/v1`.

`DeliverableVerifier.VerifyDeliverable` is the final-delivery gate for DOCX bytes. One call composes
the bounded OPC manifest, the Open XML SDK validator, WordprocessingML cross-part closure checks,
bounded part-aware workflow/revision scanners, semantic/package deltas, and any supplied
renderer evidence. The result is a deterministic report with a policy decision and actionable,
stable findings.

The byte-array entry point is intentional. A corrupt, encrypted, or safety-limited input cannot be
constructed as a `WmlDocument`, but it still needs a structured failure report. `WmlDocument`
overloads are conveniences for already-openable documents. Neither form mutates caller bytes or
documents.

## Entry points

```csharp
var report = DeliverableVerifier.VerifyDeliverable(finalDocxBytes);

var compared = DeliverableVerifier.VerifyDeliverable(
    baselineDocxBytes,
    finalDocxBytes,
    new DeliverableVerificationOptions
    {
        Mode = DeliverableVerificationMode.Standard,
        FailOnUnexpectedChanges = true,
    });

var canonicalJson = compared.ToCanonicalJson();
```

For a live editing session, `session.VerifyDeliverable()` verifies the same clean-save form returned
by `Save(false)`, including unsaved cached XML and excluding internal projector anchor attributes.
When the bytes will be delivered, prefer `session.PrepareDeliverable()`: it returns the exact byte
snapshot together with the report that hashes it, eliminating a verify/serialize identity gap. If
the session captured its initial state (the default), those exact opening bytes become the baseline.

The full request form adds approved deltas and companion artifacts:

```csharp
var report = DeliverableVerifier.VerifyDeliverable(
    new DeliverableVerificationRequest
    {
        BaselineBytes = baseline,
        DeliverableBytes = final,
        ExpectedSemanticChanges = approvedSemanticChanges,
        ExpectedPackageChanges = approvedPackageChanges,
        CompanionArtifacts = new[] { pdfEvidence, htmlEvidence },
    },
    new DeliverableVerificationOptions { FailOnUnexpectedChanges = true });
```

## What is checked

The operation records the status and finding count of each stage. `UnavailableEvidence` and
`SkippedPrerequisiteFailed` are different from an empty successful check; standard and strict modes
fail closed when analysis cannot complete. Safety failures (malformed/unreadable ZIP data,
encryption, or a size/expansion/resource limit) are hard prerequisite boundaries. Ordinary OPC
defects such as duplicate or dangling relationships do not suppress independent checks over XML
that preflight read safely; SDK exceptions become structured unavailable-evidence findings.

- **Bounded OPC/package inspection:** ZIP shape, duplicate entries, content types, encryption,
  expansion budgets, relationship owner/target closure, dangling relationship IDs, and exact/raw,
  ordered-content, and normalized-semantic SHA-256 identities.
- **Pinned Open XML validation:** `OpenXmlValidator` runs against the requested
  `FileFormatVersions` value (Office 2019 by default). Validation errors retain part URI and XPath.
- **WordprocessingML closure:** relationship-reachable bookmark pairs and hyperlink targets; comment markers and
  definitions; footnote/endnote references and definitions; move pairs; content-control shape,
  IDs, placeholder state, and reachable custom-XML bindings; numbering instances/abstract definitions,
  levels, and level overrides;
  complex and simple field shape; media relationship closure.
- **Workflow residue:** explicit bracketed blanks/instructions, bare underscore runs, `{{...}}`,
  `${...}`, `<<...>>`, configured exact tokens, and configured case-sensitive editorial markers,
  across relationship-reachable body/stories/notes/comments. Broad square-bracket alternatives are
  opt-in and advisory, so legal references such as `[Section 4.2]` and `[1]` do not block by default.
  Findings carry the owning part, structural path, scope, and stable anchor when one exists.
- **Revision registry:** malformed, ambiguous, and unsupported native tracked-change groups remain
  visible instead of being silently ignored.
- **Render risks and evidence:** static OOXML risks (for example altChunk, Office Math, OLE objects,
  legacy form fields, and vector media) are warnings. Actual missing fonts, font substitutions,
  unsupported content, and renderer warnings must be supplied by the renderer that observed them;
  verification does not read process-global font-warning state.
- **Semantic and package deltas:** `SemanticDiff` supplies modeled changes. A separate manifest
  comparator retains every changed part and relationship, so an unmodeled change cannot be hidden
  merely because another change in the same XML part was modeled.

The package safety limits are shared with `PackageManifestGenerator` and `SemanticDiff`. Companion
artifacts have separate per-item and aggregate byte limits. Semantic detectors additionally share
budgets for XML nodes, relationships, text, regex matches, and general steps. Reaching a finding,
work, or evidence budget produces a deterministic resource finding and makes `analysisCompleted`
false rather than producing a misleading pass.

## Findings and baseline disposition

Every finding contains:

- a stable `code`, category, severity, owning part, optional OPC location/XPath, scope, and anchor;
- a human-readable message and explicit remediation;
- `new`, `preExisting`, `resolved`, or `unclassified` disposition;
- `blocksDelivery`, which is the selected policy's decision for that finding.

Baseline matching is an exact multiset match over detector/rule version, owner/location,
scope/XPath, native anchor or deterministic structural path, and detector subject. It does not
match by message/remediation text or by aggregate counts. Duplicate observations
receive deterministic occurrence numbers. A `findingId` hashes this detector identity plus the
occurrence; message, severity, remediation wording, and policy disposition are deliberately not
part of the ID. The same underlying condition therefore keeps its identity as policy or explanatory
wording evolves.

`resolvedFindings` contains baseline observations no longer present, unless
`IncludeResolvedFindings` is false. Resolved findings never block delivery.

## Policy modes

| Mode | Decision behavior |
|---|---|
| `Standard` | Blocks new/unclassified errors, unsafe package/structure errors even when pre-existing, and unresolved workflow placeholders when `RequireNoPlaceholders` is true. An unchanged pre-existing Open XML validation error is reported but grandfathered. |
| `Strict` | Blocks every current warning or error, including pre-existing findings. |
| `ReportOnly` | Collects the same evidence, sets every `blocksDelivery` to false, and returns `NotEvaluated`. |

If nothing blocks but warning/error findings were inherited from the baseline, the decision is
`PassedWithPreExistingFindings`; otherwise it is `Passed`. Incomplete analysis returns `Failed` in
standard/strict mode regardless of the findings retained so far.

`FailOnUnexpectedChanges` requires a baseline. Semantic expectations match the complete typed
semantic-change identity (generated `chg-*` display IDs do not matter). Package expectations match
change kind and exact location, plus optional before/after raw digests and canonical values. Both
unexpected actual changes and approved changes that did not occur are errors.

## Companion artifacts

`DeliverableCompanionArtifactInput` records a stable ID, role, media type, availability, bytes,
page count, renderer fingerprint, page-map digest, source-package digest, and structured renderer
diagnostics. Available bytes are hashed into report metadata. Available PDF, HTML, and PageMap
evidence must bind to the exact delivered package, renderer fingerprint, and page count. PageMap
JSON is parsed with the shared portable schema-v1 contract and assigned a canonical digest; PDF and
HTML evidence must reference supplied PageMap bytes by that digest and agree on source, renderer,
and count. Basic role/media/format checks reject placeholder bytes such as a bare `%PDF` header.
Unavailable outputs remain explicit metadata instead of disappearing from the report.

The verifier does not create PDF or HTML and these closure checks do not prove visual fidelity. It
verifies basic format, metadata, and diagnostics for artifacts a renderer already produced. This
keeps conversion policy separate from delivery policy and avoids guessing whether a renderer
warning applies to the exact delivered bytes.

## Canonical report and schema

`ToCanonicalUtf8Bytes()` and `ToCanonicalJson()` emit compact UTF-8 JSON with deterministic member
and collection order. Package and semantic identities bind the report to exact inputs; companion
artifact metadata binds outputs to those inputs. The checked-in JSON Schema is
[`docs/schemas/deliverable-verification-v1.schema.json`](../schemas/deliverable-verification-v1.schema.json).

The report is evidence, not a digital signature. Sign or store its canonical bytes in the calling
system if tamper evidence or non-repudiation is required.

## Deliberate limits

- Static inspection cannot prove visual fidelity. Attach diagnostics and, where appropriate, a PDF,
  HTML, page map, or page images from the actual delivery renderer.
- Field results, table of contents, pagination, and linked external resources can be stale even when
  markup is valid. Static field/relationship checks report broken structure, not refreshed layout.
- Open XML conformance and Word's repair behavior are related but not identical. Both the SDK
  validator evidence and Docxodus closure findings remain visible.
- Version 1 does not expose a dedicated WASM/npm/Python/MCP request model for the full artifact and
  expected-delta payload. The .NET byte/session operations and shared internal session facade are the
  authoritative implementation seam for subsequent transport bindings.
