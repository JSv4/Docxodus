# Standalone paginated HTML and PDF export

Status: accepted design for epic #434 and issue #437, and the export dependency consumed by epic
#436 / issue #465.

This record fixes the initial public and wire contracts before #438--#444 land. Closed enums,
digest inputs, canonical byte encodings, required report fields, and default resource ceilings are
schema-v1 compatibility boundaries. A follow-on may populate an optional field already reserved by
v1, but a new property—even an optional one—or any changed boundary requires a versioned contract
and cross-language migration rather than a silent reinterpretation.

## Decision

Docxodus will keep pagination and artifact materialization in the browser/WASM package and add the
Node-only `@docxodus/export` companion package. The companion drives that same browser code in
Chromium; it does not implement a second converter, paginator, or print layout. Keeping it in a
separate package prevents Node-only browser/runtime dependencies from entering the existing
browser/WASM installation or its root export graph.

The public layers are:

| Layer | Responsibility | Must not do |
|---|---|---|
| `docxodus` | DOCX conversion, review/comment projection, pagination, PageMap, finalized offline HTML, structured render report | Import Node modules or launch a browser |
| `@docxodus/export` | Isolated Chromium lifecycle, local package-resource serving, readiness, PDF parsing/verification, temporary-file policy, Node byte/path APIs | Import the editor barrel or reimplement conversion/pagination |
| `docxodus` CLI | Parse paths and policy flags, call `@docxodus/export`, write a new output, format errors | Invent defaults that differ from the Node API |
| `DeliveryBundleService` | Request/source/revision policy, verification, relationships, atomic delivery publication | Discover Chromium or treat evaluation renderers as production |

The browser result is therefore the authoritative layout result. A batch render session returns
HTML, PDF, PageMap, and render report from one finalized page tree, with the same source digest,
document version, review/comment profiles, page count, and renderer fingerprint. The #465 adapter
uses that explicit batch contract and projects its immutable results into the delivery artifact
records. It must not emulate a render session with a hidden mutable cache across unrelated
single-artifact calls.

## Browser API

Issue #438 adds this browser/WASM surface:

```ts
type Sha256Hex = string; // exactly 64 lower-case hexadecimal SHA-256 digits
type ReviewProfile = "final" | "original" | "markup";
type CommentProfile = "hidden" | "inline" | "endnotes" | "margin";

interface ExportOptions {
  documentVersion?: number;
  expectedSourceDigest?: Sha256Hex;
  reviewProfile: ReviewProfile;
  reviewProfileAlreadyApplied?: boolean;
  commentProfile: CommentProfile;
  title?: string;
  unsupportedContent?: "warn" | "strict";
  strictFonts?: boolean;
  timeoutMs?: number;
  limits?: Partial<ExportResourceLimits>;
  signal?: AbortSignal;
}

interface PaginatedHtmlOptions extends ExportOptions {
  /** Trusted browser runtime assets only; never a document-resource base URL. */
  wasmBasePath?: string;
  /** Browser-only resolver from #442; Node callers use fontDirectories instead. */
  fontResolver?: FontResolver;
}

interface PaginatedRenderMetadata {
  pageCount: number;
  pageMap: PageMap;
  renderReport: CompleteRenderReport;
  warnings: readonly RenderWarning[];
  rendererFingerprint: Sha256Hex;
}

interface PaginatedHtmlResult extends PaginatedRenderMetadata {
  html: string;
}

convertDocxToPaginatedHtml(
  document: File | Uint8Array,
  options: PaginatedHtmlOptions,
): Promise<PaginatedHtmlResult>;
```

`reviewProfile` and `commentProfile` are required. Node APIs use `ExportOptions`, not the
browser-only fields on `PaginatedHtmlOptions`. `wasmBasePath` may relocate only the closed,
hash-verified runtime graph, and a Node adapter rejects it and `fontResolver` instead of accepting
an alternate asset or callback boundary by accident.

The operation obtains the authoritative raw source-package digest from #493 preflight. The public
digest string is the lower-case `value` of a #493 `VerificationDigest` whose `algorithm` is exactly
`SHA-256`; any other algorithm, malformed value, or unexpected manifest discriminator fails
preflight. `documentVersion` is optional provenance; standalone
conversion uses 0 for its immutable input snapshot when omitted, while session/delivery callers
supply their exact version. A supplied value must be a non-negative JavaScript safe integer. The
returned PageMap is the existing portable v1 contract, measured from the final page tree and bound
to the source digest by the render report. A #465 request whose .NET `long` version is outside that
range returns typed `document_version_unrepresentable` artifact unavailability before a host frame
is built rather than rounding it. The JavaScript API uses the same error code for an unsafe numeric
value. `expectedSourceDigest`, when supplied, must equal #493's recomputed raw-package byte digest
using constant-time digest comparison where the host provides it. The default
unsupported-content policy is `warn`, which means a visible placeholder plus a structured warning;
`strict` fails instead. `title` defaults to the empty string for every API and CLI path; callers do
not get a filename-derived or locale-derived title. The resolved title enters the canonical
layout/materialization options even though it does not affect page geometry, because it changes the
exact standalone HTML bytes. `strictFonts` defaults to false and `timeoutMs` to 120,000. Resource
limits use the shared defaults below and may only be lowered.

`reviewProfileAlreadyApplied` defaults to false. It may be true only for `final` or `original` when
the caller deliberately supplies the exact policy-derived bytes. Preflight then proves that no
native tracked revision remains, records the flag in the report, and does not populate
`derivedProfileSource`; it never tries to infer whether those bytes represent an accepted or
rejected history. The #465 adapter supplies the separately verified source-selection proof. The
flag is invalid with `markup`, and the standalone CLI does not expose it as a casual conversion
switch.

When this operation derives `final` or `original` bytes, it reruns #493 manifest preflight against
that derived package with the same effective lower limits before conversion. The report binds both
source identities. Passing the caller source once and trusting an internally rewritten ZIP without
the same safety, digest, and manifest checks is not an accepted implementation.

On success, the duplicate convenience fields are exact invariants:
`pageCount === pageMap.pages.length === renderReport.pages.length`, `rendererFingerprint` equals
both the PageMap token and `renderReport.environment.rendererFingerprint`, and `warnings` is the
same ordered value as `renderReport.warnings`. Requested HTML/PDF members and every digest binding
must agree with the exact returned bytes. Implementations return caller-owned snapshots; later
mutation of a result cannot change adapter state or a previously computed report digest.

The browser API lays out in an attached, offscreen, same-origin iframe with real layout and a usable
`defaultView`; an unattached `createHTMLDocument()` is not a render surface. Before converted source
HTML is attached, the iframe is bootstrapped with a deny-script/connect/object/form/navigation
policy and a sandbox that retains same-origin DOM access but omits script execution. Active
elements, event handlers, refresh directives, unsafe URL schemes, and automatic document-resource
URLs are removed while parsing in a detached tree and before attachment. It remains laid out using a
fixed offscreen viewport (not `display: none`) and is destroyed after verification/serialization.
The Node adapter supplies the equivalent fresh Playwright page and request-denial policy.
Pagination must use the staging element's `ownerDocument` and its `defaultView` for element
creation, ranges, tree walking, computed style, animation frames, and observers. Replacing only
`document.createElement` would leave a cross-document layout bug. This is a supported isolation
boundary, not a hidden copy of the pagination algorithm. The API treats caller inputs as immutable.
It validates `Uint8Array.byteLength` or `File.size` against the compressed-input ceiling before
allocating a copy, then synchronously snapshots a `Uint8Array` or reads a `File` into one new owned
buffer and rechecks the resulting length before its cancellable worker transfer. The transferred
buffer may detach; the caller's object must not. Options, nested arrays/objects, destinations, and
runtime policy are likewise snapshotted before the first `await`; only the caller-owned
`AbortSignal`, injected `Browser`, and browser `FontResolver` remain live capabilities.

## Standalone HTML contract

The result is a complete `<!doctype html>` document containing the finalized page boxes, not the
staging tree. It has these invariants:

- the measurement staging tree, registries, scripts, transient readiness markers, selection, and
  viewer controls are absent; synthetic page numbers, `will-change`, containment hints, and other
  measurement-only inline styles are removed;
- document CSS is inline, supported package images are data URLs, and no CSS import, stylesheet,
  script, image, media, or font URL can initiate a network request;
- safe `https`, `http`, `mailto`, `tel`, and internal-fragment hyperlinks may remain as
  user-activated links and are recorded in the report; render isolation blocks navigation while
  materializing, but the final document does not falsely disable a supported user link;
- page boxes keep their named `@page` assignment, physical point dimensions, section index, and
  canonical source-fragment metadata;
- screen-only scale, gaps, shadows, background, and host margins are separated from print rules;
  print always uses scale 1, zero browser margins, backgrounds enabled, and one physical page per
  page box;
- visible text remains ordinary DOM text, so it is selectable and searchable; and
- reopening offline must reproduce the same logical page count and section/page geometry without
  rerunning WASM or pagination.

The finalization order is normative: paginate pristine converted content; clone/materialize the
fixed page tree; remove staging, containment, viewer, selection, and other transient state; attach
the sanitized tree with final standalone styles; wait for final-tree stability; measure PageMap
from that sanitized tree; serialize; then reopen the bytes offline in a second isolated context and
verify page count and geometry. Cleanup never runs after PageMap measurement. If cleanup or final
styles change the page-tree signature, the bounded retry starts again from pristine converted
content; the exporter never returns geometry measured from a different DOM than it serializes.

The serialized document carries its own restrictive CSP, but its policy is distinct from the
temporary render sandbox: it prevents scripts, forms, objects, frames, automatic connections, and
automatic media/style loads while retaining the safe user-activated links above. Serialization
audits every URL-bearing HTML/SVG element, attribute, `srcset`, CSS token, nested supported SVG, and
data-URL media type. An unresolved or unsafe automatic resource is
a typed warning or error according to policy; it is never silently retained as an online
dependency. Fixed page boxes preserve page count and physical geometry across hosts; exact visual
and text-flow reproduction additionally requires the same fingerprinted font environment or
license-permitted embedded webfonts. Schema v1 does not claim that arbitrary system or
OOXML-embedded fonts are portable.

The offline reopen waits for the same final resource barrier, re-inventories every page and source
fragment, and compares the canonical PageMap (allowing only the documented point tolerance), not
merely the number and outer size of page boxes. PDF printing consumes that exact serialized HTML
snapshot in a second isolated context. No output-specific DOM mutation may occur after PageMap or
HTML digesting, so an HTML/PDF batch cannot claim common provenance for merely similar trees.

## Node API, batch session, and browser runtime

Issue #439 adds the Node-only surface:

```ts
interface FontLicenseAttestation {
  schemaVersion: 1;
  usage: "standalone-document-font-embedding";
  fileSha256: Sha256Hex;
  embeddingPermitted: true;
  permittedOutputs: readonly ("html" | "pdf")[];
  subsettingPermitted: boolean;
  basis: string;
  attester?: string;
}

interface RenderEnvironmentAttestation {
  schemaVersion: 1;
  usage: "docxodus-render-environment";
  chromiumProduct: string;
  chromiumBuild: string;
  executableSha256?: Sha256Hex;
  launchFlags: readonly string[];
  hostFonts: ReadonlyArray<{
    family: string;
    postscriptName: string;
    style: "normal" | "italic" | "oblique";
    weight: number;
    stretch: number;
    fileSha256: Sha256Hex;
    version: string;
  }>;
  basis: string;
}

interface NodeExportRuntime {
  browser?: Browser;                 // caller-owned; never closed by Docxodus
  browserExecutablePath?: string;    // Docxodus launches and owns this browser
  fontDirectories?: readonly string[];
  fontLicenseAttestations?: readonly FontLicenseAttestation[];
  environmentAttestation?: RenderEnvironmentAttestation;
}

type NodeExportOptions = ExportOptions & NodeExportRuntime;

interface PdfExportResult extends PaginatedRenderMetadata {
  pdf: Uint8Array;
}

convertDocxToPdf(
  document: Uint8Array,
  options: NodeExportOptions,
): Promise<PdfExportResult>;

convertDocxToStandaloneHtml(
  document: Uint8Array,
  options: NodeExportOptions,
): Promise<PaginatedHtmlResult>;

interface RenderBatchResult extends PaginatedRenderMetadata {
  html?: string;
  pdf?: Uint8Array;
}

renderDocxArtifacts(
  document: Uint8Array,
  options: NodeExportOptions & {
    outputs: readonly ("html" | "pdf")[];
  },
): Promise<RenderBatchResult>;

interface RenderFileDestinations {
  htmlPath?: string;
  pdfPath?: string;
  pageMapPath?: string;
  reportPath?: string;
}

interface RenderFileResult extends PaginatedRenderMetadata {
  written: RenderFileDestinations;
}

renderDocxFile(
  inputPath: string,
  destinations: RenderFileDestinations,
  options: NodeExportOptions,
): Promise<RenderFileResult>;
```

The Node `convertDocxToStandaloneHtml` and `convertDocxToPdf` functions are convenience projections
over `renderDocxArtifacts`; the browser `convertDocxToPaginatedHtml` remains the shared materializer
they drive. `outputs` controls only the optional HTML/PDF payload properties. PageMap and report are
always returned as verification metadata; the caller decides whether to persist them. `outputs`
contains no duplicates. An empty array is an intentional metadata-only render for PageMap/report
consumers; convenience calls require their one payload, and the result contains every requested
payload and no unrequested payload. The batch function is the integration surface for #465: one
invocation owns one authoritative finalized page tree and serialized snapshot. It may use separate
fresh contexts to reopen that immutable snapshot offline and print/verify PDF; those contexts never
repaginate or mutate the snapshot and do not create independent provenance. `renderDocxFile` infers
its byte outputs from the
destinations, requires at least one destination, and uses the same batch before creating each
requested file with an atomic no-replace commit. Separate destination files are not an all-or-none
transaction; callers requiring that property publish a fresh containing directory, as
`DeliveryBundleService` already does.

Every Node byte API rejects an over-limit `byteLength` before allocating, then synchronously copies
its `Uint8Array` at function entry, before package preflight,
browser selection, or any `await`, and uses only that owned snapshot for digesting, transport, and
rendering. `renderDocxFile` reads at most the effective compressed-input limit plus one sentinel byte
through a single opened regular-file handle; it never calls an unbounded whole-file read after a
size check. It compares handle and path identity before and after the read and rejects a source that
changes. Later caller-buffer, option, working-directory, or path changes therefore cannot split
source identity from rendered bytes.

Before the session starts, #493 preflight recomputes the source `rawPackageBytesDigest`. A caller
may provide an expected digest; mismatch is `source_digest_mismatch`, not a warning. The delivery adapter
groups requested render kinds by source digest, safe document version, review/comment profiles,
canonical layout-options digest, and runtime policy, invokes one batch per distinct key, and
projects the results. It renders the exact final, policy-baseline, or proven review bytes selected
by `DeliveryBundleService`; it never regenerates a supposedly equivalent profile source.

`RenderReport.bindings.artifactRequestIds` is an integration binding, not an input to layout. It is
an empty array for ordinary browser/Node calls. The framed delivery host supplies the sorted,
unique, bounded IDs for every delivery artifact projected from a cohort before canonical report
serialization. Adapters do not patch an already serialized report, and the IDs do not enter the
renderer fingerprint.

The #465 delivery PR (first drafted as #499, since closed) must replace that draft single-item
delivery seam with an explicit batch seam:

```csharp
public sealed record DeliveryRenderBatchContext(
    DeliveryReviewProfile ReviewProfile,
    DeliveryCommentProfile CommentProfile,
    VerificationDigest LayoutOptionsDigest,
    VerificationDigest RuntimePolicyDigest);

public sealed record DeliveryRenderBatch(
    string BatchId,
    DeliveryRenderBatchContext Context,
    IReadOnlyList<DeliveryRenderRequest> Requests);

public interface IDeliveryArtifactRenderer
{
    DeliveryRendererCapabilities Capabilities { get; }

    DeliveryRenderBatchContext DescribeBatch(
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile);

    ValueTask<IReadOnlyDictionary<string, DeliveryRenderResult>> RenderBatchesAsync(
        IReadOnlyList<DeliveryRenderBatch> batches,
        CancellationToken cancellationToken = default);
}
```

`DeliveryBundleService` validates every plan first, groups render requests by exact source digest,
document version, review/comment profiles, layout-options digest, and configured runtime policy,
then invokes `RenderBatchesAsync` exactly once for the build with all groups in stable order. At
build start it calls the renderer's pure
`DescribeBatch` once for every requested review/comment pair. The returned group-specific layout
digest includes those profiles and every pagination option; the runtime-policy digest covers the
fixed browser/resource/font policy. Repeating `DescribeBatch` for the same pair on one renderer
must be identical. Both digests are algorithm-labelled SHA-256 values over versioned canonical
materials; `RuntimePolicyDigest` describes configured policy, while the post-render fingerprint
describes the actual resolved environment. The service uses the complete context in the group key and passes it unchanged;
the renderer rejects mismatched profiles or digests. Those digests also bind the report, so grouping
cannot depend on hidden mutable adapter state. The dictionary must contain exactly one result for
every requested artifact id and no extras. A genuine one-artifact render is a one-item batch; there
is no `RenderAsync` compatibility loop for a multi-artifact group. The production adapter implements
this contract directly, while test/evaluation renderers are updated to return one-item or
deterministic fixture batches. No adapter may claim shared PageMap/report provenance for results
produced by separate render sessions.

This bundle-level call is what lets the production adapter start one framed host without keeping a
hidden process/cache across independent renderer calls. Each `DeliveryRenderBatch` still maps to
one fresh primary render context and one page tree; bounded offline-reopen and PDF-print contexts
may consume only that batch's immutable serialized snapshot. A genuine one-artifact build is one one-item batch;
there is no `RenderAsync` or multi-item compatibility loop. Empty builds do not invoke the
renderer. The result dictionary contains exactly the union of request artifact IDs across all
batches and no extras.

`@docxodus/export` pins matching exact versions of `playwright-core` and
`@playwright/browser-chromium`. The companion package's normal installation supplies the tested
Chromium; the browser/WASM `docxodus` package does not. Environments that deliberately skip the
browser download must inject or name an executable. Runtime selection is deterministic and
ordered:

1. use the caller-owned injected Chromium browser;
2. launch the explicit `browserExecutablePath`;
3. launch the exact Chromium revision installed by `@playwright/browser-chromium`;
4. fail with `browser_launch_failure` and installation guidance.

The companion never downloads at conversion time. Reproducible/offline deployments install the
lockfile-pinned package and browser artifact during image construction (or from an internal npm
mirror), retain that browser directory with the application, and verify its product/build before
accepting work.

The companion also has an exact-version peer dependency on `docxodus`; the asset manifest package
version and every asset digest must match before Chromium starts. A semver-compatible but different
core/WASM build is not accepted as equivalent. Schema v1 requires Node.js 22.13 or newer; the
published `engines` field and CI matrix are authoritative for later supported Node versions.

There is no PATH-wide browser guessing. `DOCXODUS_CHROMIUM_PATH` is a CLI convenience that maps to
the same explicit API field. The release fidelity contract is the Chromium revision paired with
the pinned `playwright-core`; injected or explicit alternatives remain usable and fingerprinted,
but a differing runtime is outside the committed visual ratchet until separately baselined.

A Playwright `Browser` object does not expose the executable identity, original launch flags, or
host font state. An injected browser is therefore always `browserObserved` when no environment
attestation is supplied, and `callerAttested` when the attestation supplies those otherwise
unobservable facts; it is never `nodeVerified`. Docxodus may report directly observable browser
facts alongside either level, but does not fill missing fingerprint fields with guesses. A browser
launched by Docxodus can be `nodeVerified` only when Node verifies its executable/build and launch
configuration and every layout-relevant font comes from a verified explicit file. Use of an
unattested system font lowers the environment result to `browserObserved` and fails strict mode.

`browser` and `browserExecutablePath` are mutually exclusive. An explicit executable is snapshotted
as a stable, bounded, ordinary executable file and hashed before launch; symlinks, devices,
identity changes, relative executable paths, and security-weakening launch configurations fail.
The owned supported runtime does not opt out of Chromium's process sandbox. A host that denies that
sandbox — the AppArmor restriction on unprivileged user namespaces that Ubuntu 23.10 and later apply
by default, most commonly — is therefore a launch failure the runtime cannot resolve on the
operator's behalf, and `browser_launch_failure` names it as one: the message and remediation say the
host policy has to change rather than sending the operator to inspect the executable and its shared
libraries. An injected browser is caller-trusted and receives a fresh context, but it remains
outside the process-sandbox and background-network guarantee because those launch facts are not
observable.

Schema-v1 fidelity support is maintained on the pinned Linux x64 release image, matching the
existing visual ratchet environment. Windows x64, Linux arm64, and macOS x64/arm64 may use an explicit or injected
Chromium and are fully fingerprinted, but are `experimental` until each has a release CI job plus a
reviewed fidelity baseline. An unavailable or unproven required Chromium capability returns
`unsupported_runtime`; documentation never converts "Playwright can launch there" into a fidelity
claim.

Each conversion gets a fresh primary BrowserContext, and every auxiliary verification/print context
is fresh and batch-scoped. `docxodus/export-browser` is a public, UI-free
subpath containing the materializer entry point. `docxodus/export-assets.json` is a versioned public
asset manifest that names and hashes the worker, pagination bundle, materializer module, and WASM
base files. The materializer and worker are self-contained browser bundles apart from manifest
entries; the manifest is a closed graph containing every emitted runtime asset and transitive
dynamic import/fetch with its media type and digest. A build fails if bundle analysis finds a
runtime dependency outside that graph, and a render fails if it requests one. `@docxodus/export`
resolves only those public exports; it never discovers private `dist` paths or imports the root
barrel, whose eager editor exports are not Node-compatible. The adapter serves those
manifest-verified assets from an allowlisted synthetic local origin. All
other HTTP(S), WebSocket, service-worker, download, popup, and navigation requests are denied. It creates an owner-marked
private temporary directory when filesystem staging is necessary and removes only that directory
after success or failure. Browser processes supplied by callers are never closed.

The preferred Node origin is a Playwright-routed, unresolvable HTTPS origin with no listening
socket. Routes, WebSocket denial, service-worker blocking, permission denial, popup/download
handlers, and the exact asset/input allowlist are installed before the first page request. Allowed
requests are GET-only, query-free, digest/length/media-type checked, and bounded; the one input URL
is single-use. The report inventories every denied document request. This proves context-level
document isolation. The owned release runtime additionally strips proxy/update/crash-report
network configuration and uses the pinned offline launch policy; deployments needing a proof that
the Chromium process opened no socket enforce it with an OS network namespace or equivalent. An
injected or arbitrary explicit browser cannot claim that stronger process-level property.

Every fresh context, worker, route, and owned browser is closed on success, failure, cancellation,
or timeout; only the caller-owned `Browser` survives. Owned temporary roots are created with private
permissions and a random marker plus stable filesystem identity. Cleanup recursively removes a root
only while that identity and marker still match; an ambiguously replaced or modified root is
preserved and reported rather than risking deletion of foreign data. API byte/count limits are not
an operating-system sandbox: production deployments also bound Chromium/Node memory, CPU,
processes, file descriptors, and temporary storage with cgroups, job objects, containers, or their
platform equivalent.

The companion also exposes a framed-stdio `docxodus-export-host` for non-Node consumers. The .NET
delivery adapter starts one host for a bundle build and sends one logical request containing all
distinct render batches. Protocol v1 is a bounded sequence, not one enormous base64 JSON value: a
strict JSON control frame declares a digest/length-keyed table of unique source blobs and the stable
batch plan, followed by exact-length raw blob frames. The response uses the same scheme for
HTML/PDF/PageMap/report blobs and indexes results by batch id and artifact id. Every blob descriptor
binds its SHA-256, media type, byte length, and unique id; the receiver verifies each before use.
The host rejects blob bytes inside JSON, duplicate sources/ids/properties, missing/extra/out-of-order
frames or results, trailing bytes, and declared lengths beyond the aggregate budget before
allocation. One source used by several profile cohorts crosses the pipe once.

The host tears down at the build boundary. Each batch still owns a fresh primary context, but the
adapter never starts one Node process per key or artifact. Any batch failure makes the logical call
an error: verified partial blobs may be named only in the failed diagnostic envelope and are never
returned as nominal delivery artifacts. The protocol bounds batch/artifact/frame counts, unique
decoded input and aggregate decoded output, JSON depth/collections/strings, every raw frame, and
stderr; stdout is reserved exclusively for protocol frames. Export error code, phase, fixed `error` severity, pending resources,
part/anchor/resource, remediation, and safe detail cross this frame unchanged. The #465 delivery PR extends the #499-draft renderer diagnostic model to
retain those fields through .NET, CLI, MCP, validation evidence, and failure reports.

Chromium printing uses `printBackground: true`, `preferCSSPageSize: true`, `tagged: true`,
`outline: false`, `displayHeaderFooter: false`, scale 1, and zero browser margins. Tags receive
structural smoke coverage but are not presented as PDF/UA conformance until reading order and
semantics are independently validated.
Chromium's inferred heading outline is likewise not enabled until it is proven against Word
heading/bookmark semantics. No global Letter/A4 format is supplied. A pinned
pure-JavaScript PDF parser in the companion package verifies page count and every effective
MediaBox/CropBox origin and dimension against finalized page metadata, accounting for inherited
page attributes, rotation, and `UserUnit`, before bytes are returned. It rejects encryption,
malformed/incremental cross-reference ambiguity, and decompression beyond the PDF parser limits;
regular-expression
page counting is test scaffolding, not the production verifier. Chromium currently writes volatile
creation/modification metadata, so schema v1 guarantees deterministic layout and independently
hashes the exact returned bytes, but does not claim byte-identical PDFs across runs. Metadata is
reported with the actual PDF SHA-256 and `pdfByteDeterministic: false`; it is not patched with a
byte-level rewrite that could invalidate PDF structure. Repeatability gates compare PageMap,
physical boxes, extracted text, links, tags, and rasters while excluding declared volatile metadata.

## CLI

The package exposes one command with the same option vocabulary:

```console
docxodus convert contract.docx --to html --output contract.html \
  --document-version 12 --review-profile final --comments endnotes

docxodus convert contract.docx --to pdf --output contract.pdf \
  --document-version 12 --review-profile markup --comments margin \
  --expected-source-digest "$EXPECTED_SOURCE_DIGEST" \
  --unsupported-content strict --limit finalPages=5000 \
  --browser-executable /opt/chromium/chrome \
  --font-directory /opt/company-fonts --strict-fonts --timeout 60000 \
  --font-license-attestations company-font-licenses.json \
  --environment-attestation build-environment.json \
  --report contract.render.json --page-map contract.pages.json
```

`--font-directory` and `--limit <ExportResourceLimits-key>=<integer>` are repeatable; duplicate or
unknown limit keys are errors. `--title`, the digest, profile, comments, unsupported-content,
strict-font, timeout, and browser-executable flags map directly to the public API. Font-license and
environment attestations are canonical JSON files conforming to the public types above; unknown
fields, duplicate license attestations or host-face identities, or incomplete required facts fail
validation rather than being ignored. Multiple distinct host faces may legitimately share one font
file digest. JSON input is bounded before allocation and requires strict UTF-8, unique properties,
bounded depth/collections/strings, and closed schema discriminators. The CLI parses these files into
the corresponding Node API fields; it does not invent a second policy shape.

An explicit `--browser-executable` wins over no setting; specifying it together with a conflicting
`DOCXODUS_CHROMIUM_PATH` is an error rather than silent precedence. SIGINT/SIGTERM abort the active
operation, close owned resources, and leave no nominally complete output.

Input and output paths must differ. Before rendering, the path layer rejects syntactically duplicate
destinations, existing destinations, existing-path aliases by file identity, and known input/output
aliases after absolute normalization and real-parent resolution. Nonexistent names that collide
only under a volume's case-folding or Unicode-normalization rules may be detected only by the
no-replace commit; the preflight does not claim otherwise. The source is the bounded stable snapshot
described above. The CLI stages every requested output in its destination directory using a
cryptographically random, exclusive, private sibling; writes, rereads when needed, and fsyncs every
stage before any commit; then commits payloads, PageMap, and the report last in deterministic order.
It uses same-filesystem hard-link creation as the no-replace commit point; link creation is atomic
and fails when the target exists on supported POSIX and Windows filesystems. It fsyncs the parent
directory where the platform provides a durability primitive and removes the private name only
after the target link exists. A filesystem without safe link semantics returns `filesystem_failure`
instead of falling back to a check-then-rename race. JSON diagnostics go to the report; human
diagnostics go to stderr; PDF/HTML bytes never go through stdout implicitly. Each destination has
atomic no-replace visibility. If a later commit fails, the publisher rolls back every earlier
destination whose stable identity and content still prove ownership, fsyncs the affected parents,
and reports any destination it cannot safely remove. A process or machine crash can still expose a
prefix of links; callers requiring crash-atomic multi-file visibility publish a fresh containing
directory as `DeliveryBundleService` does.

The no-replace contract prevents accidental overwrite and path aliases; it is not a defense against
a malicious same-identity process replacing ancestor directories between path operations. Secure
callers use a private destination directory (or #465's fresh-directory publisher) and platform
access controls. A success return means every requested link exists and was synced where supported;
the report-last order prevents a visible success report from naming an artifact that this process
had not already committed.

## Review and comment profiles

Issue #444 owns their implementation, but the shared semantics are fixed here:

| Profile | Visible revision result |
|---|---|
| `final` | Inserted/destination content and final formatting; deleted/source content hidden |
| `original` | Deleted/source content and original formatting; inserted/destination content hidden |
| `markup` | Supported insertions, deletions, moves, and before/after formatting with author/date metadata |

The caller's source bytes are immutable. Unless `reviewProfileAlreadyApplied` is true, `final` and
`original` deterministically derive an isolated in-memory DOCX copy by accepting or rejecting native revisions before HTML conversion;
the report records the source and derived package identities. `markup` renders the unchanged
source. When #465 supplies an already policy-derived exact profile source, the renderer uses those
bytes directly and never applies the policy a second time.

Comments are orthogonal: `hidden`, `inline`, `endnotes`, or `margin`. A visible profile retains
range, body, author, and date in body, headers, footers, footnotes, and endnotes. Comment topology
is not drawn: a reply renders as an independent comment and resolved state is not represented,
disclosed as `comment_thread_flattened` and `comment_resolved_state_not_rendered` (see "Revision
and comment families that are not drawn" below). An unsupported story or revision/comment family
produces a structured warning naming the family; strict unsupported-content policy fails. HTML,
PDF, CLI, and #465 use these exact strings.

## Readiness and diagnostics

The complete public lifecycle phase vocabulary is:

```ts
type ExportPhase =
  | "input_validation"
  | "package_preflight"
  | "browser_launch"
  | "wasm_initialization"
  | "docx_conversion"
  | "font_loading"
  | "image_decoding"
  | "chart_svg_materialization"
  | "pagination"
  | "running_story_placement"
  | "page_tree_stability"
  | "pdf_print"
  | "output_verification"
  | "output_write"
  | "filesystem_commit"
  | "cleanup";
```

Issue #441 implements one monotonic deadline beginning at API entry. It covers owned input
snapshotting, preflight, font discovery, browser launch, rendering, printing, and output
verification (and staging/commit for the file API):

1. `input_validation`
2. `package_preflight`
3. `browser_launch`
4. `wasm_initialization`
5. `docx_conversion`
6. `font_loading`
7. `image_decoding`
8. `chart_svg_materialization`
9. `pagination`
10. `running_story_placement`
11. `page_tree_stability`
12. `pdf_print`
13. `output_verification`
14. `output_write` / `filesystem_commit` for path calls

Fonts await `document.fonts.ready`; images await `decode()`. Current charts are synchronously
generated inline SVG, so readiness validates their complete SVG tree and referenced resources;
future asynchronous producers must register promises with the same barrier. Pagination always
begins from pristine converted HTML. After pagination, the materializer creates and attaches the
sanitized fixed page tree described above before the final ResizeObserver, mutation-counter, and
geometry checks. Those signals must remain unchanged for two animation frames and at least a 100 ms
quiet interval. If any resource, cleanup, final-style, or layout signal changes the page-tree
signature, the materializer discards the mutated document, recreates it from pristine HTML, and
paginates again. It returns only after two bounded attempts produce the same sanitized page-tree
signature; only then does it measure a fresh PageMap, serialize, and perform the offline reopen
check. Otherwise it fails `pagination_failure`. This prevents cleanup or readiness from making
PageMap geometry stale.

`browser_launch` appears twice in a Node export, and in different places. On the host it is
Chromium and its isolated context, before anything is converted. In the browser materializer it is
the isolated render realm — a script-free same-origin frame — which is created only once there is
converted HTML to lay out, so it records *after* `docx_conversion` and before `font_loading`. The
offline reopen check creates a second such realm and records it under `output_verification`.

The readiness log in the render report records both sides of the barrier. The browser materializer
records the phases it runs itself, from `input_validation` through `output_verification`; the Node
host prepends the phases it owns and the materializer cannot observe from inside the page —
`browser_launch`, the materializer bootstrap under `wasm_initialization`, the closed-runtime-graph
audit under `output_verification`, and `cleanup` — so the log reads in the order the work happened.
`output_write` and `filesystem_commit` are deliberately absent from it: a report cannot record the
outcome of writing its own bytes, and their digests are already fixed by then. They remain phases
for error and timeout reporting only. A failed report carries exactly one non-complete entry, so a
teardown failure behind an already-failed render is reported through the render's failure rather
than as a second one.

Font readiness proves availability, not identity. `FontFaceSet.check()` reports whether pending
downloads have settled rather than whether a family exists — Chromium answers true for a family it
has never heard of — so each requested family's first entry is measured through advance widths
against every generic fallback. A family that matches all of them is being silently substituted for
and is recorded `missing` with an aggregate `font_unavailable` warning; one that resolves
stays `unverified`, since rendering it proves neither which file supplied it nor its version.
Because an unresolvable family is a font-policy matter governed by `strictFonts` rather than
unsupported content, it always warns and never fails the render on its own. Image and inline-SVG
findings do route through `unsupportedContent`: they warn with an omitted resource record by
default and fail closed at their own phase under `strict`. The offline reopen check applies neither
policy — a resource that materialized and then failed from the serialized HTML is a defect in the
output, not in the source, and fails `output_verification`. A resource the policy already omitted
is excluded from that check: it was reported once against the source document, and the reopened
tree necessarily reproduces the same failure.

The production path performs synchronous WASM conversion inside the existing dedicated worker and
owns the render page/context. Its total deadline can therefore terminate the worker or close the
owned context. A Promise race on the browser main thread is not considered cancellation. The same
rule applies to the caller's `AbortSignal` and the .NET cancellation token: cancellation actively
terminates owned work and produces `operation_cancelled`, while the caller-owned browser itself is
not closed. Cleanup is always attempted with a separate bounded cleanup allowance so an expired
render deadline cannot skip it or turn a verified timeout into success. Timeout errors name the
active phase and bounded pending resources.

Warnings use stable codes, severity, phase, source/part when known, message, and remediation. The
render report records requested and resolved fonts, substitutions, missing fonts, image/chart/SVG
outcomes, unsupported placeholders, external links, page metadata, readiness timings, and any
policy decision. Missing or substituted fonts warn by default and fail under `strictFonts`.

The Node adapter deterministically and boundedly discovers TTF, OTF, WOFF, and WOFF2 files in each
ordered explicit font directory, rejects symlinks/devices/escaping or changing paths, reads
their family/face metadata, hashes them, and injects license-permitted files as local webfonts into
the isolated page. That works for owned and caller-owned browsers on all supported hosts. System
fonts may be used when no directory supplies a family, but an exact file/version must come from a
caller environment attestation or the render is marked `font_environment_unverified`; strict mode
rejects it. OOXML embedded fonts are not exported until de-obfuscation and embedding-license policy
is implemented. For TTF/OTF, OS/2 `fsType` restricted-license bits forbid injection. WOFF/WOFF2 or
caller-supplied files whose embedding rights cannot be derived require an explicit caller licensing
attestation recorded in the report; absence is a policy error, not assumed permission. A license
record is scoped to its permitted HTML/PDF outputs and states whether subsetting is permitted.
When an OS/2 no-subsetting restriction applies, PDF export fails unless verification proves that
Chromium embedded the complete permitted font program; it never assumes Chromium's subset is legal.
License basis/attester strings and all font metadata are bounded and control-free. Font resolution
records include family stack, PostScript face, style, weight, stretch, glyph coverage, exact file
identity/version, substitution decision, and license evidence; these normalized records, not local
paths, enter the report and fingerprint.

The versioned browser `FontResolver` contract used by `fontResolver` has bounded request/result
records and an `AbortSignal`; Node constructs the same resolver from `fontDirectories`. A
configured root's order is policy and is preserved. Duplicate resolved roots or ambiguous faces
fail. The resolver cannot fetch URLs, and its returned font bytes are length/digest/media-type/
license checked before injection.

The Node resolver reaches the isolated page as a Playwright binding rather than as a serialized
value. That is not an implementation detail: the resolver reads font files from the host
filesystem, which the page must never be given access to, so the only thing that crosses the
boundary is one bounded request and one bounded response. The materializer therefore sees the same
`FontResolver` function shape whether the caller supplied one directly in the browser or the host
built one from directories. With no resolver configured the inventory falls back to measured
browser observation, and every resolution is recorded `unverified` from the `browser` source.

Availability in that fallback is measured, never asked of `FontFaceSet.check()`, which reports
whether pending downloads have settled rather than whether a family exists — Chromium answers true
for a family it has never heard of. `strictFonts` rejects any outcome that is not an exact,
digest-identified, license-evidenced face with complete glyph coverage, which is why a
browser-observed environment can never satisfy it.

### Revision and comment families that are not drawn

`markup` draws insertions, deletions, moves, run-level format changes, and tracked cell
insert/delete/merge — the last as tinted, struck-through or dashed cells. Two families it does not
draw travel through the projection untouched and leave no mark a reader can see, so each is
reported rather than approximated: a custom XML revision range raises
`revision_family_not_rendered`, and the block-level property revisions — paragraph, table, section
and numbering — raise `revision_property_change_not_rendered`.

That second warning counts only what is missing. `rPrChange` is a property revision the converter
does draw, so the manifest counts it separately as `runPropertyChanges` and the warning reports
`propertyChanges - runPropertyChanges`. Splitting the count in the inventory rather than
approximating it in the export is what keeps `unsupportedContent: "strict"` honest: a document
whose only property revisions are run-level format changes is drawn in full and must not fail
closed. The split costs no extra work — it is one more counter in the pass the manifest already
makes, not a second inventory pass over the package.

`final` and `original` need no such warnings. They apply the projection and then assert the derived
package retains no revisions at all, so a family that cannot be applied fails the projection instead
of passing through unseen. This is also what keeps the source package intact: the projection derives
a new package and the original bytes are never accepted, rejected, or rewritten in place.

Any visible comment profile renders comment bodies, ranges and authors, but not the topology
recorded in `commentsExtended`: a reply is drawn as an independent comment
(`comment_thread_flattened`) and a resolved comment is drawn identically to an open one
(`comment_resolved_state_not_rendered`). Both silently change what a review PDF means, which is why
they are reported instead of approximated. Comments survive the `final`/`original` projection
unchanged, so these two are raised against the source package only; the derived preflight would
otherwise report each of them twice.

All four warnings route through the `unsupportedContent` policy, so `strict` turns the first one
raised into a closed `resource_policy_failure` in `package_preflight` rather than a warning a
caller has to notice. As everywhere else in preflight, a strict export reports the first policy
breach and stops; `warn` is what enumerates every one of them in a single pass.


## Error taxonomy and limits

Public failures are `DocxodusExportError` values with one of these codes:

- `invalid_argument`
- `invalid_document`
- `source_digest_mismatch`
- `document_version_unrepresentable`
- `conversion_failure`
- `browser_launch_failure`
- `resource_policy_failure`
- `readiness_timeout`
- `operation_cancelled`
- `pagination_failure`
- `pdf_write_failure`
- `output_write_failure`
- `output_verification_failure`
- `resource_limit`
- `unsupported_runtime`
- `filesystem_failure`

`invalid_argument` covers invalid options, attestations, destinations, and API types;
`invalid_document` is reserved for malformed, encrypted, or unsupported input packages. Digest and
version failures retain their distinct codes across the host. Each error includes fixed severity
`error`, failed phase, message, remediation, and optional bounded safe detail, pending resources,
part URI, anchor, and resource. Causes/stacks may be retained on the local Node object but are never
serialized automatically. Cleanup failure does not replace a primary failure; it is attached as a
bounded secondary diagnostic.

Retaining a cause is worth nothing if no surface reads it. The CLI therefore renders the whole
chain — `cause` links and `AggregateError` members alike, so the wrapped cleanup failure is visible
too — to stderr beneath the remediation, cycle-safe, depth-bounded, and given its own character
budget so a long `detail` cannot crowd it out. Writing a cause to an operator's terminal is not
serializing it: the chain never enters `detail`, `toJSON()`, or the render report, and the rendering
is stripped of every escape sequence and control code point except the newlines the diagnostic
itself needs.

The public limit shape is:

```ts
interface ExportResourceLimits {
  compressedDocxBytes: number;
  opcEntries: number;
  expandedOpcBytes: number;
  xmlPartBytes: number;
  opcUriCharacters: number;
  opcCompressionRatio: number;
  htmlOutputBytes: number;
  pdfOutputBytes: number;
  pageMapOutputBytes: number;
  renderReportOutputBytes: number;
  pdfParserExpandedBytes: number;
  finalPages: number;
  domNodes: number;
  automaticResources: number;
  automaticResourceBytes: number;
  renderDiagnostics: number;
  fontDirectoryEntries: number;
  fontFiles: number;
  fontFileBytes: number;
  fontTotalBytes: number;
  fontRequests: number;
  fontSampleCodePoints: number;
}

interface ExportHostLimits {
  batches: number;
  artifactRequests: number;
  protocolFrames: number;
  controlFrameBytes: number;
  uniqueSourceBytes: number;
  decodedResultBytes: number;
  controlJsonDepth: number;
  controlCollectionItems: number;
  controlStringCharacters: number;
  stderrBytes: number;
}
```

`ExportOptions.limits` accepts lower per-operation values. It cannot raise the shipped defaults;
the hard-ceiling column is the maximum a future schema-v1 deployment policy may select without a
contract-version change. `timeoutMs` is the deliberate exception: callers may raise the 120-second
default through the ten-minute hard ceiling. Every value is a positive safe integer.

| Resource | Default | Hard ceiling |
|---|---:|---:|
| Compressed DOCX input | 100 MiB | 100 MiB |
| OPC entries / expanded bytes / XML part | #493 defaults: 10,000 / 1 GiB / 32 MiB | same |
| OPC URI characters / compression ratio | #493 defaults: 2,048 / 1,000 | same |
| HTML or PDF output | 256 MiB each | 512 MiB each |
| PageMap or render-report output | 256 MiB each | 512 MiB each |
| PDF parser aggregate expansion | 1 GiB | 2 GiB |
| Final pages | 10,000 | 100,000 |
| DOM nodes | 1,000,000 | 2,000,000 |
| Automatic resources / aggregate bytes | 10,000 / 256 MiB | 100,000 / 512 MiB |
| Render diagnostics | 10,000 | 100,000 |
| Font directory entries / font files | 10,000 / 1,000 | 100,000 / 10,000 |
| Font file / aggregate bytes | 32 MiB / 128 MiB | 64 MiB / 512 MiB |
| Font requests / sampled code points | 4,096 / 65,536 | 16,384 / 262,144 |
| Total deadline | 120 seconds | 10 minutes |

The separately versioned `export-host-limits/v1` control contract defaults respectively to 64
batches, 256 artifact requests, 1,024 frames, an 8 MiB control frame, 512 MiB of unique source
bytes, 1 GiB of decoded result bytes, JSON depth 32, 4,096 collection items, 1 MiB of aggregate
control-string characters, and 1 MiB of stderr. Its hard ceilings are respectively 256, 1,024,
4,096, 32 MiB, 1 GiB, 2 GiB, 64, 16,384, 4 MiB, and 4 MiB. The protocol identifier and effective
limits are in the first control frame and mismatches fail before blobs are accepted. Per-blob limits
remain the effective export limits. Aggregate limits are sized so one otherwise valid worst-case
HTML+PDF+PageMap+report batch can traverse the host; additional cohorts share the remaining budget.

The checked-in export-limits v1 contract is the single source for the TypeScript options, WASM
boundary, Node preflight, and CLI validation. Its compressed-input value is 104,857,600 bytes,
matching the existing WASM safety boundary; #438 replaces the private duplicate constant with the
generated/shared contract. Node checks the limit before worker transfer, WASM checks the same value
defensively, and both report `resource_limit`. Raising that ceiling requires a versioned contract
change and memory/fidelity evidence, not a Node-only override.

The effective `opcEntries`, `expandedOpcBytes`, `xmlPartBytes`, `opcUriCharacters`, and
`opcCompressionRatio` values are passed into #493 manifest generation so a caller's lower ceiling
limits the inspection itself; generating with defaults and rejecting afterward is not enforcement.
#493 entry `size` and `compressedSize` values are canonical non-negative base-10 strings because
ZIP64 exceeds JavaScript's safe integer range. Export validates them and compares/sums with
`BigInt` (or equivalent checked decimal arithmetic), short-circuiting at the effective limit; it
never converts them to `number` or concatenates strings. The manifest is validated as a closed
schema, including digest algorithms, enum discriminators, duplicate properties, package kind,
main-document identity, findings, and nullable digests, before any field drives policy. This is the
required compatibility boundary with the hardened #493 contract.

Limit failures never return a nominally complete artifact. Supported package resources that fail
to decode are errors. Unsupported-but-representable content uses a visible placeholder plus warning
by default and fails under `unsupportedContent: "strict"`; omission is never a successful policy.
Automatic external resources are forbidden. User-activated external hyperlinks may remain and are
inventoried, but export never follows them.

Limits are checked before copying/decoding where possible and incrementally while constructing DOM,
resource inventories, PageMap, report, PDF parser state, and framed output; checking only after an
unbounded allocation is a defect. `pdf_write_failure` means Chromium failed to produce PDF bytes. Byte-returning library calls do no
destination write. The path/CLI layer uses `output_write_failure` for an
HTML/PDF/report/PageMap destination and `filesystem_failure` for stage or commit mechanics.

## Render report schema

`RenderReport` is canonical JSON with schema
`https://docxodus.dev/schemas/render/render-report/v1`. The initial public shape is:

```ts
interface RenderWarning {
  code: string;
  severity: "warning";
  phase: ExportPhase;
  message: string;
  remediation: string;
  detail?: string;
  partUri?: string;
  anchorId?: string;
  resource?: string;
}

interface RenderReportBase {
  schema: "https://docxodus.dev/schemas/render/render-report/v1";
  schemaVersion: 1;
  source: { rawPackageBytesDigest: Sha256Hex; byteLength: number; documentVersion: number };
  derivedProfileSource?: { rawPackageBytesDigest: Sha256Hex; byteLength: number };
  options: {
    reviewProfile: ReviewProfile;
    reviewProfileAlreadyApplied: boolean;
    commentProfile: CommentProfile;
    title: string;
    outputs: readonly ("html" | "pdf")[];
    layoutDigest: Sha256Hex;
    runtimePolicyDigest: Sha256Hex;
    policy: {
      unsupportedContent: "warn" | "strict";
      strictFonts: boolean;
      timeoutMs: number;
      limits: Readonly<ExportResourceLimits>;
    };
  };
  readiness: ReadonlyArray<{
    phase: ExportPhase;
    status: "complete" | "failed" | "cancelled";
    elapsedMs: number;
    pending: readonly string[];
  }>;
  fontIdentity?: FontConfigurationIdentity;
  fonts: readonly FontResolution[];
  resources: readonly ResourceOutcome[];
  unsupportedContent: readonly UnsupportedContentOutcome[];
  warnings: readonly RenderWarning[];
}

interface ExportRuntimeObservedFacts {
  runtimeKind: "browser" | "nodeChromium";
  playwrightVersion?: string;
  browserProduct?: string;
  browserBuild?: string;
  executableSha256?: Sha256Hex;
  launchFlags?: readonly string[];
  operatingSystem?: string;
  architecture?: string;
  locale: string;
  timezone: string;
  viewport: readonly [number, number];
  deviceScaleFactor: number;
  media: {
    colorScheme: "light" | "dark" | "no-preference";
    reducedMotion: "reduce" | "no-preference";
    forcedColors: "active" | "none";
    printMedia: true;
  };
  networkIsolation: "ownedProcessRestricted" | "contextRestricted";
}

interface ExportRuntimeAttestationEvidence {
  chromiumProduct: string;
  chromiumBuild: string;
  executableSha256: Sha256Hex;
  launchFlags: readonly string[];
  hostFontsDigest: Sha256Hex;
  basis: string;
}

interface CompleteRenderReport extends RenderReportBase {
  status: "complete";
  fontIdentity: FontConfigurationIdentity;
  environment: {
    rendererFingerprint: Sha256Hex;
    verification: "nodeVerified" | "browserObserved" | "callerAttested";
    fidelityTier: "releaseBaselined" | "experimental" | "unbaselined";
    observed: ExportRuntimeObservedFacts;
    attested?: ExportRuntimeAttestationEvidence;
    attestationDigest?: Sha256Hex;
  };
  pages: ReadonlyArray<{
    pageNumber: number;
    pageInSection: number;
    pageName: string;
    width: number;
    height: number;
    sectionIndex?: number;
  }>;
  bindings: {
    pageMapDigest: Sha256Hex;
    htmlDigest?: Sha256Hex;
    pdfDigest?: Sha256Hex;
    artifactRequestIds: readonly string[];
    pdfByteDeterministic?: false;
    volatilePdfMetadata?: Readonly<Record<string, string>>;
  };
}

interface FailedRenderReport extends RenderReportBase {
  status: "failed";
  failure: {
    code: DocxodusExportErrorCode;
    severity: "error";
    phase: ExportPhase;
    message: string;
    remediation: string;
    detail?: string;
    pending?: readonly string[];
    partUri?: string;
    anchorId?: string;
    resource?: string;
  };
  environment?: Partial<CompleteRenderReport["environment"]> & {
    verification: "nodeVerified" | "browserObserved" | "callerAttested";
  };
  partial?: {
    pages?: CompleteRenderReport["pages"];
    bindings?: Partial<CompleteRenderReport["bindings"]>;
  };
  unavailable: ReadonlyArray<{
    field: "environment.rendererFingerprint" | "bindings.pageMapDigest"
      | "bindings.htmlDigest" | "bindings.pdfDigest";
    reasonCode: "notReached" | "notRequested" | "failedVerification" | "discardedOnFailure";
    detail: string;
  }>;
}

type RenderReport = CompleteRenderReport | FailedRenderReport;
```

Page widths/heights and PageMap geometry are finite points. `callerAttested` requires a bounded
attestation whose observable runtime fields match and whose executable digest is present; otherwise
verification remains `browserObserved`. Verification strength is separate from `fidelityTier`: a
fully node-verified Windows render is still `experimental` until its ratchet is baselined. The
observed/attested facts or their canonical referenced evidence make the fingerprint reproducible;
a one-way hash alone is not presented as a recipe.

Resolved report options always contain `reviewProfileAlreadyApplied`, `title`, and `outputs`:
omission/defaulting exists only at API input. `outputs` is unique and in canonical `html`, `pdf`
order. The browser HTML call reports `["html"]`, each Node convenience call reports its one output,
and a metadata-only batch reports `[]`. Complete-report schema conditionals require `htmlDigest` and
`pdfDigest` exactly when their output is selected and permit PDF volatility fields only with PDF.
On failure, `unavailable` contains exactly one reason for every unavailable fingerprint/PageMap or
selected payload binding; unselected payloads are represented by `notRequested` and cannot appear
in `partial.bindings`. This makes a standalone report sufficient to validate its own bindings.

Schema conditionals require `derivedProfileSource` exactly when `final` or `original` was derived by
this operation, forbid it for `markup` and already-applied input, and require the derived digest to
differ from the source digest whenever revision processing changed bytes. `fontIdentity` is
required only on a complete report; the base declaration lets a safely reached failure retain it
without inventing one.

A complete report has only `complete` readiness entries and warning-severity diagnostics. A failed
report has the ordered completed entries followed by exactly one terminal `failed` or `cancelled`
entry and the matching error-severity failure envelope; cleanup diagnostics cannot change that
primary status.

Browser/WASM callers are not falsely restricted to Chromium: facts unavailable through the browser
privacy boundary are omitted, and such a report is `browserObserved` and `unbaselined` unless a
separately supported baseline says otherwise. Node export accepts Chromium only. `nodeVerified`
requires all optional Node/Chromium identity fields in `observed`; `callerAttested` requires the
closed attested evidence above and the digest of the complete versioned input attestation. An
attestation without `executableSha256` may still be inventoried, but cannot raise verification above
`browserObserved`.

#438 lands the writer and the complete closed JSON Schema above, including the resolved
`reviewProfileAlreadyApplied`, final font-resolution definitions, all font/sidecar limit keys, and
environment evidence needed by #442/#444. Those later issues populate the reserved fields; they do
not change the meaning, required fields, or closed enums under the same v1 IDs. The #501--#508 stack
must either bring that final shape forward before #501 merges or use an explicitly draft schema ID
until #508 freezes v1. Compatibility tests validate old-writer/new-schema and new-writer/frozen-v1
cases, and reject reports that only a later private schema accepts.

The report is a separate sidecar;
HTML and PDF do not embed it, avoiding a digest cycle while the report binds their bytes. Failed
attempts retain a report when execution reached reporting safely, including CLI runs with an
explicit `--report` destination, but never a `complete` artifact result. The discriminated failed
shape records why a fingerprint or artifact binding is unavailable when failure precedes layout;
it never fabricates a PageMap digest or renderer identity merely to satisfy the schema. Failures
that occur before a valid source/options base exists still use the versioned host error envelope but
do not forge a schema-valid report.

All canonical JSON in this contract uses RFC 8785 JCS over strict UTF-8 with no BOM, duplicate
properties, non-finite numbers, or unpaired surrogates. `pageMapDigest` is plain SHA-256 of the
frozen canonical PageMap bytes; HTML/PDF digests are plain SHA-256 of the exact returned
UTF-8/opaque bytes. Every digest of JSON policy/material is SHA-256 over its ASCII domain tag, one
zero byte, and then its JCS bytes. The exact schema-v1 tags are
`docxodus:layout-options:v1`, `docxodus:runtime-policy:v1`,
`docxodus:font-configuration:v1`, `docxodus:environment-attestation:v1`, and
`docxodus:renderer-fingerprint:v1`; there is no newline or length prefix. Source-package digests
retain #493's plain-byte definition. JavaScript and .NET share golden vectors covering empty
material, property order, Unicode, numeric point values, optional fields, and empty arrays before
#501 is accepted.

## Renderer fingerprint

The renderer fingerprint is a canonical SHA-256 identity over every layout-relevant input:

- Docxodus package/core/WASM and paginator contract versions;
- render-report and PageMap schema versions;
- Playwright-core version plus Chromium product/build, launch/headless flags, operating system,
  architecture, locale, timezone, viewport, device scale, and media settings;
- a canonical layout/materialization-options digest including the resolved title,
  review/comment profiles, `reviewProfileAlreadyApplied`, and every pagination option;
- the configured runtime-policy digest, including effective limits, asset/runtime selection,
  isolation policy, ordered font roots, and timeout;
- the font-configuration digest; and
- sorted requested-to-resolved font family, file identity, and version records.

The source package digest and document version are schema-bound beside the fingerprint. For a
Docxodus-launched browser with verified files and fonts, Node collects the authoritative facts and
passes the completed fingerprint into browser layout as the PageMap token. Browser-only and
unattested injected-browser callers receive a fingerprint over browser-observable facts plus the
explicit `browserObserved` verification level. Attested injected environments are
`callerAttested`, never `nodeVerified`; the report distinguishes observed fields from attested
fields. A caller uses the report's bounded observed/attested facts and font records—not the one-way
fingerprint itself—to provision the same runtime. Recomputing the fingerprint then verifies that
recipe. Any change in a bound input is visible even when output bytes happen to match.

## Supported fidelity

"Supported" means the automated gates below pass and no required resource is unavailable. It does
not mean arbitrary Word content is silently approximated.

| Capability | Release requirement |
|---|---|
| Page geometry | logical page count equals PDF count; each effective PDF MediaBox/CropBox origin and dimension, after inherited attributes, rotation, and `UserUnit`, is within 0.5 pt of its PageMap page |
| Selectable text | profile-visible text is searchable and extracted in predictable story order |
| Links | supported external and internal targets survive HTML/PDF inspection |
| Images | supported embedded raster images decode; failure is reported |
| SVG/charts | supported vector output remains vector and signals readiness; unsupported forms are reported |
| Running stories | headers, footers, footnotes, endnotes, page fields, section ownership, and continuations match the finalized page tree |
| Mixed sections | portrait/landscape, Letter/A4, explicit/section/continuous breaks, and supported columns retain order and physical geometry |
| Revisions/comments | exact shared profile semantics and warnings for unsupported families |
| Fonts | every requested family has an exact resolution or recorded substitution/missing decision |
| Visual fidelity | generated PDFs pass #443's committed corpus ratchet and environment contract |

Explicitly unsupported in schema v1 are active content, macros, OLE activation, embedded video,
external linked images/fonts/stylesheets, arbitrary HTML (`altChunk`), and any converter placeholder
family not promoted by a follow-on issue. WMF/EMF, unsupported Office Math, form fields, and other
opaque content remain visible placeholders with warnings when requested; they are never counted as
faithful rendering. Native SVG image parts are not claimed until their current placeholder path is
replaced and covered by readiness/vector tests. Password-protected/encrypted or malformed packages
fail before rendering.

Eligible text-only footnote paragraphs continue within the paragraph at deterministic Unicode-safe
boundaries, preserving inline structure, identities, ordering, and PageMap fragments. Structurally
indivisible note content retains the typed clipped-content failure rather than shipping a complete
result with missing text.

Package preflight, raw source identity, revision/comment/media inventory, safety findings, and ZIP
resource ceilings reuse #493's `PackageManifestGenerator` contract through its WASM surface. The
export path does not add a second partial ZIP inspector. The export stack remains rebased on the
hardened #493 contract and proves its algorithm-labelled digests plus decimal-string ZIP64 sizes at
the #501 acceptance boundary; code written against the former numeric entry-size shape cannot land.

## Follow-on reconciliation

| Issue / PR | Contract contribution |
|---|---|
| #438 / #501 | isolated final page-tree materialization, offline serializer, complete v1 browser types/schema/docs/tests |
| #439 / #502 | `@docxodus/export`, CLI, local runtime serving, bundle/PDF bytes and typed failures |
| #440 / #503 | mixed-section effective PDF-box and sequencing proof |
| #489 / #504 | lossless mid-paragraph footnote continuation and complete note text/PageMap coverage |
| #441 / #505 | phased readiness barrier, quiet interval, cancellation, delayed-resource tests |
| #442 / #506 | font directories/resolver, license policy, resolution evidence, fingerprint integration |
| #443 / #507 | generated-PDF raster/text/link ratchet and reproducibility documentation |
| #444 / #508 | shared final/original/markup and hidden/inline/endnotes/margin profiles |
| #465 / delivery PR (drafted as #499, closed) | consume #439's exact batch session through the delivery seam |

These remain separate PRs, but their acceptance order preserves the contract rather than merely the
current branch stack. #501 freezes the complete schema-v1 shapes used through #508 (or uses an
explicit draft ID until that is true). #507 may add the ratchet earlier, but it is not the release
gate until #508 has landed and the release candidate reruns every advertised review/comment profile,
including `original` and already-applied final/original sources. No published v1 may advertise a
profile that returns `unsupported_runtime` in its supported release runtime.

The #465 delivery PR is accepted only after it exposes the bundle-level `RenderBatchesAsync` seam above, validates
both group-context digests, represents an unsafe .NET document version as the closed
`document_version_unrepresentable` reason, and preserves the structured diagnostic envelope. It may
not loop the legacy `RenderAsync` API or use cross-call cache state. The #501--#508 stack remains
rebased on hardened #493, and each PR validates the frozen writer/schema pair before the next PR
extends behavior.
