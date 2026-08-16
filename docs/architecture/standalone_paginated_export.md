# Standalone paginated HTML and PDF export

Status: accepted design for epic #434 and issue #437.

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
type ReviewProfile = "final" | "original" | "markup";
type CommentProfile = "hidden" | "inline" | "endnotes" | "margin";

interface PaginatedHtmlOptions {
  documentVersion?: number;
  expectedSourceDigest?: string;
  reviewProfile: ReviewProfile;
  commentProfile: CommentProfile;
  title?: string;
  unsupportedContent?: "warn" | "strict";
  strictFonts?: boolean;
  timeoutMs?: number;
  limits?: Partial<ExportResourceLimits>;
}

interface PaginatedRenderMetadata {
  pageCount: number;
  pageMap: PageMap;
  renderReport: RenderReport;
  warnings: RenderWarning[];
  rendererFingerprint: string;
}

interface PaginatedHtmlResult extends PaginatedRenderMetadata {
  html: string;
}

convertDocxToPaginatedHtml(
  document: File | Uint8Array,
  options: PaginatedHtmlOptions,
): Promise<PaginatedHtmlResult>;
```

`reviewProfile` and `commentProfile` are required. The operation obtains the authoritative raw
source-package digest from #493 preflight. `documentVersion` is optional provenance; standalone
conversion uses 0 for its immutable input snapshot when omitted, while session/delivery callers
supply their exact version. A supplied value must be a non-negative JavaScript safe integer. The
returned PageMap is the existing portable v1 contract, measured from the final page tree and bound
to the source digest by the render report. A #465 request whose .NET `long` version is outside that
range returns typed `document_version_unrepresentable` unavailability rather than rounding it.
`expectedSourceDigest`, when supplied, must equal #493's recomputed canonical digest. The default
unsupported-content policy is `warn`, which means a visible placeholder plus a structured warning;
`strict` fails instead. Resource limits use the shared defaults below and may only be lowered.

The browser API lays out in an attached, offscreen, same-origin iframe with real layout and a usable
`defaultView`; an unattached `createHTMLDocument()` is not a render surface. Before converted source
HTML is parsed, the iframe is bootstrapped with a deny-script/connect/object/navigation CSP and a
sandbox that retains same-origin DOM access but omits script execution. It remains laid out using a
fixed offscreen viewport (not `display: none`) and is destroyed after verification/serialization.
The Node adapter supplies the equivalent fresh Playwright page and request-denial policy.
Pagination must use the staging element's `ownerDocument` and its `defaultView` for element
creation, ranges, tree walking, computed style, animation frames, and observers. Replacing only
`document.createElement` would leave a cross-document layout bug. This is a supported isolation
boundary, not a hidden copy of the pagination algorithm. The API treats caller bytes as immutable:
it synchronously copies any `Uint8Array` at API entry, or reads a `File` into a new owned buffer,
before its cancellable worker transfer. The transferred buffer may detach; the caller's object must
not.

## Standalone HTML contract

The result is a complete `<!doctype html>` document containing the finalized page boxes, not the
staging tree. It has these invariants:

- the measurement staging tree, registries, scripts, transient readiness markers, selection, and
  viewer controls are absent; synthetic page numbers, `will-change`, containment hints, and other
  measurement-only inline styles are removed;
- document CSS is inline, supported package images are data URLs, and no CSS import, stylesheet,
  script, image, media, or font URL can initiate a network request;
- external hyperlinks may remain as inert user-activated links and are recorded in the report;
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

Serialization audits every URL-bearing element and CSS rule. An unresolved automatic resource is
a typed warning or error according to policy; it is never silently retained as an online
dependency. Fixed page boxes preserve page count and physical geometry across hosts; exact visual
and text-flow reproduction additionally requires the same fingerprinted font environment or
license-permitted embedded webfonts. Schema v1 does not claim that arbitrary system or
OOXML-embedded fonts are portable.

## Node API, batch session, and browser runtime

Issue #439 adds the Node-only surface:

```ts
interface FontLicenseAttestation {
  fileSha256: string;
  embeddingPermitted: true;
  basis: string;
  attester?: string;
}

interface RenderEnvironmentAttestation {
  chromiumProduct: string;
  chromiumBuild: string;
  executableSha256?: string;
  launchFlags: string[];
  hostFonts: Array<{
    family: string;
    style: string;
    weight: number;
    fileSha256: string;
    version: string;
  }>;
  basis: string;
}

interface NodeExportRuntime {
  browser?: Browser;                 // caller-owned; never closed by Docxodus
  browserExecutablePath?: string;    // Docxodus launches and owns this browser
  fontDirectories?: string[];
  fontLicenseAttestations?: FontLicenseAttestation[];
  environmentAttestation?: RenderEnvironmentAttestation;
}

interface PdfExportResult extends PaginatedRenderMetadata {
  pdf: Uint8Array;
}

convertDocxToPdf(
  document: Uint8Array,
  options: PaginatedHtmlOptions & NodeExportRuntime,
): Promise<PdfExportResult>;

convertDocxToStandaloneHtml(
  document: Uint8Array,
  options: PaginatedHtmlOptions & NodeExportRuntime,
): Promise<PaginatedHtmlResult>;

interface RenderBatchResult extends PaginatedRenderMetadata {
  html?: string;
  pdf?: Uint8Array;
}

renderDocxArtifacts(
  document: Uint8Array,
  options: PaginatedHtmlOptions & NodeExportRuntime & {
    outputs: Array<"html" | "pdf">;
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
  options: PaginatedHtmlOptions & NodeExportRuntime,
): Promise<RenderFileResult>;
```

The Node `convertDocxToStandaloneHtml` and `convertDocxToPdf` functions are convenience projections
over `renderDocxArtifacts`; the browser `convertDocxToPaginatedHtml` remains the shared materializer
they drive. `outputs` controls only the optional HTML/PDF payload properties. PageMap and report are
always returned as verification metadata; the caller decides whether to persist them. The batch
function is the integration surface for #465: one invocation owns one page/context and returns
every requested output from that page tree. `renderDocxFile` infers its byte outputs from the
destinations, requires at least one destination, and uses the same batch before creating each
requested file with an atomic no-replace commit. Separate destination files are not an all-or-none
transaction; callers requiring that property publish a fresh containing directory, as
`DeliveryBundleService` already does.

Every Node byte API synchronously copies its `Uint8Array` at function entry, before preflight,
browser selection, or any `await`, and uses only that owned snapshot for digesting, transport, and
rendering. `renderDocxFile` reads one owned byte snapshot through a single opened file handle and
rejects a source whose size/identity changes while it is being read. Later caller-buffer or path
changes therefore cannot split source identity from rendered bytes.

Before the session starts, #493 preflight recomputes the source `rawPackageBytesDigest`. A caller
may provide an expected digest; mismatch is `invalid_document`, not a warning. The delivery adapter
groups requested render kinds by source digest, safe document version, review/comment profiles,
canonical layout-options digest, and runtime policy, invokes one batch per distinct key, and
projects the results. It renders the exact final, policy-baseline, or proven review bytes selected
by `DeliveryBundleService`; it never regenerates a supposedly equivalent profile source.

PR #499 replaces the draft single-item delivery seam with an explicit batch seam before it leaves
draft:

```csharp
public sealed record DeliveryRenderBatchContext(
    DeliveryReviewProfile ReviewProfile,
    DeliveryCommentProfile CommentProfile,
    VerificationDigest LayoutOptionsDigest,
    VerificationDigest RuntimePolicyDigest);

public interface IDeliveryArtifactRenderer
{
    DeliveryRendererCapabilities Capabilities { get; }

    DeliveryRenderBatchContext DescribeBatch(
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile);

    ValueTask<IReadOnlyDictionary<string, DeliveryRenderResult>> RenderBatchAsync(
        DeliveryRenderBatchContext context,
        IReadOnlyList<DeliveryRenderRequest> requests,
        CancellationToken cancellationToken = default);
}
```

`DeliveryBundleService` validates every plan first, groups render requests by exact source digest,
document version, review/comment profiles, layout-options digest, and configured runtime policy,
then invokes `RenderBatchAsync` exactly once per group. At build start it calls the renderer's pure
`DescribeBatch` once for every requested review/comment pair. The returned group-specific layout
digest includes those profiles and every pagination option; the runtime-policy digest covers the
fixed browser/resource/font policy. Repeating `DescribeBatch` for the same pair on one renderer
must be identical. The service uses the complete context in the group key and passes it unchanged;
the renderer rejects mismatched profiles or digests. Those digests also bind the report, so grouping
cannot depend on hidden mutable adapter state. The dictionary must contain exactly one result for
every requested artifact id and no extras. A genuine one-artifact render is a one-item batch; there
is no `RenderAsync` compatibility loop for a multi-artifact group. The production adapter implements
this contract directly, while test/evaluation renderers are updated to return one-item or
deterministic fixture batches. No adapter may claim shared PageMap/report provenance for results
produced by separate render sessions.

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

Schema-v1 fidelity support is maintained Node.js on pinned Linux x64, matching the existing visual
ratchet environment. Windows, macOS, Linux arm64, and macOS arm64 may use an explicit or injected
Chromium and are fully fingerprinted, but are `experimental` until each has a release CI job plus a
reviewed fidelity baseline. An unavailable or unproven required Chromium capability returns
`unsupported_runtime`; documentation never converts "Playwright can launch there" into a fidelity
claim.

Each conversion gets a fresh BrowserContext. `docxodus/export-browser` is a public, UI-free
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

The companion also exposes a framed-stdio `docxodus-export-host` for non-Node consumers. The .NET
delivery adapter starts one host for a bundle build, sends one length-bounded envelope containing
all distinct render batches, reads one length-bounded envelope with results keyed by batch id and
artifact id, and tears it down at the build boundary. Each batch still owns a fresh context, but
the adapter never starts one Node process per key or artifact. The host rejects duplicate, missing,
or extra ids. Export error code, phase, severity, pending resources, part/anchor, remediation, and
safe detail cross this frame unchanged. PR #499 extends its current renderer diagnostic model to
retain those fields through .NET, CLI, MCP, validation evidence, and failure reports.

Chromium printing uses `printBackground: true`, `preferCSSPageSize: true`, `tagged: true`,
`outline: false`, `displayHeaderFooter: false`, scale 1, and zero browser margins. Tags receive
structural smoke coverage but are not presented as PDF/UA conformance until reading order and
semantics are independently validated.
Chromium's inferred heading outline is likewise not enabled until it is proven against Word
heading/bookmark semantics. No global Letter/A4 format is supplied. A pinned
pure-JavaScript PDF parser in the companion package verifies page count and every
MediaBox/CropBox against finalized page metadata before bytes are returned; regular-expression
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
fields, duplicate font digests, or incomplete required facts fail validation rather than being
ignored. The CLI parses these files into the corresponding Node API fields; it does not invent a
second policy shape.

Input and output paths must differ. Before rendering, the path layer rejects duplicate destinations,
existing destinations, and input/output aliases after absolute normalization, existing-path file-id
checks, and real-parent resolution for new targets. The CLI writes and fsyncs an owner-marked
private sibling file, then uses a same-filesystem hard-link create as the commit point; link creation
is atomic and fails when the target exists on POSIX and Windows. It removes the private name only
after the target link exists. A filesystem without safe link semantics returns `filesystem_failure`
instead of falling back to a check-then-rename race. JSON diagnostics go to the report; human
diagnostics go to stderr; PDF/HTML bytes never go through stdout implicitly. Each destination has
atomic no-replace visibility, but several destinations can be partially visible after a crash or
later commit failure; the returned error lists committed destinations and the CLI never claims
multi-file transactionality.

## Review and comment profiles

Issue #444 owns their implementation, but the shared semantics are fixed here:

| Profile | Visible revision result |
|---|---|
| `final` | Inserted/destination content and final formatting; deleted/source content hidden |
| `original` | Deleted/source content and original formatting; inserted/destination content hidden |
| `markup` | Supported insertions, deletions, moves, and before/after formatting with author/date metadata |

The caller's source bytes are immutable. For the standalone API, `final` and `original` may derive
an isolated in-memory DOCX copy by accepting or rejecting native revisions before HTML conversion;
the report records the source and derived package identities. `markup` renders the unchanged
source. When #465 supplies an already policy-derived exact profile source, the renderer uses those
bytes directly and never applies the policy a second time.

Comments are orthogonal: `hidden`, `inline`, `endnotes`, or `margin`. A visible profile retains
range, body, the ordered reply tree, author, date, and resolved state in body, headers, footers,
footnotes, and endnotes. Missing extended-comment metadata is represented as unknown rather than
invented. An unsupported story or revision/comment family produces a structured warning naming the
family and owning part; strict unsupported-content policy fails. HTML, PDF, CLI, and #465 use these
exact strings.

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

Issue #441 implements a single deadline across the rendering phases:

1. `wasm_initialization`
2. `docx_conversion`
3. `font_loading`
4. `image_decoding`
5. `chart_svg_materialization`
6. `pagination`
7. `running_story_placement`
8. `page_tree_stability`
9. `pdf_print`

The coordinator explicitly loads every computed font family/sample and then awaits
`document.fonts.ready`; images decode in parallel and produce structured complete/failed resource
outcomes. Chart/SVG producers use `data-docxodus-materialization`, `-state`, and `-id` attributes;
the barrier waits for all `pending` producers and validates the resulting SVG. Pagination returns
an explicit ready result with page count and inventory diagnostics. Pagination always
begins from pristine converted HTML. After pagination, the materializer creates and attaches the
sanitized fixed page tree described above before the final ResizeObserver, mutation-counter, and
geometry checks. Those signals must remain unchanged for two animation frames and at least a 100 ms
quiet interval. If any resource, cleanup, final-style, or layout signal changes the page-tree
signature, the materializer discards the mutated document, recreates it from pristine HTML, and
paginates again. It allows at most three attempts and returns only after two consecutive stable
attempts produce the same sanitized page-tree signature; only then does it measure a fresh PageMap
and serialize.
The Node PDF path re-runs the font, image, graphic, and stable-tree barriers in the exact reopened
document Chromium prints and verifies that its page count is unchanged. Otherwise export fails
with the exact phase and pending resources. This prevents cleanup or readiness from making
PageMap geometry stale.

The production path performs synchronous WASM conversion inside the existing dedicated worker and
owns the render page/context. Its total deadline can therefore terminate the worker or close the
owned context. Every browser phase also owns an abort signal that disconnects observers and cancels
timed waits. Progress crosses the browser/Node boundary so the outer watchdog reports the active
phase and its current pending resources instead of masking it as a generic Node timeout.

Warnings use stable codes, severity, phase, source/part when known, message, and remediation. The
render report records requested and resolved fonts, substitutions, missing fonts, image/chart/SVG
outcomes, unsupported placeholders, external links, page metadata, readiness timings, and any
policy decision. Missing or substituted fonts warn by default and fail under `strictFonts`.

The Node adapter discovers TTF, OTF, WOFF, and WOFF2 files in each explicit font directory, reads
their family/face metadata, hashes them, and injects license-permitted files as local webfonts into
the isolated page. That works for owned and caller-owned browsers on all supported hosts. System
fonts may be used when no directory supplies a family, but an exact file/version must come from a
caller environment attestation or the render is marked `font_environment_unverified`; strict mode
rejects it. OOXML embedded fonts are not exported until de-obfuscation and embedding-license policy
is implemented. For TTF/OTF, OS/2 `fsType` restricted-license bits forbid injection. WOFF/WOFF2 or
caller-supplied files whose embedding rights cannot be derived require an explicit caller licensing
attestation recorded in the report; absence is a policy error, not assumed permission.

## Error taxonomy and limits

Public failures are `DocxodusExportError` values with one of these codes:

- `invalid_document`
- `conversion_failure`
- `browser_launch_failure`
- `resource_policy_failure`
- `readiness_timeout`
- `pagination_failure`
- `pdf_write_failure`
- `output_write_failure`
- `output_verification_failure`
- `resource_limit`
- `unsupported_runtime`
- `filesystem_failure`

Each includes the failed phase, safe detail, and remediation. Causes may be retained on the Node
object but are not serialized automatically. The public limit shape is:

```ts
interface ExportResourceLimits {
  compressedDocxBytes: number;
  opcEntries: number;
  expandedOpcBytes: number;
  xmlPartBytes: number;
  htmlOutputBytes: number;
  pdfOutputBytes: number;
  finalPages: number;
  domNodes: number;
  automaticResources: number;
  automaticResourceBytes: number;
}
```

`PaginatedHtmlOptions.limits` accepts lower per-operation values, while `timeoutMs` controls the
single total deadline. Schema v1 uses these defaults; callers may lower them but cannot raise a
hard ceiling:

| Resource | Default | Hard ceiling |
|---|---:|---:|
| Compressed DOCX input | 100 MiB | 100 MiB |
| OPC entries / expanded bytes / XML part | #493 defaults: 10,000 / 1 GiB / 32 MiB | same |
| HTML or PDF output | 256 MiB each | 512 MiB each |
| Final pages | 10,000 | 100,000 |
| DOM nodes | 1,000,000 | 2,000,000 |
| Automatic resources / aggregate bytes | 10,000 / 256 MiB | 100,000 / 512 MiB |
| Total deadline | 120 seconds | 10 minutes |

The checked-in export-limits v1 contract is the single source for the TypeScript options, WASM
boundary, Node preflight, and CLI validation. Its compressed-input value is 104,857,600 bytes,
matching the existing WASM safety boundary; #438 replaces the private duplicate constant with the
generated/shared contract. Node checks the limit before worker transfer, WASM checks the same value
defensively, and both report `resource_limit`. Raising that ceiling requires a versioned contract
change and memory/fidelity evidence, not a Node-only override.

Limit failures never return a nominally complete artifact. Supported package resources that fail
to decode are errors. Unsupported-but-representable content uses a visible placeholder plus warning
by default and fails under `unsupportedContent: "strict"`; omission is never a successful policy.
Automatic external resources are forbidden. User-activated external hyperlinks may remain and are
inventoried, but export never follows them.

`pdf_write_failure` means Chromium failed to produce PDF bytes. Byte-returning library calls do no
destination write. The path/CLI layer uses `output_write_failure` for an
HTML/PDF/report/PageMap destination and `filesystem_failure` for stage or commit mechanics.

## Render report schema

`RenderReport` is canonical JSON with schema
`https://docxodus.dev/schemas/render/render-report/v1`. The initial public shape is:

```ts
interface RenderWarning {
  code: string;
  severity: "warning" | "error";
  phase: ExportPhase;
  message: string;
  remediation: string;
  partUri?: string;
  anchorId?: string;
  resource?: string;
}

interface RenderReportBase {
  schema: "https://docxodus.dev/schemas/render/render-report/v1";
  schemaVersion: 1;
  source: { rawPackageBytesDigest: string; byteLength: number; documentVersion: number };
  derivedProfileSource?: { rawPackageBytesDigest: string; byteLength: number };
  options: {
    reviewProfile: ReviewProfile;
    commentProfile: CommentProfile;
    layoutDigest: string;
  };
  readiness: Array<{
    phase: ExportPhase;
    status: string;
    elapsedMs: number;
    pending: string[];
  }>;
  fonts: FontResolution[];
  resources: ResourceOutcome[];
  unsupportedContent: UnsupportedContentOutcome[];
  warnings: RenderWarning[];
}

interface CompleteRenderReport extends RenderReportBase {
  status: "complete";
  environment: {
    rendererFingerprint: string;
    verification: "nodeVerified" | "browserObserved" | "callerAttested";
  };
  pages: Array<{
    pageNumber: number;
    width: number;
    height: number;
    sectionIndex?: number;
  }>;
  bindings: {
    pageMapDigest: string;
    htmlDigest?: string;
    pdfDigest?: string;
    artifactRequestIds: string[];
    pdfByteDeterministic?: false;
    volatilePdfMetadata?: Record<string, string>;
  };
}

interface FailedRenderReport extends RenderReportBase {
  status: "failed";
  failure: {
    code: DocxodusExportErrorCode;
    phase: ExportPhase;
    message: string;
    remediation: string;
  };
  environment?: {
    rendererFingerprint?: string;
    verification: "nodeVerified" | "browserObserved" | "callerAttested";
  };
  partial?: {
    pages?: CompleteRenderReport["pages"];
    bindings?: Partial<CompleteRenderReport["bindings"]>;
  };
  unavailable: Array<{
    field: "environment.rendererFingerprint" | "bindings.pageMapDigest"
      | "bindings.htmlDigest" | "bindings.pdfDigest";
    reason: string;
  }>;
}

type RenderReport = CompleteRenderReport | FailedRenderReport;
```

The checked-in JSON Schema and canonical writer land with #438. The report is a separate sidecar;
HTML and PDF do not embed it, avoiding a digest cycle while the report binds their bytes. Failed
attempts retain a report when execution reached reporting safely, including CLI runs with an
explicit `--report` destination, but never a `complete` artifact result. The discriminated failed
shape records why a fingerprint or artifact binding is unavailable when failure precedes layout;
it never fabricates a PageMap digest or renderer identity merely to satisfy the schema.

## Renderer fingerprint

The renderer fingerprint is a canonical SHA-256 identity over every layout-relevant input:

- Docxodus package/core/WASM and paginator contract versions;
- render-report and PageMap schema versions;
- Playwright-core version plus Chromium product/build, launch/headless flags, operating system,
  architecture, locale, timezone, viewport, device scale, and media settings;
- a canonical layout-options digest including review/comment profiles and every pagination option;
- the font-configuration digest; and
- sorted requested-to-resolved font family, file identity, and version records.

The source package digest and document version are schema-bound beside the fingerprint. For a
Docxodus-launched browser with verified files and fonts, Node collects the authoritative facts and
passes the completed fingerprint into browser layout as the PageMap token. Browser-only and
unattested injected-browser callers receive a fingerprint over browser-observable facts plus the
explicit `browserObserved` verification level. Attested injected environments are
`callerAttested`, never `nodeVerified`; the report distinguishes observed fields from attested
fields. A caller reproduces a render with the same source bytes, layout options, exact runtime from
the fingerprint, and reported font configuration. Any change in those inputs is visible even when
output bytes happen to match.

## Supported fidelity

"Supported" means the automated gates below pass and no required resource is unavailable. It does
not mean arbitrary Word content is silently approximated.

| Capability | Release requirement |
|---|---|
| Page geometry | logical page count equals PDF count; each PDF MediaBox/CropBox is within 0.5 pt of its PageMap page |
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

Issue #489's current whole-paragraph footnote splitter can clip a note tail and omit its PageMap
geometry. Full footnote fidelity is explicitly unsupported until #489 lands; the exporter detects
the clipped/oversized condition and fails rather than shipping a complete result with missing text.

Package preflight, raw source identity, revision/comment/media inventory, safety findings, and ZIP
resource ceilings reuse #493's `PackageManifestGenerator` contract through its WASM surface. The
export path does not add a second partial ZIP inspector. The corrected #493 implementation is an
ancestor of the export stack and is exercised by the browser and Node/CLI production gates.

## Follow-on reconciliation

| Issue | Contract contribution |
|---|---|
| #438 | isolated final page-tree materialization, offline serializer, browser types/docs/tests |
| #439 | `@docxodus/export`, CLI, local runtime serving, batch/PDF bytes and typed failures |
| #440 | mixed-section PDF MediaBox/CropBox and sequencing proof |
| #441 | phased readiness barrier, quiet interval, delayed-resource tests |
| #442 | font directories, resolution report, strictness, fingerprint integration |
| #443 | generated-PDF raster/text/link ratchet and reproducibility documentation |
| #444 | shared final/original/markup and hidden/inline/endnotes/margin profiles |
| #489 | lossless mid-paragraph footnote continuation before export can claim complete note text/PageMap coverage |
| #465 / PR #499 | consume #439's exact batch session through the existing delivery seam |

These remain separate PRs in dependency order. #489 may land independently on `main`; #438 builds
on this decision; #439 builds on #438; #440–#442 and #444 stack on the first vertical PDF path;
#443 is the release gate. #439 supplies the production adapter contract, and a final commit on the
existing draft #499 consumes it without creating a second #465 PR. The production export stack is
rooted in the corrected #493 preflight contract.
