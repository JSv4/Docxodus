# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Changed
- **`ReplaceTextRange` with a needle that matches nothing now fails with the new
  `TextNotFound` error code (#490)** — `text_not_found` on the wire — carrying the anchor id
  and the needle, instead of returning an empty list a caller could not distinguish from a
  successful replacement. The same structured failure surfaces identically when the op runs
  as a mutation batch step: the step fails (rolling back an atomic batch) with
  `text_not_found` rather than the former `internal_error` or, post-#497, a silent no-op.
  **A caller that relied on the empty-list no-op** should pass `ExpectedMatchCount = 0`
  (`expectedMatchCount` on the wire) to assert absence as a successful no-op;
  `MaxReplacements = 0` likewise remains a successful found-but-unconsumed no-op.
- **Comment export profiles report what they cannot draw (#444).** Under any visible comment
  profile, comment threading raises `comment_thread_flattened` and resolved state raises
  `comment_resolved_state_not_rendered`, because the topology in `commentsExtended` is not read: a
  reply is drawn as an independent comment and a resolved comment is indistinguishable from an open
  one. (#444 also introduced revision warnings — `revision_family_not_rendered` and
  `revision_property_change_not_rendered` — but #538 and #539 below draw those families in this
  same cycle, so neither warning ships.) Both route through the `unsupportedContent` policy, so
  **an existing caller passing `unsupportedContent: "strict"` now gets a closed
  `resource_policy_failure` where such a document previously exported**; pass `"warn"` to keep the
  old outcome. The package manifest gains `revisions.runPropertyChanges`, splitting run-level from
  block-level property revisions in the inventory without a second pass. The profile contract is now also pinned where it
  was previously asserted without evidence: PDF text extraction per profile (deleted content
  extracts in document order under `markup` and never under `final`; inserted content never under
  `original`), revision rendering inside headers, footers, footnotes and endnotes, a comment range
  overlapping an insertion and a deletion, and per-profile comment-body rendering (`endnotes`
  lists every referenced body; `inline` drops a range-less reply body — the flattening the warning
  discloses; `hidden` renders none). The design doc no longer promises a threaded reply tree or
  resolved-state rendering the converter does not perform.

- **Render report schema v1 → v2 (#442).** `fonts[]` entries were a closed five-field record
  with a `browser | embedded | configured` source; they now carry the full resolver-backed
  resolution shape, `status` gains `load_failed`, and `source` becomes
  `browser | configured | attested`. `embedded` is gone: OOXML embedded fonts are not
  exported until de-obfuscation and embedding-license policy exist. `fontIdentity` changes
  from `{schemaVersion, digest, verification}` to the resolver and substitution contract
  identities it now binds. A consumer pinned to
  `docxodus/render-report.schema.json` receives the v2 document; the v1 schema stays in
  `docs/schemas/` for reading archived reports. Because the report is bound into the renderer
  fingerprint, fingerprints from before this change do not compare equal to ones after it.

### Added
<<<<<<< HEAD
- **The converter draws comment topology (#540).** `WmlToHtmlConverter` now reads
  `commentsExtended.xml` — where Word keeps the *shape* of a comment set — and carries both
  facts it records into every visible comment mode. A reply (`w15:paraIdParent`) nests
  beneath its thread root: an inner `ol.comment-replies` in the endnote-style section, a
  nested note that travels with the root as one page-positioned unit in margin mode (the
  pagination engine maps every id in a thread to the root note and dedupes per page), and —
  since a Word-authored reply has no range of its own — an inline marker that carries the
  reply body, its parent id, and a "Reply by …" title instead of dropping all three. A
  resolved comment (`w15:done`) is muted (`comment-resolved`) with a "Resolved" header badge,
  and its range highlight is muted in every mode. `CommentInfo` gains `Resolved` and
  `ParentId`. A malformed parent graph degrades to the flat rendering; without a
  commentsExtended part nothing changes. The standalone export's `comment_thread_flattened`
  and `comment_resolved_state_not_rendered` warnings are retired with the gaps they
  disclosed.
=======
- **The converter draws block-level property revisions (#539).** A tracked change to
  paragraph, numbering, table, row, cell or section properties (`w:pPrChange`,
  `w:numberingChange`, `w:tblPrChange`, `w:tblGridChange`, `w:trPrChange`, `w:tblPrExChange`,
  `w:tcPrChange`, `w:sectPrChange`) previously passed through `WmlToHtmlConverter` untouched,
  leaving no mark a reader could see under tracked-changes rendering. Each family now marks
  its block — a titled goldenrod change bar on the paragraph
  (`rev-para-format-change` / `rev-section-format-change`), a dotted outline on the table, row
  or cell (`rev-table-format-change`, `rev-row-format-change`, `rev-cell-format-change`), and
  a marker div where a trailing body `w:sectPr` records a section-property change — with the
  author and date under `IncludeRevisionMetadata`. Outside tracked-changes rendering nothing
  is emitted. The standalone export's `revision_property_change_not_rendered` warning is
  retired with the gap it disclosed; `revisions.runPropertyChanges` in the package manifest
  remains as inventory.
>>>>>>> origin/main
- **The converter draws custom XML revision ranges (#538).** `w:customXmlInsRangeStart`/
  `End`, `Del`, `MoveFrom`, and `MoveTo` — a reviewer's tracked add/remove/move of a custom
  XML structural wrapper — previously passed through `WmlToHtmlConverter` untouched, leaving
  no mark a reader could see under the `markup` profile. Each range boundary now renders as a
  bracket marker span (`rev-cxml-…-start`/`-end`) drawn as CSS-content brackets in the owning
  family's revision color, titled with what happened and, under `IncludeRevisionMetadata`,
  the author and date. Outside tracked-changes rendering the boundaries stay invisible and
  the enclosed content renders unchanged. The standalone export's
  `revision_family_not_rendered` warning is retired with the gap it disclosed.
- **Field-based internal cross-references from the session surface (#545).**
  `DocxSession.InsertCrossReference(anchorId, characterOffset, bookmarkName, options?)` inserts
  a Word-faithful `REF` field targeting an existing bookmark — a real field Word re-resolves on
  refresh, unlike an internal hyperlink. `CrossReferenceOptions` maps to the field's switches
  (`ReferenceNumber` → `\r`, `Hyperlink` → `\h`, `IncludePosition` → `\p`), and the written
  `w:fldSimple` carries a cached result run — the bookmarked text, the target's auto-number
  under `\r` (`0` when unnumbered, as Word shows), and/or the `above`/`below` position word —
  so renderers that do not recompute fields display a faithful snapshot. A missing or
  incoherent bookmark fails with `MissingBookmarkTarget`. Rippled to every transport: the WASM
  bridge and npm gain `insertCrossReference` (typed `CrossReferenceOptions`), the stdio host
  and `docx-scalpel` gain `insert_cross_reference`, and the MCP server's `docxodus_links` tool
  gains the batchable `insert_cross_reference` action.
- **Delivery-receipt verification on every transport (#520).** The portable JSON change
  receipt (#458) can now be verified from all four client surfaces, each routing through the
  new single-owner facade `DeliveryOps.VerifyChangeReceiptJson` (artifacts as
  `{"artifactId": "<base64>"}`, snake_case enums, malformed input answered with a structured
  invalid verdict): the WASM bridge gains `DocumentConverter.VerifyDeliveryReceipt`, npm gains
  `verifyDeliveryReceipt(receiptJson, artifacts?)` with typed
  `DeliveryReceiptVerificationResult`/`DeliveryArtifactVerification` results, the stdio host
  and `docx-scalpel` gain `verify_delivery_receipt`, and the MCP server gains the sessionless
  `docxodus_verify_receipt` tool (receipt and artifacts by path in the document scope, or the
  receipt inline). A vendored cross-language fixture (`TestFiles/Delivery/DR001-*`) is
  verified by the C#, Python, and browser suites alike so a canonical-format drift is caught
  on every side of the wire. Receipt building stays on the typed .NET surface driven by the
  #465 delivery operation — remote consumers verify; they do not compose.
- **Verified delivery bundles (#465).** `DeliveryBundleService.BuildAsync` derives a requested
  set of document, render, and evidence artifacts from exact baseline and working snapshots
  under an explicit two-part revision policy (`preserve`/`accept`/`reject` for pre-existing
  versus generated revisions), proves the review document's accept-to-final and
  reject-to-baseline directions, and returns immutable bytes plus a canonical
  `delivery-bundle-manifest/v1` — or publishes the same result as a fresh, atomically committed
  directory via `DeliveryBundleDirectoryPublisher`. The schema-v1 artifact vocabulary covers
  baseline/policy-baseline/working/review/final DOCX, standalone HTML, final and review PDF,
  PageMap, render report, package manifests, semantic and package deltas, comprehensive
  deliverable validation, the reversibility proof, and the deterministic change receipt
  (receipt issuance requires authoritative transaction evidence and stages the final
  deliverable's canonical verification result as the receipt's validation evidence).
  `DeliveryBundleVerifier.VerifyJson` re-verifies a manifest against independently supplied
  artifact bytes. Production HTML/PDF rendering crosses one injected
  `IDeliveryArtifactRenderer` boundary — a pure per-profile-pair `DescribeBatch` whose
  layout-options and runtime-policy digests enter the group key, and one `RenderBatchesAsync`
  call per build carrying every group in stable order; `DocxodusExportHostRenderer` sends that
  call as a single `docxodus-export-host` framed request with digest-deduplicated sources,
  process-owned executable authority, and no PATH discovery, and represents a document version
  outside JavaScript's safe range as the closed `document_version_unrepresentable`
  unavailability before any frame is built. Ships as the `docxodus-deliver` CLI
  (`tools/delivery`) and the MCP `docxodus_deliver` tool over the same service. See
  `docs/architecture/delivery_bundle.md`.
- **`docxodus_compare` MCP tool (#466).** A sessionless agent-server tool that turns stored
  document versions into one native tracked-changes redline: `baselinePath` plus `revisedPath`
  (two-way `DocxDiff` compare) or `revisedPaths` with per-reviewer `authors` (N-way
  consolidate), written to `outputPath`. Every path resolves through the document store's
  containment check like `docxodus_open`'s, the response summarizes generated revisions by
  author, and the tool is refused as a `docxodus_mutations` step — it mutates no session.
  Backed entirely by the existing `DocxDiffOps` facade; no new engine surface.

- **All nine #466 scenario families ship in the eval corpus.** Clause insertion into a real
  auto-numbered list (renumbering proven through rendered labels), comment + threaded reply +
  footnote + bookmark-anchored internal link (a field-based `REF` cross-reference is not yet
  authorable — #545 tracks the op), body edit under part-identity protection of the signature
  block, header, and page-number-field footer, content-control template fill (via the corpus's
  first programmatic fixture builder, since the tool surface fills but cannot create controls),
  two-reviewer compare-and-consolidate through `docxodus_compare` with a `redline` invariant
  group asserting the written redline's attributed revisions through a fresh session, and an
  edit over pre-existing revisions and comments that must both survive. The
  master-services-agreement fixture now carries a footer page-number field and a pre-existing
  comment; every scenario runs in the fast deterministic tier.

- **Workflow evaluation runner completed (#466).** The eval runner now supports multi-anchor
  steps (`targets` maps an argument name to a content-resolved target, unlocking range formats,
  moves, and bookmark spans), a `changedPartsMustBeWithin` part-URI allowlist that turns "how
  many parts changed" into "the right parts changed", and `trackedRevisions`/`comments`
  invariant groups read through the same MCP dispatcher the steps drive — so a scenario can
  assert a pre-existing reviewer's markup is still live outside the reversibility path. The
  scenario schema is now enforced rather than decorative: every scenario file is validated
  against it, its invariant vocabulary is pinned to the checkers' vocabulary in both
  directions, and each new checker ships a negative control proving it fails when violated.
  Every run writes a per-scenario `scorecard.json` — the machine-readable engine baseline
  later agent-scored runs are compared against — CI uploads eval artifacts on failure, and a
  weekly `eval-corpus` workflow executes the opt-in tier under `eval/scenarios/corpus/`
  (`DOCXODUS_RUN_EVAL_CORPUS=1`). Fixtures declare their own `expectedContent` anchor, so the
  corpus is no longer limited to a single fixture. Delivery change receipts stay deliberately
  out of the eval artifacts: their lineage must come from mutation evidence captured at
  execution time, which belongs to the #465 delivery operation.
- **Multi-author Consolidate coverage and schema conformance for the reversibility proof (#464).**
  The proof suite now exercises `DocxDiff.Consolidate` output: a redline carrying two distinct
  reviewer authors — disjoint edits and a policy-resolved merge conflict — proves reject ≡ shared
  base and accept ≡ policy-resolved composite at the story-text and modeled-semantic levels, with
  every revision classified as generated for its own reviewer (multi-author attribution is not an
  ownership conflict). A new contract test validates the emitted canonical proof JSON — from the
  verifier and from the shared `VerificationOps` facade, in both completed and fail-closed shapes —
  against the checked-in `redline-reversibility-proof-v1` schema, with negative controls proving
  the validation rejects each constraint kind the schema relies on.

- **Workflow evaluation scaffold (#466).** Added `eval/`, a corpus of deterministic
  document-workflow scenarios, and `Docxodus.Tests/Eval/`, the scripted caller that runs them.
  A scenario declares a fixture, the tool calls that perform the task, and the invariants that
  decide the run — scored on task completion, target precision (distinct anchors in the #457
  change set), collateral package change (#456 manifests), deliverable validity (#463), redline
  reversibility (#464), and HTML rendering. The caller drives the MCP tool surface rather than
  the .NET API, so the engine baseline is measured where an agent meets it and a failure can
  never be attributed to model planning. Fixtures are build scripts replayed over a blank
  document, so the corpus carries no third-party bytes. Ships the term-replacement,
  notice-period-amendment, and table-economics scenarios as the deterministic fast subset; PDF
  and visual-regression scoring stay with #443's ratchet.

- **Redline reversibility proof on every client surface (#464).** The proof engine landed with
  #497 but was reachable only from .NET. `VerificationOps.ProveRedlineReversibility` now owns the
  wire shape, and the canonical `redline-reversibility-proof/v1` document is available from the
  WASM bridge, the npm package (`proveRedlineReversibility`, plus an off-main-thread worker path
  because three packages are inspected and two rebuilt), the stdio host
  (`prove_redline_reversibility`), `docx-scalpel` (`prove_redline_reversibility`, decoded into
  typed frozen dataclasses), and the MCP server (`docxodus_track_changes` action
  `prove_reversibility`, which proves the session's clean-save checkpoint against two documents
  resolved through the document store). The rebuilt packages stay in-process — the proof already
  carries their digests and the divergences between them and each expected document.

- **Long footnote paragraphs keep a complete PageMap when they continue (#489).** A note paragraph
  taller than the maximum note band, and a long leader followed by short tails, are now covered by
  `standalone-export.spec.ts`: both must continue across pages with `running_story_placement`
  complete, no pending work, and PageMap fragments on every page the note reaches. The footnote
  fixture generator accepts per-paragraph word counts so uneven continuation pressure can be
  constructed directly.

- **Generated-PDF fidelity ratchet (#443).** `@docxodus/export` PDFs now run through the same
  Poppler raster contract the browser-page benchmark uses, over a ten-document pinned corpus with
  recorded provenance. Conversion, page count, physical geometry, semantic content and chart-vector
  contracts are unconditional gates that no raster severity or disposition can waive (the chart
  contract requires a chart case to emit vector path operations at all); SSIM and ink
  metrics ratchet separately against a numbers-only record, and text and link extraction are gated
  independently so a text regression cannot hide behind an acceptable raster score. Extraction from
  the *reference* PDF is reported rather than gated, so a LibreOffice or Poppler change cannot fail
  a contract that names Docxodus. An environment
  fingerprint covering LibreOffice, Chromium, Poppler and the font contract means a changed
  environment reports `environment-changed` rather than being misattributed to the renderer. The
  benchmark stays `workflow_dispatch`-only until #444 lands, per the release-gate ordering in the
  design doc.

- **Verified font runtime (#442).** `fontDirectories` is live: the Node adapter
  deterministically discovers TTF/OTF/WOFF/WOFF2 files across the ordered directories,
  rejects symlinks and escaping or changing paths, reads family and face metadata, hashes
  every file, and gates injection on OS/2 `fsType` — a restricted-license face fails the
  export closed rather than being silently dropped or substituted, and every WOFF/WOFF2 file
  requires an explicit caller attestation regardless of what its OS/2 bits would otherwise
  permit, since embedding rights cannot be derived from the compressed format alone. The
  resolver reaches the isolated page as a Playwright binding, never as serialized bytes, so
  the materializer only ever sees the versioned `FontResolver` contract and the page never
  gains filesystem access; a Node-backed resolver's own failure (a rejected symlink, a
  resource limit) crosses back with its original code and remediation intact rather than
  collapsing into a generic message. Resolution records now carry the request identity,
  family stack and kinds, style, weight, stretch, sample digest, resolved face, file digest
  and version, face match, metric compatibility, glyph coverage, license evidence, and a
  single `verified` flag standing in for the whole strictFonts question, and the renderer
  fingerprint binds the resulting configuration identity. `strictFonts` finally enforces: it
  rejects any outcome that is not an exact, digest-identified, license-evidenced face with
  complete coverage, replacing the `unsupported_runtime` stub that previously rejected every
  configured font environment outright. The visual-parity font contract now derives from the
  production substitution contract instead of duplicating it, keeping only deployment-specific
  package hints on the test side.

- **Deterministic print-readiness barrier (#441).** The export barrier now proves what it used
  to assume. Requested font families are probed for actual availability instead of being taken
  on trust once `document.fonts.ready` resolves — `FontFaceSet.check()` answers true for families
  Chromium has never heard of, so resolution is measured through advance widths against every
  generic fallback. A family the environment silently substituted for is recorded as `missing`
  rather than `unverified` and raises one aggregate `font_family_unavailable` warning; an
  available family stays `unverified`, because the barrier proves the browser can render it and
  not which file it came from. Undecodable images and unmeasurable inline SVG now route through
  the `unsupportedContent` policy as `image_decode_failed` / `chart_svg_unmeasurable` warnings
  with an omitted resource record, instead of collapsing into an untyped `conversion_failure`;
  under `strict` they fail closed at their own phase. The offline reopen check keeps assertion
  semantics, so a resource that materialized and then did not survive serialization is still an
  `output_verification_failure`. On the Node side the render report's readiness log now covers
  the host-owned phases the browser materializer cannot see from inside the page —
  `browser_launch`, `wasm_initialization`, `output_verification` and `cleanup` — prepended in
  the order the work actually happened, so a timeout in any of them names its phase and pending
  resources instead of surfacing a bare error code.
- **Mixed-section physical PDF geometry (#440).** The shared paginator now transfers continuous
  spill pages to the section that supplies their body while preserving predecessor-owned stories
  on the shared page, carries section-specific header/footer distances and logical page numbering,
  selects odd/even stories from each section's one-based page position, inserts story-free blank
  pages for odd/even section starts, and honors `footnoteLayoutLikeWW8` when a pre-break footnote
  meets a continuous section. PDF verification resolves inherited MediaBox/CropBox, rotation, and
  `UserUnit` before comparing physical origins and dimensions. A generated six-page Letter/A4
  portrait/landscape fixture proves explicit, next-page, continuous, two-column, header/footer,
  footnote, and page-field sequencing through the production Node renderer, including a scaled
  screen-view reprint with unchanged physical boxes and text placement.
- **Deterministic delivery change receipts (#458).** Added a versioned, canonical JSON
  receipt that binds source/delivered package identities to every supplied mutation transaction,
  normalized requests, explicit undo/redo lineage, requested/derived/unexpected package-change
  attribution, required typed semantic evidence, optional validation/reversibility evidence, and
  independently hashed clean DOCX, review DOCX, HTML, PDF, image, and report artifacts. Exactly
  one clean DOCX must match the delivered raw package bytes, while exact #457 canonical semantic
  bytes cover the source-to-delivered comparison and every state-changing transaction. Privacy
  profiles support hash-only, structural-summary, and full-evidence output. Page citations are
  accepted only when their exact PageMap bytes project the claimed coordinates for a reachable
  document version/package digest and match the render fingerprint and artifact. Undo/redo is
  checked as a LIFO state machine, and any indexed package attribution may point only to a
  successful retained operation. Failed atomic receipts retain the complete requested-operation
  list with explicit not-executed/rolled-back status, while a genuinely successful no-op may retain
  an empty edit-result list. Strict typed semantic reconstruction, authoritative clean-DOCX inventory
  recomputation, portable path normalization, canonical collection-order checks, aggregate object
  budgets, stream-charged JSON collections, and capped canonical writers fail closed under forged or
  oversized input. The portable verifier detects receipt,
  record, artifact, semantic, lineage, and citation-binding tampering. See
  [`docs/architecture/delivery_change_receipt.md`](docs/architecture/delivery_change_receipt.md).
- **Node and CLI standalone PDF export (#439).** Added the separately published
  `@docxodus/export` companion with immutable byte and stable-file APIs, one-pass HTML/PDF batches,
  a strict `docxodus convert` CLI, and a length-framed integration host. It drives the shared
  hashed browser materializer through an isolated pinned-Chromium context, denies requests outside
  the closed runtime graph, prints the finalized offline page tree with exact CSS page sizes and
  zero margins, and parser-verifies page geometry, tags, searchable text, links, vector charts,
  digests, and volatile PDF metadata. Typed failures retain structured reports, file publication is
  fsynced/atomic/no-replace, caller-owned browsers remain open, and CI publishes a viewable success
  and failure artifact gallery.
- **Standalone paginated HTML materialization (#438).** Added the UI-free
  `docxodus/export-browser` entry point for immutable DOCX-to-offline-HTML export, with a finalized
  page-box-only tree, post-cleanup PageMap, digest-bound render report, structured warnings and
  failures, resource limits, font-environment disclosure, pristine-tree stability retry, offline
  reopen verification, and a closed hashed runtime-asset graph. The paginator now operates in the
  owning DOM realm, and converter relationships resolve from their actual story parts so header and
  footer resources survive. Includes a file-picker/download example and Playwright artifacts for
  the HTML, PageMap, report, and offline screenshot.
- **Deterministic, non-mutating DOCX package manifests (#456).** Added a versioned
  verification artifact that inventories every OPC entry and content type, preserves duplicate
  occurrences, resolves every package/part relationship, reports dangling references and
  malformed/encrypted packages, extracts renderer-relevant facts, and separates exact package,
  ordered OPC-content, and normalized-semantic SHA-256 identities. XML normalization ignores only
  documented serialization choices through an explicit known-content-type allowlist, preserves
  unknown vendor-extension/custom XML whitespace and `xml:space`, and
  keeps Strict and Transitional namespaces distinct. Declared and actual payload reads now also
  have a configurable absolute per-entry byte ceiling alongside the aggregate and compression-ratio
  limits. Available from .NET, live sessions,
  WASM/npm (including workers), the stdio Python host/client, and MCP
  `docxodus_get_content(format: "manifest")`.

  `isValid` is a claim about the package, so it is reserved for defects a real file does not
  have: package-absolute relationship targets (`Target="/word/document.xml"`, the form the Open
  XML SDK writes) resolve normally, and empty directory-only ZIP entries — which 7-Zip, Windows'
  *Send to → Compressed folder*, and several Word templates emit — are a `directory_entry`
  warning, inventoried with a trailing-slash URI and excluded from both content digests; a
  trailing-slash entry with payload is invalid and remains in both identities. ZIP64 entry sizes
  are decimal strings on the JSON wire, avoiding silent precision loss in JavaScript. ASCII ZIP
  item names are mapped losslessly to Unicode logical OPC names without collapsing reserved or
  opaque percent escapes. A
  finding whose real cause is one unreadable file is reported once against that file rather than
  once per part: an unusable `[Content_Types].xml` yields `content_types_unreadable` instead of
  `missing_content_type` on every entry, and an unparsed `.rels` part yields
  `relationship_part_unreadable` instead of `dangling_relationship` on every reference it owns.
  Breaching the entry-count limit now suppresses both content digests, because an inspection that
  stopped early cannot distinguish two packages that differ only past the cut, and declared
  expansion is summed over the whole central directory so the two limits cannot be played against
  each other. A part that declares an XML content type but whose bytes are not XML now
  contributes those bytes to `normalizedSemanticDigest` rather than voiding the package's
  identity; a part merely *skipped* by `MaxXmlPartBytes` still leaves the digest `null`, because
  a larger budget would have normalized it and the identity must not depend on the caller's
  options. See
  [`docs/architecture/package_manifests.md`](docs/architecture/package_manifests.md), including
  its **Known limits of schema v1** section.
- **Stable, comprehensive semantic DOCX diff (#457).** `SemanticDiff.Compare` and
  `DocxDiff.GetSemanticChanges` return the versioned, deterministic
  `docxodus.semantic-changes` schema with owning part/path, side-specific anchors
  and scopes, `insert`/`delete`/`move`/`modify`, and closed typed before/after
  values. Coverage spans text/structure/formatting/styles/numbering, tables,
  sections and page setup, all story/note/comment families, links/bookmarks/SDTs,
  images/media/relationships, revisions/annotations, and opaque package parts.
  Serialization-only XML, coordinated relationship-id rewrites, and
  relative-versus-absolute internal relationship targets are suppressed, while
  target swaps at relationship-reference locations remain visible. Unknown XML
  remains visible with whitespace-, comment-, and processing-instruction-preserving
  normalized digests; full-part style/numbering/theme residuals cover values outside
  the typed IR registries. Package inspection reuses #456's manifest generator, XML normalizer,
  relationship resolution, safety limits, and finding locations; it runs before SDK parsing and has
  configurable entry-count, part-URI, per-entry, aggregate-decompressed-byte, and
  compression-ratio bounds, with duplicate part names and relationship ids rejected. Invalid
  manifests fail closed even when the package-change supplement is disabled.
  Typed table and cell values are read from that table's own `w:tblPr`/`w:tblGrid` and that cell's
  own `w:tcPr`, so a nested table's style, width, and column widths are reported against the nested
  table's anchor instead of the containing one. Integers that a package can carry outside the v1
  safe range — `wp:extent/@cx`, `w:gridCol/@w`, `w:bookmarkStart/@w:colFirst`, and declared entry
  sizes — are emitted losslessly as decimal strings rather than failing the comparison.
  The formal v1 JSON Schema is published at
  [`docs/schemas/semantic-changes-v1.schema.json`](docs/schemas/semantic-changes-v1.schema.json).
  The existing redline, revision,
  edit-script JSON, and `DocxSession.GetDiff` APIs are unchanged. Available through
  .NET, WASM/npm, the Python host/client, and MCP
  (`docxodus_get_content format:"semantic_changes"`). Session comparison keeps the
  opening package when `CaptureInitialProjection` is enabled (the MCP open tool can
  decline with `captureInitialProjection:false`) and compares its checkpoint
  serialization against an isolated checkpoint of the current state, so an
  unedited session reports zero changes for any openable package. The surface
  shares the sibling entry points' read pipeline — strict→transitional and
  `mc:AlternateContent` normalization plus the `OnCompatibilityWarning`/
  `ThrowOnCompatibilityWarning` gate — and the package safety limits are
  `PackageManifestOptions` defaults, declared once. Serialization bookkeeping
  (rsids, `w14:paraId`/`textId`, annotation-part indentation) is never a change;
  package records use the IR's `hdr{N}`/`ftr{N}` scope vocabulary; and a
  bookmark/revision/binding whose containing block the IR aligned in place is
  never a `move`. The header/footer path grammar pins the `w:type` kind vocabulary
  (`default`/`first`/`even`). Design and measured 1,000-paragraph guard:
  [`docs/architecture/semantic_diff.md`](docs/architecture/semantic_diff.md).
- **Idempotent mutation transaction identities for the MCP server** (issue #449).
  An applying `docxodus_mutations` batch may carry a caller-chosen root
  `transactionId` (non-blank, at most 256 Unicode scalar values). The first
  terminal response — success, partial, structured failure, precondition failure,
  or safely-caught exception — is retained for the lifetime of that open session,
  and an identical retry returns it byte-for-byte without executing anything or
  re-evaluating preconditions, so generated anchors, timestamps, versions and
  `packageHash` all survive a lost response. Results gain a top-level
  `transaction: { schemaVersion, transactionId, requestFingerprint }`; the
  fingerprint is a SHA-256 over a canonical rendering that excludes only the root
  `sessionId`/`transactionId`. Reusing an id for a different request returns
  `transaction_conflict`. Preview/dry-run batches, nested step args, and the other
  tools reject transaction ids rather than ignoring them. Retention is bounded per
  session by both a count and a byte budget (128 responses, 32 MiB) followed by
  1,024 response-less tombstones; there is no TTL and the number of open sessions
  is not bounded, and once a tombstone expires a late retry applies again — both
  documented as hazards in
  [`docs/architecture/docx_agent_server.md`](docs/architecture/docx_agent_server.md).
  Idempotency is MCP-only: `execute_batch` through WASM/npm and the stdio host has
  no equivalent. Adds `EditErrorCode.InvalidTransaction`, `TransactionConflict`,
  `TransactionResultEvicted` and `TransactionIncomplete`, rippled to npm
  `EditErrorCode` and Python `EditErrorCode`. MCP session dispatch is now
  serialized per session so a retry cannot race another action on the same
  document. Coverage: `McpMutationTransactionTests` MCP449,
  `python/tests/test_transaction_error_codes.py`, and
  `npm/tests/transaction-error-codes.spec.ts`.
- **Structural tracked revisions and one live revision registry (#455).** *(Breaking — see
  the Breaking changes section below.)* `DocxSession` now enumerates and resolves the
  structural revision families it previously ignored —
  table-cell insert/delete/merge (`w:cellIns`/`w:cellDel`/`w:cellMerge`), content-control
  (`w:sdt`) insert/delete envelopes, and numbering revisions (`w:numPr/w:ins`,
  `w:numberingChange`) — through the same part-aware registry that already owned
  content, paragraph-mark, row, move, and `*PrChange` revisions. `RevisionListEntry`
  gains `Family`, `ConstituentIds`, `PartUri`, `Scope`, `AffectedAnchors`,
  `ResolutionStatus`, and `Diagnostic`; `AcceptAllRevisions`/`RejectAllRevisions` become
  ordinary undoable session mutations built on the selective resolver instead of a
  whole-document `RevisionProcessor` byte transform plus a session rebind. Structural
  editing in `TrackedChangeMode.RenderInline` emits native markup for table row
  insert/delete and column insert, and for single-paragraph list application, removal,
  and level changes; shapes with no safely reversible encoding return
  `TrackedOperationUnsupported` without mutating or recording history, and a table with
  an unresolved cell revision returns `UnresolvedStructuralRevision`. Rippled through
  `DocxSessionOps`/JSON, WASM/npm, the stdio host + `docx-scalpel`, and MCP
  `docxodus_track_changes` (now also usable as a `docxodus_mutations` batch step). See
  `docs/architecture/docx_mutation_api.md` and `docx_agent_server.md`.
- **Native content-control operations (#452).** `DocxSession` now enumerates and fills
  Word structured-document tags as first-class objects. `ListContentControls` /
  `GetContentControl` return every `w:sdt` in outer-before-inner story order under a
  stable `sdt:{scope}:{unid}` anchor derived from the native `w:sdtPr/w:id`, with family,
  placement, owning part, parent/depth, native metadata, data binding, current text, list
  item values, and an explicit `CanMutate`/`UnsupportedReason` decision.
  `FillContentControlText`, `FillContentControlRichText`, `SetContentControlChecked`,
  `SetContentControlDate`, `SelectContentControlItem`, `FillContentControlPicture`,
  `AddRepeatingSectionItem`, and `RemoveRepeatingSectionItem` mutate through the wrapper
  without rebuilding it, preserving `w:sdtPr` metadata and the placeholder definition.
  Data-bound controls fail closed unless `bindingPolicy: detach_target` removes the
  target's own binding; a bound or locked ancestor always fails closed, and no Custom XML
  part is ever edited. `sdt` becomes an AnchorIndex kind in both the WML projector and the
  IR emitter, and `ListInlineSpans` reports outer-to-inner `ContentControlAnchorIds`.
  Rippled through the JSON facade, WASM/npm, the stdio host and `docx-scalpel`, and the
  new `docxodus_content_controls` MCP tool. Design: `docs/architecture/native_content_controls.md`.
- **Canonical table addressing and complete table-operation ripple (#450, absorbing
  #471).** Tables now expose explicit stable identities for the `w:tbl`, every
  `w:tr`, every physical `w:tc`, and every `w:tblGrid/w:gridCol`, plus
  `GetTableMetadata`, `ResolveTableCellAnchor`, and
  `ResolveTableCellCoordinate` for bidirectional anchor ↔ zero-based Word-grid
  coordinates. All cell operations now consume one unambiguous canonical `tc`
  anchor; a legacy paragraph inside a cell is translated to its nearest cell for
  the compatibility window, while `tbl`/`tr`/unrelated anchors fail with
  `TableAnchorMigrationRequired` and remediation. This also fixes nested-table
  operations accidentally retargeting an outer cell. Shape-changing edits return
  `EditResult.TableAnchors` with deterministic retained (before/after location),
  added, and invalidated table/row/column/cell identities. Missing or
  underspecified `tblGrid` columns are read-only virtual metadata until a column or
  width transaction materializes real `gridCol` anchors and reports the virtual
  identities invalidated. `DocumentStructure` keeps its path-based `Id` for
  compatibility and adds canonical `AnchorId`; its table geometry now honors
  `gridBefore`/`gridAfter`, horizontal spans, and actual vertical-merge row spans.
  Rippled through JSON/ops, WASM/npm (including `SetTableRowOptions`), the Python
  host/client (all table operations), and MCP. MCP table actions now use distinct
  schema fields: `anchorId` for insertion, `tableAnchorId` for table reads, and
  `cellAnchorId` for every cell operation. Existing merge OOXML semantics are
  preserved; emitting new tracked table revisions remains #455. Coverage:
  `DocxSessionTableAddressingTests` DT250–DT257, the existing table/MCP suites, and
  `python/tests/test_table_addressing.py`.
- **Portable, renderer-authored `PageMap` and exact page citations** (issue #454).
  Browser pagination can now materialize a versioned map of physical pages and every
  canonical `kind:scope:unid` source fragment, with page-relative point geometry,
  story/table ownership, page style identity, document version, and renderer
  fingerprint. `DocxSession` validates and registers external maps; search, structural
  find, and scoped projection APIs optionally attach citations across .NET, WASM/npm,
  stdio/Python, and MCP. Mutations stale maps automatically, fingerprint mismatches are
  rejected, and continuous/no-map layouts return typed unavailable results instead of
  guessed pages. `paginateHtml`, React `PaginatedDocument`, and
  `navigateToPageCitation` expose materialization and preview navigation. The MCP inline
  preview remains explicitly continuous pending #434. See
  [`docs/architecture/page_map.md`](docs/architecture/page_map.md).
- **Intrinsically isolated mutation preview** (issue #446). `.NET` `PreviewBatch`,
  Ops/JSON, WASM/npm, stdio/Python, and MCP now run the identical atomic or explicit
  `best_effort` batch path on a complete shadow package instead of applying to the live
  session and undoing. The clone carries every OPC part/relationship/media/custom-XML
  payload plus version, mutable configuration, diff baseline, and id generators, while
  caches and undo/redo history remain independent; failure, interruption, disposal, and
  abandonment therefore cannot touch live bytes or history. Rich apply/preview receipts
  include predicted versions, per-step created/removed/modified anchors and patches,
  revision/comment/annotation deltas, warnings, a canonical package-content SHA-256, and
  optional scoped/full shadow-only HTML. Deterministic previews and applies have exact
  receipt/hash equivalence. Operations that generate anchors/OOXML ids or timestamps are
  explicitly semantic-equivalence-only (same outcomes and structure/content/relationship
  effects modulo generated metadata) and emit warnings. This supersedes the undo-depth,
  redo-destruction, and crash window described in #468.

  The preview HTML profile has a single owner (`HtmlConversionOps.PreviewDocumentOptions`
  / `PreviewBlockOptions`), reached from the browser through the new
  `RenderPreviewHtml` / `RenderPreviewBlockHtml` bridge exports, so every surface's
  preview of the same batch describes the same document (tracked changes, comments,
  annotations, notes, and headers/footers shown) rather than the editor's authoring
  render. Receipt change-set membership is compared on each entry's serialized wire
  projection rather than CLR equality, matching what the browser client compares.
  `packageHash` is `null`, never `""`, when it could not be computed, so an absent hash
  cannot satisfy a replay-equality assertion. `MutationPreviewHtmlMode` is exposed to
  Python as an enum (`docx_scalpel.MutationPreviewHtmlMode`).

  **Cost note.** Receipt enrichment is unconditional on both the apply and the preview
  path: each batch inspects revisions, comments, and annotations twice (each forcing an
  anchor index) and computes a package-content hash, which serializes and hashes a full
  package checkpoint. A preview additionally clones the package and opens a second
  `WordprocessingDocument`, roughly doubling peak memory for its duration — material for a
  large document on a browser WASM heap. There is deliberately no opt-out in this release;
  whether to gate enrichment behind a setting remains an open public-API decision.
- **Atomic multi-step mutation batches** (issue #445). `DocxSession.ExecuteBatch`
  and the reusable nested-safe `BeginTransaction` primitive checkpoint the complete
  OPC package, relationship topology, anchor/revision generators, mutable session
  configuration, version, and both undo/redo cursors. Atomic mode is the default:
  all available preflights run before step zero; success advances the version once
  and creates one undo unit; any failed or thrown step restores the exact package
  and history state and returns its index/tool/action/error with `rolledBack: true`.
  Explicit `best_effort` retains sequential partial-success behavior. The contract
  is available through .NET/Ops/JSON, WASM/npm, stdio/Python, and MCP;
  MCP's legacy `apply` spelling is now a deprecated alias for `best_effort`.
  Structural table operations are batchable on every surface, and a batched step's
  receipt keeps its full `tableAnchors` mapping so a caller can address the cells the
  same batch just created. Guards a preflight can decide read-only are evaluated at
  the batch-start state and not re-evaluated per step; `expectedMatchCount` is the
  one exception and is enforced by the replacement itself, at that step's turn.
- **Optimistic mutation preconditions and a monotonic document version** (issue
  #447). Every `DocxSession` starts at version `0` and advances exactly once for
  each committed mutation, undo, or redo; failures and successful no-ops leave it
  unchanged. `MutationPreconditions` can guard the expected version, target
  anchor/hash/exact visible text or range/kind/scope, and find/replace occurrence
  count. A mismatch returns `PreconditionFailed` with structured expected/actual
  values plus the current version and target metadata, without changing bytes or
  undo history. The same camel-case shape is exposed by the WASM/npm, stdio/Python,
  and MCP transports; `AnchorInfo` now includes `contentHash` and `visibleText`.
- **Exact occurrence-count replacement.** `ReplaceOptions.ExpectedMatchCount`
  requires the live literal-match count before `ReplaceTextRange` proceeds. Guard
  evaluation, counting, and the whole multi-match rewrite share one mutation gate
  and one undo snapshot, so duplicate text cannot turn a stale plan into a partial
  replacement.
- **First-class native image inspection and editing across every session surface (#453).**
  `DocxSession` now enumerates image occurrences across body, headers, footers, footnotes,
  endnotes, and comments. It can insert, replace, resize, describe, reposition, or remove the
  canonical DrawingML subset. PNG/JPEG/GIF/BMP/TIFF bytes are validated by signature and dimensions;
  WebP, external links, legacy VML, multi-picture/non-canonical DrawingML, and unsupported
  floating layouts remain truthfully enumerable but read-only. Image relationships are owned by
  the actual story part, identical media is reused across owners, and undo/redo restores bytes,
  content type, exact media URI, owner-local relationship ids, and external targets. Runtime
  capabilities, points-versus-EMU units, 96-DPI default sizing, size caps, and base64-only JSON
  transports are exposed through .NET, JSON ops, WASM/npm, stdio/Python, and MCP
  (`docxodus_images`). Three behaviours worth knowing before you adopt it:
  - **Orphan cleanup deletes only provably unreferenced relationships, and only on the mutation
    path.** The sweep asks whether the relationship id appears in *any* attribute of the owning
    part, not whether it appears in a whitelist of `r:embed`/`r:link`/`r:id`. OOXML names image
    relationships through more attributes than the DrawingML pair (VML/OLE spellings such as
    `o:relid` and `r:href`), and a whitelist would silently and irrecoverably drop media
    referenced any other way. Because orphaning is something a *mutation* does, the sweep runs
    when an op's edit lands rather than when the document is serialized — covering the transforms
    that drop a `w:drawing` without any image API involved, including one whose edit lands in a
    story part other than the one it resolved.
  - **`Save` and `ConvertToHtml(session)` are read-only with respect to relationships.**
    `ConvertToHtml(session)` is implemented as `session.Save(persistAnchorIds: true)`, so
    save-time normalization would otherwise run on a caller who only asked to render. Neither
    now changes relationship topology or media: an orphan already present in the opened bytes
    survives any number of renders and saves and is cleaned up only by the next mutation, so
    opening and saving a document back unchanged no longer silently deletes media the session
    never touched.
  - **`ReplaceImage` is dimension-preserving.** It rewrites `r:embed` only; `wp:extent` and
    `a:xfrm/a:ext` keep their EMUs, so the new media renders into the old box. `ListImages()`
    reports the new intrinsic pixels immediately, and
    `SetImageDimensions(id, w, h, preserveAspect: false)` re-fits from them. `PreserveAspect`
    always scales the *current rendered box*, never the media's intrinsic ratio.

  Pictures whose `a:blip` carries an extension list (an SVG `asvg:svgBlip` and its raster
  fallback) are enumerable but `canMutate:false`: replacing only the fallback would leave an
  SVG-aware renderer showing the old image while the API reported success.
- **First-class hyperlinks and bookmarks across every editing surface (#448/#451/#469/#470).**
  `DocxSession` can enumerate and mutate external or bookmark-target hyperlinks and paired,
  multi-paragraph bookmarks with exact character spans. External relationships are owned and
  reused by the containing body/header/footer/footnote/endnote part; internal links use
  `w:anchor` without a package relationship. Rename retargets inbound links atomically, removal
  refuses live targets, malformed/cross-part ranges return structured errors, destructive edits
  cannot orphan markers, and undo/redo restores relationship topology. "Inbound reference" covers
  both consumer families: `w:hyperlink/@w:anchor` **and** `REF`/`PAGEREF`/`NOTEREF`/`HYPERLINK \l`
  cross-reference fields in `w:instrText` and `w:fldSimple/@w:instr` — so renaming a bookmark a
  table of contents points at retargets those fields instead of leaving "Error! Bookmark not
  defined." behind, and removing one is refused while a field still cites it. That guard also gates
  structural deletion, so on a document with a table of contents deleting a heading is refused
  (in default, untracked mode) until the citing field goes with it. Word's own
  `_GoBack`/`_Toc*`/`_Ref*`/`_Hlt*`/`_Hlk*` namespace is closed to *creation* (Word reallocates
  it); bookmarks Word already placed there stay fully readable and mutable. Adding a hyperlink
  over a span relocates the zero-width `w:bookmarkStart`/`End` and `w:commentRangeStart`/`End`
  markers inside it into the new `w:hyperlink` rather than stranding them after it, and a
  cross-part `MoveBookmark` takes a fresh document-global `w:id` instead of carrying a
  part-scoped one into a part that may already use it. Splitting a paragraph inside a hyperlink
  gives each half its own identity, so both remain individually addressable. The same contract is
  exposed through JSON ops, WASM/npm, stdio/Python, and MCP (`docxodus_links`); Markdown links now
  use the same owner-aware promotion and orphan cleanup. Listing and mutation reach the comments
  story part as well, so `ProjectionScopes.Comments` is a real scope rather than a silent empty
  result. Coverage includes Open XML validation, save/reopen identity, exact run-format
  boundaries, repeated story-scoped bookmark ids, tracked limitations, and relationship cleanup,
  plus peer suites in `python/tests/test_links_bookmarks.py` and
  `npm/tests/docx-session-links.spec.ts`. This supersedes the earlier tracked-move clone policy:
  a tracked block move containing bookmark markers now fails before snapshot instead of creating
  two simultaneously-live copies of a globally unique bookmark name — and `ValidMoveTargets`
  mirrors that rejection, so a drag UI never advertises a drop the engine will refuse.
- **Complete inspect-before-edit formatting surface (#448).** `DocxSession` now exposes an explicit
  style catalog (`ListStyles`), direct-versus-effective paragraph/run formatting
  (`GetFormatting`), and enumerable mutation-compatible run spans (`ListInlineSpans`). Effective
  properties reuse `FormattingAssembler`'s document-default and style-chain rollups. List
  membership now includes its query `AnchorId`, abstract-level `Start`/`LevelText`, and effective
  indentation; section info includes its mutation-ready body `AnchorId`, and section identifiers
  are stored Unids rather than positional fallbacks. The same JSON schema is rippled through the
  WASM/npm, stdio/Python, and MCP surfaces (`get_content` formats `styles`, `formatting`, `spans`;
  `info` is per-anchor). Returned style ids, anchors, and spans are tested by feeding them unchanged
  into their matching mutation APIs. Table geometry and inline memberships remain separate work.

  **Breaking (source):** `ListMembership` and `SectionInfo` gained a `required` `AnchorId`
  property. External code that constructs either record with an object initializer must now set
  it; consumers that only read the records are unaffected.

  **Known limitation.** "Effective" is a shorter cascade than the render oracle: it excludes the
  numbering-level `w:pPr` and the table-style/`w:tblStylePr` layers, so a list item whose indent
  lives only in `w:abstractNum/w:lvl/w:pPr/w:ind` reports `LeftIndentTwips = 0`, and a run bolded
  by a `firstRow` table style reports `Bold = false`. Both are pinned by tests and stated in
  `docs/architecture/docx_mutation_api.md`; `GetListMembership` exposes the real numbering
  indentation in the meantime.
- **Style/formatting introspection is annotation-independent and cycle-safe.** Three fixes to the
  new read surface: effective paragraph properties no longer vary with whether
  `ListItemRetriever`'s (lazily cached) list annotations happen to be present, so `GetFormatting`
  answers the same for a projected and an unprojected session and cannot be changed by an
  intervening `GetListMembership`; `ST_OnOff` values now parse through `PtUtil.ToBoolean` — the
  parser the renderer uses — so `w:val="False"` reads as false instead of true, and a value outside
  `ST_OnOff` reads as unknown instead of true (this also aligns `IsDefault`/`SemiHidden`/
  `QuickFormat` and the default-paragraph-style lookup with the resolver that consumes `w:default`);
  and `FormattingAssembler`'s four `w:basedOn` walkers gained cycle guards, so a document declaring
  `A basedOn A` (or `A → B → A`) no longer hangs `ListStyles`, which rolls up every style in the
  catalog including ones no content references.
- **Block renders stamp `data-source-anchor-id`.** `RenderBlocksHtml` / `RenderBlockHtml` now carry
  the canonical source anchor id from the original package onto the rendered block, matching the
  full render. Previously only the full and paginated renders emitted it, so an incrementally
  re-rendered block silently dropped the addressing attribute `npm/src/pagination.ts` resolves page
  citations by. The id is carried, never re-derived from the throwaway shell — a shell hoists note
  paragraphs into its body and would otherwise stamp a `body`-scoped id onto footnote content.
- **`DocxSessionSettings.UndoMemoryBudgetBytes`** (wire `undoMemoryBudgetBytes`,
  Python `undo_memory_budget_bytes`) — an approximate ceiling on the memory held
  by undo/redo snapshots, default **128 MiB**. `UndoDepth` never bounded memory:
  each undo step is a deep clone of every snapshot-scoped part, so one step costs
  whatever the *document* costs. Measured on `TestFiles/NVCA-Model-COI.docx`
  (**144 KB on disk**), a single snapshot retains **≈7.5 MB** — so the previous
  default of 50 steps was **≈375 MB of live DOM for a 144 KB file**, with nothing
  bounding it. The ring now evicts oldest-first (redo before undo) until under
  *both* bounds. Set to `0` for the previous depth-only behavior, which also skips
  measurement entirely.
- **`DocxSession.UndoCount` / `RedoCount` / `UndoMemoryBytes` /
  `UndoHistoryTrimmedForMemory`** — makes the bound observable. The last is sticky
  and reports that the *budget*, not the depth, discarded history, so an editor can
  explain why undo stops short of the configured depth instead of appearing broken.

### Breaking changes
- **`RevisionListEntry` is a required-init record, not a positional one (#455).** It was
  `RevisionListEntry(string Id, string Type, string Author, string? Date, string Text,
  string? AnchorId)`; it is now an init-only record with the members above. Positional
  construction and deconstruction no longer compile in .NET, and in Python the fourth
  positional argument of `docx_scalpel.types.RevisionListEntry` is now `family`, not
  `date` — a positional caller silently rebinds. Construct by name on both surfaces.
- **Revision ids changed format from `revNNN` to opaque `rev2-<hex>` (#455).** Ids are
  emitted in the new format only. A legacy `revNNN` id is still accepted as *input* to
  `AcceptRevision`/`RejectRevision` when it identifies exactly one live revision, and fails
  with `RevisionNotFound` otherwise. Do not pattern-match, sort, or derive `w:id` values
  from an id — re-list and use the value verbatim.
- **MCP `docxodus_track_changes` `accept_all`/`reject_all` return a full `EditResult`,
  not `{"success": true}` (#455).** Callers now get `modified`/`removed` anchors and, on
  failure, a typed `error`. The old shape had no failure representation at all.
- **Bulk revision resolution fails closed and has no `force` mode (#455).** Accepting or
  rejecting everything used to be a whole-document `RevisionProcessor` transform that
  always succeeded. It is now the selective resolver over every registry entry and
  refuses the entire operation — mutating nothing — on the first entry it cannot resolve
  safely: a missing or non-numeric `w:id`, one `w:id` shared by two live revisions in one
  part, `w:customXmlMoveFromRange*`/`w:customXmlMoveToRange*` ranges, `w:ins`/`w:del`
  under `m:ctrlPr`, `w:del` on a run's `w:rPr` or a paragraph's `w:numPr`, a `w:sdt`
  envelope whose range topology is not Word's two-pair shape, an unattached
  `w:numberingChange`, or a malformed cell marker. `RevisionProcessor.AcceptRevisions`/
  `RejectRevisions` still handle all of these and stay public, so the previous behaviour
  is reachable over saved bytes. Full table in `docs/architecture/docx_mutation_api.md`.
- **`EditErrorCode.TrackedOperationUnsupported` moved down the enum (#455),** shifting the
  numeric values of about ten members. The wire is unaffected (every transport serializes
  the name), but a .NET consumer that persisted or compared the integer values must
  recompile.

### Changed
- **Manifest inspection limits reach the Python client (#523).** `generate_package_manifest`
  accepts an optional `PackageManifestInspectionLimits`, so a stdio caller can constrain
  inspection of an untrusted package the way the browser export already could.
- **A batch step whose mutation records no edits is now a successful no-op (#458).**
  `MutationBatchStep.Mutation` may return an empty edit-result list; previously the
  step failed with `internal_error` and rolled an atomic batch back. This means an
  edit that matches nothing (for example a `replace_text_range` whose search text
  is stale) no longer blocks the rest of the batch — callers that need "this edit
  must land" semantics should pass `ExpectedMatchCount` (or a preflight) so a
  zero-match step fails explicitly instead of vacuously succeeding. Delivery
  receipts record such steps as portable no-op evidence.
- **A footnote too long for its page now keeps flowing instead of being clipped
  (#489).** Browser pagination reserves at most 60% of the body height for the note
  band. Continuations were already partitioned at whole-paragraph boundaries, but
  one ordinary paragraph taller than the band was still drawn past its clipped
  bottom on every attempt. Footnote paragraphs now reuse the conservative DOM Range
  fragmenter used by body flow, measured inside the exact initial-note or
  continuation wrapper, and drain across as many note areas as needed. Every
  fragment keeps its canonical source anchor for `PageMap` citations. Content with
  unsafe or indivisible layout retains the visible clipped fallback for that one
  element, while later sibling paragraphs continue instead of disappearing inside
  the same clipped band. Documents whose notes fit paginate unchanged.

  Two rules govern the fallbacks. A note is only deferred to the next page when
  that page's band is larger than what the current page can already offer —
  otherwise the deferral evacuates the citing page and buys nothing. And the
  fragmenter's below-line-breaking fallbacks (arbitrary grapheme boundaries, and
  forcing a first fragment that does not fit) apply only to an element that owns
  the whole band; with earlier siblings already packed, a paragraph that will not
  start belongs intact in the next note band rather than cut mid-word. A queue of
  whole notes larger than the band no longer claims the whole page either: the
  band takes what fits and the rest of the page still sets body text.
- **Word's authored footnote separator stories are rendered, and a continued note
  gets the continuation rule (#489).** `w:separator` and `w:continuationSeparator`
  were previously dropped, and the paginator drew its own two-inch rule above
  every note band. Both reserved stories are now converted into the hidden
  paginated-footnote registry as inert, non-addressable templates; the paginator
  clones the normal story where a note starts and the continuation story on a
  page that carries a note over, falling back to typed 2in/full-width rules when
  a document defines neither. Identities are stripped from these repeated clones
  so no id, editor anchor, or `PageMap` anchor is duplicated across pages. A
  marker-only story renders exactly its rule: the pipeline's synthetic
  empty-paragraph run no longer counts as authored content and no longer draws a
  blank line under the continuation rule.
- **The GitHub Pages landing page serves THE DOCX ARCADE on a phone, keeps its
  navigation, and gives the arcade thumb controls** — three fixes to the same
  problem, that the demo's mobile visit was its worst one:
  - *The arcade is the phone default.* Below 620px `docs/demo/index.html`
    mounts the arcade into its frame instead of the plain editor — the same
    shipped `createRibbonEditor` surface, with a game running on one of its
    paragraphs. Nobody drafts a contract on a 390px screen, and the arcade
    demonstrates the engine harder anyway: a document rewritten and re-rendered
    ten times a second, still saving as a real `.docx`. The choice is taken in a
    `<head>` script before first paint, so the copy framing the frame can never
    advertise the demo the page did not mount; `?demo=editor` / `?demo=arcade`
    pin either on any screen, and the page links to the other one both ways.
  - *The nav survives the phone.* Every link but the CTA was `display: none`
    below 620px, which made the landing page a dead end on exactly the device
    that most wants the other demo pages. They are now a horizontal scroll
    strip — six chips, one swipe, nothing behind a second tap.
  - *Floating controls, where the thumbs are.* The arcade's controls moved into
    `docs/demo/arcade-dock.js`, shared by both hosts and sized from its own
    host element rather than a viewport media query (a narrow card in a wide
    page is narrow). Wide, it is the cabinet's original one-bar dock. Compact,
    a slim HUD strip keeps play/pause and pacing, everything else moves behind
    a `⋯` sheet, and a thumb D-pad plus a round **FIRE** button float over the
    bottom corners of the game — replacing a wrapped four-arrow row that ate
    the bottom of the screen and sat nowhere near a thumb. `FIRE` sends
    `Space`, which the old touch row could not send at all: jump in the
    platformer, the weapon in the raycasters, the coin drop on the attract
    screen — Freedoom could previously be walked but never fought on a phone.

    Demo-site content only; no library, WASM, or npm-package surface changes.
- **`DocxSessionSettings.UndoDepth` default 50 → 20.** See above: at 7.5 MB per
  snapshot the old default was a latent OOM in a browser WASM heap. 20 steps pairs
  with the 128 MiB budget so that on a typical document the two bounds roughly
  agree (~17 steps) and the budget takes over as documents grow.
- Wire defaults for session settings are now read from `DocxSessionSettings` rather
  than repeated as literals in each surface. `DocxSessionJson.ParseSettings` and the
  MCP dispatcher had each hardcoded `undoDepth = 50` independently of the .NET
  default — they could (and would) have silently drifted apart.

### Removed
- **BREAKING — the vendored CPOL-licensed PEG parser and the Excel formula path
  built on it.** `Docxodus/PegBase.cs` carried a 2008 third-party header reading
  `Licence:CPOL` (Code Project Open License) beneath the repository's own
  `Licensed under the MIT license` banner — two incompatible license claims in one
  file. CPOL is not OSI-approved and is generally treated as incompatible with MIT
  redistribution, yet the file shipped inside a package declaring
  `<PackageLicenseExpression>MIT</PackageLicenseExpression>`, and neither `LICENSE`
  nor `README.md` disclosed it. Deleted rather than disclosed, because the code it
  supported was unreachable in practice:
  - `Docxodus/PegBase.cs` (2109 lines, CPOL), `Docxodus/ExcelFormula.cs` (833
    lines, machine-generated in 2012 from an `ExcelFormula.txt` grammar that is
    not in the repository, so it could not be regenerated or maintained) and
    `Docxodus/SSFormula.cs` (105 lines).
  - The public namespaces `Peg.Base` and `ExcelFormula` disappear with them
    (`PegNode`, `PegBaseParser`, `PegByteParser`, `PegCharParser`, `PegException`,
    `ParseFormula`, `FileLoader`, `TreePrint`, and the rest — 20 public types).
  - `WorksheetAccessor.FormulaReplaceSheetName` and `WorksheetAccessor.CopyCellRange`,
    the only two consumers, are removed with it. Both had **zero callers** anywhere
    in the library, tests, CLI tools, WASM bridge, npm package or Python client, and
    `WorksheetAccessor` has no test coverage at all. `CopyCellRange` was not kept in
    a formula-unaware form on purpose: copying a range while leaving relative
    references unadjusted silently corrupts a workbook, which is worse than not
    offering the operation.

  No replacement is provided. Callers needing Excel formula rewriting should use a
  dedicated spreadsheet library. This removes public API and so requires a major
  version bump at release time.

### Fixed
- **Inline and margin comment markers are no longer links to nowhere (#563).** The `[n]`
  comment marker always carried `href="#comment-{id}"`, but only `endnotes` mode renders that
  target — inline mode has no comments section, and margin mode's note ids are stripped when
  pagination clones notes into the page margin. Every commented document therefore failed a
  strict (`unsupportedContent: "strict"`) inline/margin standalone export on a
  `fragment_target_unavailable` dangling link, and quietly degraded to an inert marker under
  the default `warn`. The converter now emits the `href` only in `endnotes` mode; in inline
  and margin modes the marker keeps its element name, id, classes and metadata but is not an
  anchor, so nothing downstream has to repair a link the converter knew was dangling.

- **`PreserveInputRevisions` now carries a modified block's own tracked changes through the
  redline (#517).** A comparison whose right side already contained tracked changes silently
  lost them from a *modified* block — the fine renderers emit the accepted view, so accepting
  the redline could never recover the intended final's review state (the WC012-Math shape: a
  `w:ins` nested inside `m:r`, dropped because math is opaque to token diffing and the whole
  paragraph is a modify). Under the flag, a modified right block that carries its own markup
  now lowers to the whole-block replacement pair whose insert side emits the original element
  with the markup intact — nested inside this diff's `w:ins` when it sits within an atomic
  construct, which is exactly what keeps accept-all, reject-all, and selective resolution all
  correct. Preserved wrappers also keep their **original `w:id`s** (first emission; duplicates
  renumber off a counter seeded above every preserved id), so the input's revision identity
  survives into the redline and `RedlineReversibilityVerifier.Prove` now classifies the
  carried revision as intended-final review state instead of failing closed with
  `intended_final_revision_missing_from_redline`. Default (non-Preserve) comparisons are
  unchanged and the fail-closed proof pin for them remains.

- **The markdown projection shows simple-field results (#559).** A `w:fldSimple` — a
  Word-authored `REF`, `STYLEREF`, `SEQ`, or simple `PAGE` field — was dropped by the
  projection's inline walk, so its visible cached result was absent from the markdown/text an
  agent reads, while the HTML conversion rendered it and the flat text the span machinery
  addresses included it (the complex `fldChar` spelling of the same field already projected,
  so the two spellings disagreed). The cached-result runs now project as ordinary inline
  text on both the oracle and the IR emitter, byte-parity preserved; the field stays atomic
  for mutation addressing, exactly like a hyperlink wrapper. Inline `w:sdt`/`w:smartTag`
  carriers deliberately remain projected-out.
- **Even/odd running stories follow the page number again (#536).** The paginated renderer's
  #527 hardening switched even-header/footer selection (and the matching band heights) to the
  page's one-based position in its section, which flips every story of a section that begins
  on an even page — `DB001-Sections` lost ink parity with LibreOffice on three of six pages
  and turned the weekly visual-parity ratchet red on `main`. ECMA-376 §17.10.5 hangs the even
  story on "even numbered pages": the page NUMBER, which keeps counting across a section
  boundary unless `w:pgNumType` restarts it — and a restart moves the parity with it, so a
  section that restarts at 1 on a physically even page shows its odd story, exactly as Word
  treats a front-matter restart. Selection now keys on the displayed page number (first-page
  selection under `w:titlePg` is unchanged), and the multi-section parity case measures ink
  F1 1.0 on all six pages again — the value the ratchet record already pins, so no record
  refresh is needed.
- **Resolving away a document's only footnote/endnote no longer leaves a reference-less
  husk, and the reversibility proof accepts Word's two "no notes" spellings (#516, #552).**
  When revision resolution removed a note's last reference — rejecting the insertion that
  carried it, or accepting its deletion — the note survived in its part as an empty husk.
  `AcceptRevision`/`RejectRevision`/`AcceptAllRevisions`/`RejectAllRevisions` now delete a
  note whose last reference that resolution itself removed (a note that was already
  dangling beforehand is untouched, and the notes part itself always stays — Word never
  prunes one, per the WC/RP oracle corpus). Separately, `RedlineReversibilityVerifier`
  treats an absent notes part and Word's eagerly-created separator-only part as the same
  statement: when strict normalized identity fails, both sides are re-compared with
  separator-only notes parts elided (disclosed via the new
  `separator_only_notes_part_normalized` info finding), while the raw/OPC digests and
  reported package identities stay strict. The WC035 note-appearance pairs now run the
  full RRS004 reversibility corpus contract in all four directions.
- **Semantic diff: a nested table's format-only change now gets typed records (#511).** When a
  table nested inside a cell changed only its formatting — `w:tblPr`, `w:tblGrid`, or a nested
  paragraph's `w:pPr` — and no content hash moved, `SemanticDiff` descended into the cell only on
  a `ContentHash` difference, so the nested table produced no `table.style`/`table.width`/
  `table.properties` record at all; the change was visible only as the outer table's moved
  `formatDigest`. `IrCell` now retains a `FormatFingerprint` (the ordered fold of its blocks'
  format fingerprints, which the reader already computed for the enclosing table) and the cell
  recursion gate consults it alongside `ContentHash`, so format-only nested changes report fully
  on the nested table's own anchor. Unchanged documents still compare with zero extra recursion.
- **`@docxodus/export` docs asserted the font runtime was still unimplemented (#442).** The
  README's "Runtime boundaries" section claimed `--font-directory` fails with
  `unsupported_runtime`, the `original` review profile is fail-closed, and the generated-PDF
  fidelity ratchet is future work — all three shipped. The stale claims are replaced with the
  delivered contracts, a "Reproducible font configuration" section documents the license-safe
  metric-substitute set (Carlito / Caladea / Liberation) with packages and directories for a
  reproducible CLI render, the CLI help no longer marks `--font-directory` as pending, the
  export design doc names the shipped `font_unavailable` warning code instead of
  `font_family_unavailable`, and the `docxodus` README's export-browser section no longer
  claims `original` and `strictFonts` are pending or that the render report is schema v1.
- **Empty paragraphs lost their line box in paginated export (#443).** A paragraph Word serialized
  as an explicit run holding an empty `w:t` — how it commonly writes a blank line, including
  signature-table spacer rows — counted as having content, so it received no placeholder run and
  reached the browser as nothing but a span the converter had already emptied. That left no
  paragraph-mark line box, a canonical paragraph anchor with zero geometry, and a strict PageMap
  rejecting otherwise-valid documents. An empty `w:t` is no longer treated as content, and the
  cleanup that deletes empty spans now recognizes them: it tested `XElement.IsEmpty`, which is a
  serialization property true only of an element with no child nodes at all, so a span built around
  an empty `w:t` carried an empty text node and survived. Both spellings of a blank paragraph now
  produce exactly one placeholder span.
- **Chromium launch failures say why, and an unavailable sandbox is named as a host policy**
  (issue #525). The Node export runtime already attached the real reason a launch failed to
  `DocxodusExportError.cause`, but the CLI rendered only code, phase, message, remediation, and
  detail, so the diagnostic was discarded at the one surface an operator reads. On any host that
  restricts unprivileged user namespaces — Ubuntu 23.10 and later do by default — every export
  failed with "Chromium could not be launched" plus a remediation pointing at the executable and
  its shared libraries, which is exactly where the problem is not. `docxodus convert` now prints
  the whole cause chain to stderr beneath the remediation, unwrapping `AggregateError` so a wrapped
  cleanup failure is visible too, cycle-safe, depth-bounded, and holding its own character budget so
  a long detail cannot crowd it out. The rendering strips every escape sequence and control code
  point except newline, which Chromium's multi-line launch log needs. Causes still reach stderr
  only: they never enter `detail`, `toJSON()`, or the render report. A launch that fails because the
  host denies Chromium's process sandbox is additionally recognized as its own condition and
  remediated as the host policy it is — the runtime never answers it by dropping `chromiumSandbox`,
  so the operator is told which knob is actually theirs. Every surface that reports a remediation,
  the framed host included, carries the corrected wording.
- **Worker resource-limit failures are classified by code, not message text** (issue #523).
  The standalone export matched worker error message wording to decide whether a failure was
  a resource limit, so rewording a message would silently downgrade the typed `resource_limit`
  code the export contract requires. Worker responses now carry a machine-readable
  `errorCode`. The same pass removes redundant work on the export path: page verification
  uses a map instead of a per-page linear scan, runtime assets are verified concurrently
  while still reporting the first failure in declaration order, images decode in parallel,
  the footnote clipping check runs its cheap guard before per-descendant geometry reads,
  `sha256` no longer copies its input, the manifest worker returns only the representation
  its caller requested, and pagination no longer stamps fragment identities that the export
  path immediately overwrites.
- **Live-session package digests are no longer timestamp-dependent** (issue #521).
  The session's checkpoint clone serialized through `ZipArchive`, which stamps entries
  with wall-clock time at 2-second DOS granularity, so two `GetPackageManifest()` calls
  on identical logical content could disagree on `rawPackageBytesDigest` — an
  intermittent CI flake and a hazard for any workflow comparing raw digests across
  independent serializations (idempotent-retry and no-op-preview detection included).
  `ZipPackageOutputNormalizer` now pins every entry timestamp to the ZIP DOS epoch,
  which is exactly what Word writes; `Save()` output already carried epoch stamps via
  `System.IO.Packaging`, so saved-package bytes are unchanged.
- **Package manifests recognize the content-type spellings Word writes, and VML parts
  digest as XML** (issues #512, #513). Five allowlist/counter spellings in the manifest
  generator could never match a real package — glossary
  (`wordprocessingml.document.glossary+xml`), commentsExtended/commentsIds/people
  (`wordprocessingml.*`, not `vnd.ms-word.*`), and stylesWithEffects
  (`vnd.ms-word.stylesWithEffects+xml`) — silently disabling documented whitespace
  suppression, glossary story facts, and the threaded-comment/people facts (always zero).
  `IsXml` additionally recognizes the declared `vmlDrawing` content type (the one OOXML XML
  part type without a `+xml` suffix), so a reindent-only VML resave no longer flips the
  normalized semantic identity. The semantic package detector consumes the same corrected
  vocabulary for its opaque-part whitespace policy, so the manifest normalizer and the
  semantic diff can never disagree about the same part bytes.
- **Selectively rejecting a move no longer strands the moved text as orphaned
  `w:delText` (#515).** The comparison engine serializes move-source runs as `w:delText`
  inside `w:moveFrom` (delete-grade, unlike Word's own `w:t` serialization), but
  `AcceptRevision`/`RejectRevision`'s resolver only performed the `delText → t` restore
  when unwrapping a surviving `w:del`. Rejecting an engine-generated move therefore
  removed the `w:moveFrom` wrapper and left its payload behind as schema-orphaned
  `w:delText` — on `WC004-Large` ~1.3k characters of restored text were missing from the
  rejected body, and the stranded nodes then surfaced as phantom unresolvable "delete"
  revisions in the live registry. A surviving `w:moveFrom` now restores its payload
  exactly like a surviving `w:del`, matching `RevisionProcessor.RejectRevisions`'s
  unconditional `delText → t` transform; Word-authored moves (already `w:t`) are
  unaffected. Covered by the WC004-Large rows of
  `RedlineReversibilityFixtureSweepTests` (RRS004 content round-trip + RRS005 regression
  pin on the corruption shape).
- **npm — a Web Worker call no longer detaches the caller's `Uint8Array`.** Every
  `createWorkerDocxodus` entry point that takes document bytes (`convertDocxToHtml`,
  `compareDocuments`, the session opens, and the new `generatePackageManifest`) put the
  caller's own buffer in the worker `postMessage` transfer list, so after the call the array
  the caller still held was zero-length and unusable — and when the caller passed a
  `subarray`, the transfer additionally shipped the unrelated prefix and suffix bytes of the
  backing buffer. Worker requests now transfer an exact-view clone. The cost is one copy per
  call; the previous behaviour silently destroyed caller-owned data. Covered by
  `npm/tests/worker.spec.ts` ("Uint8Array subviews are cloned exactly before transfer").
- **Fuzzing harnesses for the package-manifest parser (`tools/manifest-fuzz/`).** A
  self-contained feedback-driven havoc fuzzer, an AFL++/SharpFuzz coverage-guided harness, and
  a full-oracle corpus replayer, all enforcing the generator's contract on arbitrary bytes:
  never throws, byte-identical determinism, no caller-buffer mutation, canonical JSON always
  parses, no hangs — under both default and adversarially small safety limits. The baseline
  campaign (~484M executions: 180.4M havoc + 303.5M coverage-guided + a 24,049-input frontier
  replay) recorded zero contract violations; the README documents the runbooks and the
  positive-control procedure that validates the crash detector itself.
- **Tracked insertions are part of a paragraph's visible text, so an agent can re-find
  the edit it just made.** Under `TrackedChangeMode.RenderInline`, text a mutation
  inserted landed inside `w:ins`, and `InlineRuns` — the shared walk behind the flat
  string that every offset-addressed op works over — did not descend into it. The result
  was a split brain on a document the caller had just edited: `Project().Markdown` and the
  anchor's `TextPreview` showed the new text while `Grep`/`docxodus_search` returned
  nothing, `ReplaceTextRange` could not address it, and `apply_format_by_substring`
  reported `offset_out_of_range`. Insertions arriving from Word were skipped the same way,
  so a search over an incoming redline silently missed every inserted span.
  `w:ins`/`w:moveTo` now join the transparent inline containers; `w:del`/`w:moveFrom`
  deliberately do not, because their content is `w:delText` and deleted text is not
  visible text. Making that text addressable also made it a valid target for ops that insert
  or partition a paragraph's top-level children, and those cannot split a revision wrapper —
  so a note/endnote citation and `SplitParagraph` whose requested offset is strictly inside
  `w:ins`/`w:moveTo` now fail closed with `unsupported_inline_boundary`. Both previously
  reported success after silently landing at the wrapper's far edge. Offsets exactly at a
  wrapper's edge remain exact, and image insertion already refused this case.
  Coverage: `DocxSessionSurgicalTrackedChangesTests` DS409-DS410,
  `DocxSessionNoteAuthoringTests` DS339-DS340, and `DocxSessionPR131RegressionTests`
  DS095-DS096. Found by the issue #435 acceptance smoke and its PR #491 adversarial review
  (`tools/mcp-server/smoke/epic-435-workflow.json`).
- **An ordinary Word picture is no longer reported read-only because of a rendering
  hint.** The blip-extension guard that refuses to replace a picture holding a second
  image payload (an SVG `asvg:svgBlip` beside its raster fallback, or an artistic-effects
  `a14:imgProps/a14:imgLayer` original) tested for the mere presence of an `a:extLst`. Word
  writes `a14:useLocalDpi` on most inserted pictures, so `ListImages` reported
  `canMutate: false` and `ReplaceImage`/`RemoveImage`/`SetImageDimensions` refused with
  `UnsupportedImageMarkup` on a large share of real documents. The test is now structural —
  an extension is a second payload only when it names its own image relationship — which
  keeps every genuine dual-payload picture refused.
- **Tracked `DeleteRange` / `DeleteSection` no longer hard-remove block content
  controls.** Block `w:sdt` envelopes now use Word-native paired custom-XML
  deletion ranges while paragraphs, tables, and nested controls are marked
  recursively, so accept removes the selection and reject restores locked or
  data-bound controls intact. Anchors retained beneath a control are reported as
  `Modified`, structural fall-through anchors as `Removed`, and ranges containing
  unsupported `w:customXml` wrappers fail atomically with
  `IncompatibleElementType` instead of silently deleting them. Paragraph-mark
  revision properties are also inserted in schema order for styled headings.
- One undo step is now always retained even when a single snapshot exceeds the whole
  budget, so undo cannot silently become unavailable on exactly the large documents
  where a mistaken edit is most expensive to lose.
- **A `DocxSession` mutation that threw partway through left the document
  partially mutated, permanently** — and reported it to the caller as an
  ordinary failed `EditResult`. Every mutation records a pre-op snapshot before
  entering its `try`, but 34 of the 44 `catch` blocks discarded that snapshot
  (`_ = _history.PopForUndo()`) instead of applying it, so whatever the op had
  already changed survived AND the record that could have reversed it was thrown
  away. Eight more restored without checking `PopForUndo().ok`, which would
  dereference a null snapshot and throw an unhandled `NullReferenceException`
  straight out of the typed-error envelope. All 44 now route through one
  `RollbackFailedOp()` helper that restores the pre-op snapshot and cannot throw.
  The failure is reachable from ordinary input, not just synthetic faults: a
  stray `U+0000` or unpaired surrogate in a markdown payload — routine in LLM
  output and pasted text — throws from XML writing deep inside
  `InsertFootnote` / `SetHeaderText` / `AddComment` / `ReplaceText`, after those
  ops have already created parts. Clean-failure paths (an op that detected a
  problem and returned *without* mutating) still discard deliberately, so a
  rejected edit cannot evict a real one from the bounded undo ring.
  `DocxSession.LastRollbackError` is new and is non-null only in the one case
  that remains unrecoverable — the rollback itself failing — signalling that the
  session should be reopened from bytes.
- **`Undo()` did not revert writes to `settings.xml` or `styles.xml`.** Snapshot
  membership followed what the *projector reads* rather than what *ops write*, so
  both parts sat outside it. Undoing the first `InsertFootnote` in a document left
  the `w:footnotePr` settings declaration and the generated
  `FootnoteText`/`FootnoteReference` styles behind forever; the same applied to
  `EnsureHeaderFooterVisible`'s `w:titlePg`/`w:evenAndOddHeaders` and to
  `AddComment`'s `CommentText`/`CommentReference` styles. Both parts are now
  snapshot-scoped. The numbering part remains deliberately excluded — list ops are
  additive-only by design, which is what keeps undo correct without it.

## [9.9.0] - 2026-08-13

### Added
- **Visual-parity evidence framework: LibreOffice reference-version contract,
  Word-reference evidence store, and reduced environment cases** (issues #402,
  #403, #404) — the disposition system now has a contracted evidence chain on
  both sides of the comparison:
  - *Reference-version contract (#403)*: the benchmark is contracted to
    **LibreOffice 25.8** (declared once in
    `npm/tests/visual-parity/environment-contract.ts`), asserted at run start
    with install guidance mirroring the font contract's failure mode — an
    out-of-contract host fails in the first second naming the TDF build to
    install, the bundled-font removal step, and the known cross-version
    rendering differences (starting with the 24.2-vs-25.8 footnote separator).
    CI installs that exact TDF build instead of unpinned `ubuntu-latest` apt
    (which carries 24.2 — every scheduled run would have died at the
    fingerprint check after twenty minutes of rendering). The ratchet's
    environment fingerprint additionally gains the **Poppler major.minor**
    (record schema 2): pdftoppm sits between the reference PDF and every
    recorded number. The pure ratchet spec proves the failure message and the
    contract/record/CI agreement on every pull request without LibreOffice.
  - *Word-reference evidence store (#402)*: `word-reference.json` is a
    committed, numbers-only record of what Microsoft Word renders for each
    corpus fixture — page counts, page geometry, ink extents, named per-case
    measurements, and the Word/OS versions used; never binaries, never an
    image corpus. The only manual step is exporting each fixture to PDF with a
    licensed Word; `npm run capture:word-reference` automates everything
    downstream under the benchmark's own 96-DPI Poppler contract and shared
    ink model, optionally recording advisory three-way comparisons against a
    benchmark run. Dispositions cite recorded data via
    `disposition.wordEvidence`, which the pure spec refuses unless the cited
    case is actually measured. All 21 corpus cases are seeded `pending`; the
    capture procedure is documented in
    `npm/tests/visual-parity/WORD_REFERENCE.md`.
  - *Reduced environment cases (#404)*: `visual-parity-reductions.spec.ts`
    (now part of `npm run test:visual-parity`) reduces the three
    environment-attributed cases to minimal generated documents measured
    identically in both engines: `landscape-section` to paragraph pitch
    (uniform in both engines, 29 vs 30 px/paragraph — a 1 px/line same-font
    line-box delta), `inline-image` to extent fidelity (the declared
    `wp:extent` renders exactly in both engines; following text resumes 20 vs
    16 px below), and `tracked-deletion` to the heading line box (identical
    24 px advance; 17- vs 19-row glyph rasterization spread). Each corpus
    disposition now cites its reduced case instead of a whole-fixture
    impression.
  - The ratchet record was refreshed from a clean full-corpus run in the
    reconstructed contract environment, which reproduced all untouched cases
    to five decimal places and re-measured the eight cases moved by renderer
    PRs #417–#421 (which had landed without a record refresh); their corpus
    dispositions were re-triaged from the new evidence.
- **Chart families beyond clustered bar/column render from cached data**
  (issue #411) — `WmlToHtmlConverter` now projects stacked and percent-stacked
  bar/column charts, pie and doughnut charts (per-point `c:dPt` colors,
  `c:firstSliceAng`/`c:holeSize` honored), line charts, and area charts into
  inline SVG, with 3-D variants (`bar3DChart`, `pie3DChart`, `line3DChart`,
  `area3DChart`) rendered as their 2-D projection. Previously everything but a
  clustered `c:barChart` rendered as a blank extent. Line-series colors are
  read from the stroke (`a:ln/a:solidFill`), date axes (`c:dateAx`) format
  cached serial day numbers as dates, dense category axes thin their labels to
  fit, and a percent-stacked value axis pins at 100% with `%`-suffixed ticks.
  New SVG element classes `docx-chart-slice`, `docx-chart-line`, and
  `docx-chart-area` join `docx-chart-bar`; `data-chart-type` gains
  `column-stacked`/`bar-stacked`/`-percent-stacked` suffixes plus `pie`,
  `doughnut`, `line`, and `area`.
- **Visual-parity corpus second wave** (issue #400) — nine new tracked cases
  covering the shapes one-fixture-per-category coverage had left invisible:
  stacked/pie/line chart families, floating square- and tight-wrapped images,
  nested tables, a continuous two-column (`w:cols`) section, endnotes, and a
  realistic legal contract (cached TOC + multilevel heading numbering +
  (a)/(i) sub-clauses + cached REF cross-references + signature table). Five
  fixtures are existing tracked TestFiles; four are authored deterministically
  by `TestFiles/VP/make-vp-fixtures.py` and committed under the same blob-hash
  guard. All nine entered `unattributed` (strict-gating) and were triaged from
  the first measured run — surfacing three renderer bugs (continuous section
  breaks render as page breaks and `w:cols` is ignored, #413; the paginated
  print layout drops the endnotes section, #414; the list-number suffix tab
  overshoots the declared text indent to the next default tab stop, #415) and
  pinning the chart-family (#411) and floating-image text-wrap (#412) feature
  gaps as measured, non-hidden corpus results. `ratchet.json` now records all
  21 cases; BASELINE.md carries the first measured run and each case's triage.

### Fixed
- **The arcade/observatory canvas no longer tilts on devices whose monospace
  font lacks the block-drawing glyphs** — the follow-on to the wrapped-rows
  fix below, and the same root question asked of the other grid property. The
  canvas is a 92×26 character grid, which holds only if every cell advances
  the same width; that is a property of the font the device actually resolves,
  which the document cannot state. The art draws with box drawing (U+2500…),
  block elements (`█ ░ ▓`) and geometric shapes (`▶ ◀ ▲ ▼`), and Android's
  monospace face covers none of them — each lands in a PROPORTIONAL fallback
  (Noto Sans Symbols 2), displacing every cell after it by a different amount
  on each row. Measured on the phone rig with that font situation reproduced:
  the attract screen's five block-letter rows come out 8–13% wider than their
  neighbours, a **12.1-cell** spread — the title card's `X` reads as a `K` and
  the logo smears off the right edge, worst exactly where the art is densest.
  The canvas is now pinned to a self-hosted 17 KB subset,
  `docs/demo/fonts/docxodus-canvas-mono.woff2` (DejaVu Sans Mono 2.37 under
  the Bitstream Vera license, renamed as that license requires; built and
  hash-pinned by `docs/demo/tools/build-canvas-font.sh`), whose every one of
  538 codepoints advances identically — taking the platform out of the loop
  rather than hoping its fallback matches. Same spread with the pin: **0.01
  cells**. `createCanvasPin` also neutralizes kerning, ligatures and inherited
  letter/word spacing, which are per-glyph adjustments a grid cannot survive
  either; the bespoke `npm/examples/ascii-animation.html` driver, which never
  called the pin, now does (so it gains the wrapped-rows guarantee too). The
  saved `.docx` is unchanged and still says Courier New — this is a display
  pin, and Word has the real font. Guarded on both sides: a headless test
  (`docs/demo/tools/canvas-font.test.mjs`, in `npm run test:demo-logic`)
  drives all four phenomena, the whole title-card sweep and all three
  cartridges and asserts every character they can draw is inside the shipped
  subset, and `demo-arcade-mobile.spec.ts` asserts sub-0.1-cell row-width
  spread on the attract screen and all three cartridges, one-cell advance for
  every character actually on screen, and — as the control that keeps the
  others honest — that aborting the font request restores the >1-cell tilt.
- **Paragraph space-before is suppressed at the top of later pages, matching
  Word's section-aware pagination rule** (issue #428) — page placement now
  treats converted `p` and outline-heading blocks as Word paragraphs: the
  authored top spacing participates in normal same-page margin collapsing and
  remains visible on the first page of a document or section, but the clone
  placed first on a later page drops that spacing. Natural overflow,
  `w:pageBreakBefore`, hard-break markers, and keep-with-next measurement all
  use the same decision, while table and other non-paragraph margins remain
  untouched. A generated four-page DOCX pins natural and forced breaks,
  section-start and same-page controls; direct paginator probes cover outlined
  headings, hard breaks, and non-paragraph scope. On the tracked legal contract,
  page-2/page-3 first ink moved from row 115 to 99 px (Word 100, LibreOffice 99),
  and a same-environment filtered rerun improved mean SSIM 0.69467 → 0.72874
  and mean ink F1 0.57203 → 0.70363.
- **Cached TOC field-result hyperlinks now match Word's black, undecorated presentation**
  (issue #427): measured Word output contains zero blue pixels across the cached TOC entries even
  though their runs reference the blue, underlined `Hyperlink` character style; an ordinary link
  with the same styles remains decorated. The converter now derives that exception from the
  existing complex-field annotations: `FieldRetriever` recognizes `TOC`, and run styling removes
  only hyperlink color and underline while inside its cached result. A generated DOCX and a native
  cross-paragraph field-scope test prevent paragraph-style, anchor-name, and blanket-link fixes.
- **Mobile Chrome no longer garbles the arcade/observatory animations with
  extra wrapped rows, and the converter opts document text out of
  device-driven inflation** — on Android the demo's 92-cell frame rows render
  WIDER than authored: the system has no Courier New (Chrome substitutes a
  wider monospace face), and mobile Chrome's text autosizer (the system "Text
  scaling" accessibility setting) inflates text-heavy blocks outright while
  the section's authored column width stays fixed. Once a row outgrows the
  6.5in text column, `overflow-wrap: break-word` folds it onto a second line
  and the frame stacks into garbage — extra lines multiplying as the animation
  fills the grid (sparse early frames stay narrow, which is why screens
  "initially looked good" and degraded as the logo swept in). Three layers:
  (1) the converter's always-emitted document-layout CSS now carries
  `-webkit-text-size-adjust: 100%; text-size-adjust: 100%` on EVERY document
  element, not just `body` — the property is non-inherited, and the field
  evidence shows an intermediate-ancestor rule (the ribbon chrome's existing
  `.dxr` opt-out, present in the shipped bundle) does not protect the document
  inside it on a real phone; a word processor never inflates document text —
  the viewport's fit-to-width zoom is the legibility affordance. (2) The
  GitHub Pages demo pages carry the same page-level opt-out so the live site
  (pinned to a released engine) is fixed without waiting for a release.
  (3) The arcade/observatory drivers pin the canvas paragraph to
  `white-space: pre` (keyed to its stable Unid), the mechanism-agnostic
  guarantee: a frame row can never wrap no matter what a platform does to
  glyph widths — the worst case degrades to a clipped right edge instead of
  stacked rows. Guarded by `demo-arcade-mobile.spec.ts` under the new
  `chromium-pixel5` Playwright project, which recreates the phone's condition
  directly (inflates the document text 175% past the column and asserts the
  frame stays 26 line boxes; without the pin it measures 130+), and by HC056
  asserting the new layout-CSS rule.
- **List-number suffix tabs now advance to the paragraph's text indent instead
  of overshooting to the next default tab stop** (issue #415) — the
  unknown-font text-width estimation charges a flat 0.6 em per character,
  roughly double the real width of narrow-glyph list markers like `(a)`/`(iii)`
  (Times New Roman `(a)` at 12 pt is ~0.94 em, not 1.8 em). The overestimated
  marker pushed the computed pen position past the paragraph's text-indent tab
  stop, so the marker's suffix tab resolved to the next `w:defaultTabStop`
  multiple and every numbered clause body started ~0.25" too far right — in
  exactly the hanging-indent legal-list shape (`w:ind w:left="1080"
  w:hanging="360"`) where the marker plainly fits. Generated list-marker runs
  now measure through a character-class-aware estimate
  (`MetricsGetter.EstimateTextWidth`: hairline / narrow / wide / uppercase /
  CJK-fullwidth em classes modeled on Times New Roman & Calibri averages), so
  the suffix tab lands on the text indent whenever the number genuinely ends
  before it — matching Word's rule and LibreOffice's rendering. The
  character-class estimate is deliberately scoped to marker runs: general tab
  layout reserves (stop − estimated following-text width) ahead of the
  browser's real glyphs, and browser fallback fonts advance at ~0.6 em, so the
  flat general estimate must stay put or column-edge tab lines (TOC entries)
  would overflow and wrap; a marker's wrapper is (chosen stop − marker start)
  with a flex tab absorbing any glyph-width error, so only its stop choice
  needs realistic widths.
- **Paginated print layout no longer drops the endnotes section, and endnote
  markers render in Word's default lowercase-roman format** (issue #414) — in
  `PaginationMode.Paginated` the converter emitted `section.endnotes` as a
  body-level sibling of the pagination staging/container divs, but the
  paginator only flows content found inside `[data-section-index]` wrappers in
  the staging area, so the endnotes never reached a page. The section is now
  appended inside the last section div in the staging area and flows onto the
  final page(s) as an ordinary block, the way Word and LibreOffice lay endnotes
  out. Independently, footnote/endnote display numbers were hardcoded decimal:
  the converter now resolves the effective `w:numFmt` (section-level
  `w:footnotePr`/`w:endnotePr` over the settings-part declaration, falling back
  to the spec defaults — decimal for footnotes, **lowerRoman for endnotes**),
  renders citation markers and paginated-registry labels in that format, and
  stamps the notes-section `<ol>` with the matching CSS `list-style-type`.
- **Continuous section breaks no longer render as page breaks, and `w:cols`
  multi-column sections lay out as columns** (issue #413) — the converter now
  stamps the section wrapper with its start type (`data-section-type`, from
  `w:type`) and column geometry (`data-cols`/`data-col-gap`, from `w:cols`,
  with the same geometry applied inline as CSS `column-count`/`column-gap` for
  the continuous view). The paginator groups a `w:type="continuous"` section
  into the page run its predecessor started — provided the page box (size and
  margins) is unchanged, Word's own condition for honoring the break — instead
  of always opening a fresh page, and flows a multi-column section as balanced
  CSS-multicol container blocks that split across pages at block boundaries.
  Unequal explicit `w:col` widths collapse to the equal-column approximation,
  and pages of a merged run carry the leading section's headers/footers and
  page numbering.
- **Body text now wraps around floating anchored pictures** (issue #412) —
  `WmlToHtmlConverter` renders a picture anchored with `wp:wrapSquare`,
  `wp:wrapTight`, or `wp:wrapThrough` as a CSS float instead of an inline
  object, so surrounding lines flow beside the image the way Word and
  LibreOffice lay them out rather than resuming below it. The float side comes
  from a one-sided `wrapText`, the anchor's `wp:align` token, or the
  offset-placed object's center against the governing section's column (or
  page) center; the anchor's `distT/R/B/L` clearances become margins.
  `wrapTight`'s polygon degrades to its bounding box; `wrapTopAndBottom`,
  `wrapNone`, and a centered object keep their previous placement.

## [9.8.0] - 2026-08-11

### Added
- **Arcade attract screen** (`docs/demo/ascii-arcade.js` `introFrame` +
  `startArcade({ intro })`): the Arcade now opens on a title card — "OS LEGAL
  presents DOCXODUS", drawn with an original 7×5 ASCII block font over a
  twinkling starfield, a typewriter credit line, a left-to-right sweep reveal,
  and a blinking "PRESS SPACE TO START" prompt — rendered on the SAME canvas
  paragraph as the games through the per-frame `raw.replaceXml` +
  `editor.refresh()` loop, so the title screen is a Word paragraph too (pause
  it, put your caret in it). Space (or picking a cartridge, or Restart) drops
  the coin into the selected cartridge; `?intro=0` skips it, which the
  cartridge specs use; a dedicated spec covers the reveal, the incremental
  frame path, and the coin drop.
- **A real Doom-format level is playable inside a Word document** — Arcade
  cartridge 3, *Freedoom E1M1* (`docs/demo/freedoom-e1m1.js` +
  `docs/demo/tools/wad2cart.mjs`; demo content only, no library changes). The
  Docx Dungeon's DDA raycaster is now a level-pack player: the same renderer,
  controls, and MAP-panel document round-trip run both the hand-drawn 24×16
  maze and a 126×109 grid rasterized from the Freedoom project's E1M1
  (only that derived data is BSD-3-Clause; Docxodus remains MIT, and the full
  scoped notice plus immutable source provenance are retained in the generated
  module). Demo and test assets are excluded from the npm tarball by an explicit
  runtime-only allowlist and a package-boundary regression check. The MIT
  notices now credit John Scrudato IV for Docxodus work while retaining the
  inherited Microsoft notice required by the upstream license. `wad2cart.mjs`
  parses the classic binary lumps (THINGS/LINEDEFS/SIDEDEFS/VERTEXES/SECTORS),
  rasterizes blocking linedefs at 32 map units per cell (doors/lifts/stairs
  resolved to their player-friendly state — the grid world has no door
  mechanic), flood-fills walkability from the player-1 start, places `§`
  sigils on the level's own keycard/armor/weapon spots and the `*` gate at the
  exit switch, and BFS-proves every objective reachable before emitting.
  Levels larger than the 24×16 MAP band scroll it as a player-following window
  whose typed edits land on exactly the world cells it shows. The level's 45
  single-player monster placements come along as a combat layer: billboard
  enemies (zombieman/sergeant/imp/demon glyphs) with line-of-sight wake
  checks and chase AI, melee damage, a Space-triggered hitscan sidearm along
  the view center, HP/kills HUD, death + respawn-at-start with progress
  kept, sigils that heal — and the MAP band round-trips enemies as entities,
  so typing `&` into the paused document conjures one. Spec
  `npm/tests/demo-arcade-freedoom.spec.ts` proves real play through the live
  `raw.replaceXml` + `editor.refresh()` loop: a BFS autopilot drives the
  game's own input seam to a Freedoom pickup (fighting the monsters that
  find it on the way), and `DOCXODUS_DOOM_MARATHON=1` plays the entire
  level — every sigil, then the exit — to the win banner.
- **Visual-parity regression ratchet** (issue #395,
  `npm/tests/visual-parity/ratchet.ts` + `ratchet.json`): the weekly scheduled run
  now compares itself against a committed, numbers-only per-case record — page
  counts, severity, mean SSIM, worst ink F1, disposition; no images, no paths, no
  artifact hashes — and fails when any case gets worse beyond a documented
  tolerance (0.0005 SSIM, 0.001 ink F1). Previously the run uploaded an artifact
  that expired in 14 days and nothing compared one run against the previous, so a
  renderer regression was caught only if someone downloaded and eyeballed it in
  time. The ratchet is deliberately **broader than strict mode**: strict gates only
  severe cases the renderer owns, while the ratchet covers every case at every
  severity, so a `close` case sliding to `minor` is caught. Because CI installs
  LibreOffice from unpinned `ubuntu-latest` apt (issue #403), the record carries an
  environment fingerprint (LibreOffice major.minor, Chromium major, `fonts.conf`
  SHA-256); a mismatch reports `environment-changed` and demands a refresh rather
  than blaming Docxodus for a reference-renderer release. Refresh the record
  deliberately with `DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1` in the PR that changes
  rendering, so improvements and accepted regressions are reviewed in the diff; a
  passing run still lists every improvement it measured, so a stale record
  announces itself. The comparison layer is pure and `visual-parity-ratchet.spec.ts`
  exercises it on every pull request without LibreOffice or a renderer, which keeps
  "a deliberately introduced regression fails, naming the case and the signal" a
  continuously proven property. The full artifact upload is unchanged.
- **Visual-parity font-substitution contract** (issue #379,
  `npm/tests/visual-parity/fonts.conf` + `font-contract.ts`): font policy for the
  LibreOffice benchmark is now a shared contract instead of a host observation.
  `fonts.conf` pins each declared Office family to a license-safe metric-compatible
  substitute (Calibri/Calibri Light → Carlito, Cambria → Caladea, Times New
  Roman/Arial/Courier New → Liberation), and BOTH renderers load it via
  `FONTCONFIG_FILE` — LibreOffice per subprocess, Chromium at browser launch
  (scoped to the benchmark opt-in in `playwright.config.ts`, so ordinary specs and
  committed snapshots keep host fonts). Enforcement is layered: an fc-match
  assertion fails the run naming the package to install, an in-browser
  canvas-width check catches a browser launched without the contract, and a
  cross-renderer wrapping probe (one generated paragraph per family; line counts
  must match) detects drift the other layers can't — negatively validated, since
  Calibri Light wraps differently without the contract on a stock host. Every
  `summary.json` records the resolved family/file/version set and the contract
  file's SHA-256, so baseline deltas trace to the renderer or a declared contract
  change, never silent host-font drift. On the corpus: nine cases byte-identical,
  `tracked-deletion` improved (the shared Calibri Light pinning succeeds where the
  rejected renderer-only fallback failed), `fields-and-tabs` now measures its
  honest, lower score.
- **Visual-parity attribution dispositions** (`npm/tests/visual-parity/corpus.ts`):
  every benchmark corpus case now carries a reviewed `disposition` — `renderer-bug`,
  `environment`, `reference-deviation`, `unsupported-feature`, or `unattributed` —
  with a mandatory rationale and optional tracking reference, because severity alone
  conflates "our bug" with "different comparison environment" and "LibreOffice
  deviates from the OOXML evidence". Dispositions flow into `metrics.json` and
  `summary.json` (`aggregate.severeByDisposition`, `aggregate.strictGatingCases`),
  and strict mode (`DOCXODUS_VISUAL_PARITY_STRICT=1`) now gates only on severe cases
  the renderer owns (`renderer-bug`/`unattributed`) plus conversion errors, instead
  of failing on every severe case regardless of whose difference it is. New corpus
  entries default to `unattributed`, which gates, so an untriaged severe case cannot
  hide.

- **THE DOCX ARCADE** (`docs/demo/arcade.html` + `docs/demo/ascii-arcade.js`):
  playable video games whose screen is a live Word paragraph inside the shipped
  ribbon editor — the interactive sequel to the DOCX Observatory, and pure demo
  content (no library changes; the page pins the already-published engine).
  Two cartridges: *Pilcrow's Quest*, a side-scrolling platformer starring ¶
  (run/jump/stomp, coins, spikes, flagpole), and *The Docx Dungeon*, a
  Doom-style DDA raycaster whose right-hand MAP panel is part of the same
  paragraph (a red directional marker and a line-of-sight view cone on the
  map track the 3-D camera, so the two views read as one world). A capture-phase listener owns WASD/arrows/Space only while
  playing; pausing (Esc, the dock button, or just clicking the screen) hands
  the keyboard back to the editor with no mode switch. The signature trick is
  document-as-level-source: on resume the driver blurs the edit (the editor
  commits on blur), re-parses the game world from the session's XML, and
  typed terrain — `#` bricks, `$` coins, `^` spikes, `&` gremlins, or any
  letter as a 3-D dungeon wall — becomes real. A box-drawing bezel frames
  every row so the markdown blur-commit can never read a game row as a
  heading/bullet or split the paragraph on a blank line. Each frame remains a
  Unid-preserving `raw.replaceXml` + `editor.refresh()` (one block repainted
  incrementally, ~90–150 runs/frame at 10 fps), every frame is undoable, and
  the ribbon's Save downloads the current frame as a real .docx. The page
  boots on tap inside iframes, ships an embed-snippet copy button, on-screen
  touch controls, and `?cart=`/`?engine=` overrides; spec
  `npm/tests/demo-arcade.spec.ts` drives boot, keyboard steering, the
  pause→type→resume loop, and the save round-trip.

### Changed
- **The `merged-table` visual-parity case is attributed** (issue #399,
  `npm/tests/visual-parity/corpus.ts`): the corpus's last `unattributed` disposition.
  The recorded premise — "the whole perceptual delta is fill/border *color*" — is
  wrong. Both engines paint the identical theme-derived values (`#4472C4` header,
  `#D9E2F3` bands, `#8EAADB` borders), and the style's cached literals equal those
  values exactly under Word's tint formula, so there is nothing to disagree about.
  Horizontal extents match to the pixel; the residual is row **height**, ~1px per row
  accumulating to ~3px, which large solid fills amplify perceptually — which is how the
  case holds ink F1 1.00000 while its SSIM sits at 0.96348. That isolates to font
  line metrics by elimination: no `w:trHeight`, `w:tblCellMar` top/bottom both 0, and
  `w:line="240"` single spacing, so a row IS one font line box. Moves to `environment`,
  leaving the corpus with no `unattributed` case — a tidiness result, not a gating one,
  since strict mode fires only on *severe* cases and this one is `minor`. Also records
  that Docxodus applies table-style conditional formatting from Word's per-row/per-cell
  `w:cnfStyle` hints rather than deriving band membership from `w:tblLook`, so a
  hand-authored table without them renders unshaded.

### Fixed
- **Border colour resolves `w:themeColor` instead of its cached literal** (issue #399,
  `Docxodus/WmlToHtmlConverter.cs`): `w:color` on a border is a *cache* of the last
  theme resolution, not the authority — when `w:themeColor` is present, the theme
  entry plus any `w:themeTint`/`w:themeShade` is what the border is. Shading already
  resolved this way, so a single table style whose fill and border reference the same
  accent colour derived them from two different sources. For any Word-written file the
  two agree (Word rewrites the cache when it applies the theme), so no rendered output
  changes and no benchmark number moves; it changes documents whose theme was swapped
  without a cache rewrite. Found while reducing the `merged-table` benchmark case,
  which could not expose it: a generated table that makes cache and theme **disagree**
  can.
- **Automatic line spacing is a multiple of the font's line box, not of font-size**
  (issue #396, `Docxodus/WmlToHtmlConverter.cs`): OOXML's `w:lineRule="auto"` gives
  line spacing in 240ths of a *line*, where a line is the font's own single-line
  height. It was emitted as `line-height: <n>%`, and CSS percentages resolve against
  **font-size**, so every line was short by the ratio between the font's natural line
  box and its em square — about 19% for Calibri/Carlito, on `w:line="259"`, which is
  Word's own default for documents created since Word 2013. PR #372 had already built
  the correct model (`line-height: normal` plus `calc(1lh * multiplier)` on the
  paragraph's inline children, so nothing font-specific is hard-coded) but enabled it
  only for empty paragraph marks; it now applies to every paragraph, and the dead
  percentage branch is gone. Measured against LibreOffice at 11pt, the line advance
  goes from 15.69px to 19.33px, which is LibreOffice's exactly. Visible wherever the
  laid-out text height *is* a rendered dimension: a `spAutoFit` DrawingML textbox's
  height error falls from −16.0px to +2.7px, and TOC entries that were displaced by
  up to 14.7px now land within 0.12px. Also resolves the "content sits 28px lower"
  reading of the `numbered-lists` case, which was the accumulated error rather than a
  page-margin disagreement. On the visual-parity corpus: **severe cases 7 → 0**, mean
  ink F1 0.848523 → 0.981644, and the strict-gating set is now empty.
- **A superscript or subscript no longer makes its line taller** (issue #396,
  `GenerateDocumentLayoutCss`): CSS counts a `vertical-align: super` inline box when
  it sizes the line box, so a paragraph grew merely because it carried a footnote
  reference — 25.31px instead of 19.42px on the tracked case, pushing every glyph on
  the line down. Word rides superscripts inside the existing line. `sup, sub {
  line-height: 0 }` joins the always-emitted document-layout stylesheet as a third
  Word layout invariant CSS defaults do not match, alongside over-long-word breaking
  and table max-width. This is the same device the converter already used to keep the
  compacted pieces of a `w:br` from contributing a line box.
- **Paginated footnote area sits on the bottom margin line** (issue #378). Word
  anchors the footnote area to the bottom of the text column — the last note line
  ends ON the bottom margin line, notes stack with no spacing of their own
  (FootnoteText is single-spaced, zero spacing-after), and the separator rule is
  drawn about one line above the first note. The paginated note container carried
  web chrome inside the bottom-anchored box — `line-height: 1.4`, 4pt inter-note
  margins, a 6pt separator gap, a trailing `↩` back-reference link, and an
  `N.`-style number label — which together lifted the visible note ink ~13px off
  the margin and drew ink no print renderer shows. The paginated registry now
  renders the bare superscript number and no backref (the web-view `<ol>` section
  keeps both), and the note-area CSS uses `line-height: normal`, zero item margins,
  and a 3pt separator gap. On the tracked `footnote` benchmark case the note-text
  bottom now sits flush with LibreOffice's (row-exact at 96 DPI) and the
  separator-to-note gap matches. Generated-DOCX browser regression
  `npm/tests/pagination-footnote-geometry.spec.ts` pins the note block and the
  separator separately.
- **Paginated headers, footers, and body text sit at the distances `w:pgMar`
  declares.** `w:header` and `w:footer` are distances from the PAPER EDGE to the
  top of the header story and the bottom of the footer story — four independent
  numbers with `w:top`/`w:bottom`, not two nested boxes. The paginator ignored
  both, anchoring the header bottom-aligned above the top margin and the footer
  top-aligned below the bottom margin, which pulled the two stories toward the
  body by exactly `margin − distance` (25 px on a Word-default page) and left the
  top and bottom of the sheet blank. `resolvePageBands()`
  (`npm/src/page-geometry.ts`) is now the single owner of the model — the header
  grows down from `w:header`, the footer grows up from `w:footer`, and the body
  band starts at `w:top` unless the header has already run past it and ends at
  `w:bottom` unless the footer has already climbed above it — and
  `PaginationEngine.getPageBands()` is the one place band placement, the flow
  loop's page budget, and the footnote area's anchor all read, so they cannot
  disagree about where the body ends. A story taller than its margin now pushes
  the body instead of drawing over it, and the note block follows the body's real
  bottom edge rather than the raw bottom margin. The generated
  `pagination-running-content-geometry` regression pins first/even/odd header and
  footer coordinates, the inheriting second section's own distances, band
  disjointness, the overflowing-story case, and a header-less section. Against
  LibreOffice, `DB001-Sections.docx` improves from severe (SSIM 0.99586, worst ink
  F1 0.00000) to close (0.99934 / 0.99797) and `DB005-Headers-With-Images.docx`
  from severe (0.99773 / 0.00000) to close (0.99921 / 0.96486);
  `DB002-Landscape-Section.docx` also improves (0.91746 / 0.50174 →
  0.92607 / 0.60042). `PageDimensions` gains correctly named `headerDistance` /
  `footerDistance`; `headerHeight` / `footerHeight` remain as deprecated aliases.
- **Two tests failed on evidence they never meant to assert — the wall clock and a shared
  runner's spare CPU.** `PreAcceptInputRevisionsTests` and `DocxCompareTests` each compared two
  SEPARATELY PRODUCED packages by raw `DocumentByteArray`, but a DOCX is a ZIP that stamps every
  entry with the time it was written: the two calls normally land in the same timestamp granule and
  agree, and on a loaded runner they sometimes straddle one and differ on a byte unrelated to the
  behavior under test. Both now compare the part set and each part's bytes through the new
  `PackageEquivalence.AssertSamePackage`, which keeps the entire claim — same parts, same content —
  and drops only container metadata; two packages produced four seconds apart fail the old
  assertion and satisfy the new one. `IrAlignerAdversarialTests`' scale guard divided tiny CPU
  samples and repeatedly crossed its 12x cliff on unchanged code (12.41x, then 12.82x after all
  three retry rounds). It now compares per-thread managed allocations for 4x input against the
  same 12x limit: that deterministically catches quadratic materialization without treating
  scheduling or CPU-cache pressure as a product regression.
- **Floating DrawingML text boxes now honor their OOXML anchor geometry in paginated output.**
  The converter preserves page/margin/column/paragraph/line/character origins, offsets and
  alignments, stored extents, relative sizing, wrap clearances, and internal text insets; the
  paginator resolves those values only after the anchor paragraph has landed on a page. Generated
  DOCX browser regressions pin the outer box and text origin across the distinct coordinate bases.
  On the tracked shape fixture, the box moves from 5 px right and 19 px high to exact horizontal
  bounds and a top origin within 1 px (ink F1 0.50719 → 0.63049), without changing inline drawings.
- **Right, center, and decimal tab stops retain their OOXML targets across native and WASM
  rendering.** Tab measurement no longer includes an invented trailing blank; unavailable-font
  estimates are normalized to CSS layout units; empty tabs advance with an explicit box instead of
  an extra nonbreaking-space glyph; and the aligned segment gives browser font-metric drift to a
  flexible tab remainder. Dot/hyphen/underscore leaders now fill that exact remainder with CSS
  rules, including cached paginated TOCs whose `w:webHidden` tab and page-number runs sit inside a
  hyperlink. Hyperlink containers suppress browser-default styling while preserving explicit Word
  run formatting. Generated DOCX and Chromium regressions cover right/center/decimal targets,
  current and following-run widths, dot/no-leader variants, and the tracked cached-TOC print path.
- **Cached clustered column and horizontal bar charts render instead of disappearing.** A
  `c:barChart` drawing with portable cached categories/values is now projected to accessible inline
  SVG at its stored Word extent, including title, legend, value grid, axis labels, gap/overlap,
  theme/default series colors, and DrawingML font sizes. Rendering does not require the optional
  embedded workbook, a JavaScript chart library, or an Office process; unsupported chart families
  continue through the existing fallback path. An independently generated workbook-free DOCX pins
  the semantic output. In the tracked LibreOffice benchmark, `HC043-Chart.docx` improves from a
  blank severe result (SSIM 0.92933, ink F1 0) to close (SSIM 0.98687, ink F1 0.96817).
- **Empty paragraphs measure automatic line spacing from the font's native line
  box.** OOXML `w:lineRule="auto"` is a multiple of the font's single-line
  height, but the converter expressed it as a percentage of CSS font size,
  which under-measures an empty paragraph mark (its whole height *is* that
  single line box). Empty paragraphs now emit `line-height: normal` plus
  `calc(1lh * <multiple>)` on their inline content; populated paragraphs keep
  the percentage form. A 200-document tracked-fixture comparison went from
  199/200 to 200/200 LibreOffice page-count matches, and `DB005-Headers-With-
  Images.docx` now paginates to 5 pages like LibreOffice instead of 4. The four
  `npm/tests/__snapshots__/tabs-visual.spec.ts` baselines whose fixtures end in
  an empty paragraph were regenerated (2–7 px taller).
- **Omitted header/footer references inherit from the preceding section.** Both
  the paginated header/footer registry and `WmlToHtmlConverter.GetDocumentMetadata`
  treated a section that omits a default/first/even reference as having an empty
  story; OOXML instead links that story type forward until it is explicitly
  replaced, so inherited running content went missing on later-section pages and
  was mis-reported in section metadata. `GetDocumentMetadata` also no longer
  throws `NullReferenceException` on a body with no `w:sectPr` at all, which
  `CollectSectionData` explicitly allows (`null` ⇒ defaults).

## [9.7.0] - 2026-08-08

### Changed
- **Editor surface restyled to the OS-Legal house design system**
  (github.com/Open-Source-Legal/OS-Legal-Style): the ribbon chrome
  (`npm/src/ribbon-chrome.ts`, style version 6) moves from the ad-hoc blue/gray
  palette to the house tokens — deep-teal accent `#0F766E`, slate foreground
  scale, white surfaces over a warm `#FAFAFA`/`#F1F5F9` canvas, 6/12px radii,
  and 0.03–0.06-opacity "light touch" shadows; Inter leads the UI font stack
  (hosts load it, system-ui fallback), Save becomes the surface's one primary
  accent button, and the loading overlay trades the neon cyan/violet look for
  the house dark-slate `#0F172A` + teal family with a Georgia serif headline.
  The GitHub Pages demo pages follow: `docs/demo/index.html` is re-skinned to
  the light "warm precision" look (Georgia headlines, teal CTAs, white cards),
  `player.html` to the dark-sidebar variant, and `app.html`'s frame syncs with
  the new desk color. No DOM, `data-dxr` addressing, or behavior changes.

### Added
- **`frame` option on the ribbon surface** (`RibbonOptions.frame`,
  `"card" | "flush"`): `"card"` makes the surface carry its own embed boundary —
  14px rounded corners, hairline border, house shadow, `overflow: clip` so the
  sticky chrome, scrollbars, and loading overlay respect the curve. Default
  `"flush"` for `mountRibbon` (full-bleed hosts unchanged); `createRibbonEditor`
  defaults to `"card"` since it is the drop-into-a-page path (the three
  `docs/demo/` hosts pin `"flush"` — they frame the surface themselves). The
  document scroll area also gains soft top/bottom edge fades (sticky gradient
  veils) in both frames, so content dissolves at the chrome instead of
  hard-clipping.

## [9.6.0] - 2026-08-08

### Added
- **ASCII-animation demo: DOCX as a rendering canvas** (`npm/examples/ascii-animation.html`,
  served with the Playwright webroot). Four procedurally generated natural phenomena — ocean
  swell, pond ripples, a rain squall, and a hearth fire — animate inside a live Word document:
  one paragraph is the framebuffer, each frame is that paragraph's OOXML (colored monospace runs
  plus `w:br` line breaks, paragraph `w:shd` as the sky) swapped in with
  `DocxSession.raw.replaceXml` and re-rendered incrementally with `session.renderBlock`. The
  replacement XML keeps the paragraph's Unid, so the anchor — and therefore the `data-anchor`
  DOM identity — survives every frame; *Save .docx* mid-animation downloads a real Word document
  holding the caught frame. The page seeds its document through the agentic surface
  (`CreateBlankDocx`, `replaceText`, `insertParagraph`, `applyFormat`, `insertFootnote`) and
  reports live per-frame telemetry (replaceXml / renderBlock / DOM-swap ms, run count, fps).
  Scene generators budget color changes deliberately: every color change inside a row is its own
  `w:r` and the converter pays ~1 ms per run, so ink follows smooth bands while glyphs carry the
  detail (an unbanded fire frame costs ~1300 runs ≈ 550 ms renders; banded it sits nearer 150).
  Spec: `npm/tests/ascii-animation.spec.ts` (boots, frames advance on a stable anchor, every
  scene draws, saved bytes round-trip through a fresh session).
- **`DocxEditor.refresh()` — repaint after external session mutations** (npm). Public method for
  hosts that drive the editor's `sessionHandle` directly (an agent pipeline, `raw.replaceXml`, a
  batch import): "the session changed behind the editor's back — reconcile the DOM." Continuous
  mode patches incrementally from the render plan (a Unid-preserving single-block mutation
  repaints just that block); paginated mode, or anything the reconciler cannot prove, remounts.
- **The Observatory inside the real editor surface** (`npm/examples/ascii-animation-editor.html`).
  The same four phenomena animated in the SHIPPED ribbon editor rather than a bespoke exhibit
  page: the loop mutates the editor's own session (`raw.replaceXml` on the canvas paragraph's
  stable anchor) and repaints through `editor.refresh()`, so pausing needs no mode switch — the
  sea is an ordinary editable block, the whole ribbon works on the document mid-frame, Undo
  rewinds the animation frame by frame (every frame is an undoable mutation), and the ribbon's
  own *Save* downloads the caught wave. Clicking the water pauses the loop and drops the caret
  in. Scene generators, frame→OOXML builder, and document seeding moved to the shared
  `npm/examples/ascii-scenes.js` module both pages import. Spec:
  `npm/tests/ascii-animation-editor.spec.ts` (frames advance incrementally — never a per-frame
  remount — on a stable anchor; pause → type via real editing gestures → one saved DOCX holds
  both the frozen frame and the human's edit; every scene draws).
- **Observatory on the GitHub Pages demo** (`docs/demo/observatory.html`, linked from the landing
  nav). A fourth Pages page hosting the same effect through the *published* package: it imports
  `createRibbonEditor`/`DocxSession`/`getWasmExports` from the pinned CDN embed bundle. The
  phenomena are demo content, not library machinery, and are deliberately NOT shipped in the npm
  package: `ascii-scenes.js` lives canonically in `docs/demo/` (the one place Pages needs a
  physical copy — it deploys `docs/` verbatim, no build step), and pretest copies that single
  file into the test webroot for the two `npm/examples/ascii-animation*` pages. The frame loop +
  dock wiring moved into `startObservatory()` in that module, shared by all three hosts.
  Pinned to `docxodus@9.6.0` — the page needs `DocxEditor.refresh()`, which 9.5.0 predates — so
  it shows its boot-failure card until that release publishes, then heals with no further change.
  Spec: `npm/tests/demo-observatory.spec.ts` drives the page fully locally via the `?engine=`
  override (boots the shipped surface, animates incrementally, save round-trips).

## [9.5.0] - 2026-08-08

### Fixed
- **Dragging a block in the editor now shows what is moving and where it will land.** The drag
  handle floats in the page margin, so the natural gesture — press it and pull straight down —
  keeps the pointer in the gutter, where it never crosses a paragraph box. Because the drop
  target was resolved by asking which element the pointer was over, that gesture found no target
  for its whole duration: no drop line, and a release that silently did nothing. Only a drag
  steered into the text column worked, which is why drops appeared to succeed with no preceding
  feedback at all.
  - Drop position is now resolved from the pointer's **vertical position** against the blocks
    (`resolveDropAt`), not from element hit testing, so the gutter counts as being over the
    document. Blocks are measured once per drag; scrolling is corrected by one subtraction.
    Cost is 0.0 ms per `dragover` on a 234-block charter, and the flow no longer carries one
    registered drop listener per block.
  - A block the source cannot legally reach (across a section break, say) still offers no drop
    and draws nothing — the line never snaps to a distant legal boundary the user did not point at.
  - Three signals now run for the whole gesture: the source block dims, a preview chip naming the
    block follows the cursor (replacing the browser's ghost of the 26 px grip), and a 2 px accent
    line with a leading dot marks the insertion point, positioned by `transform` so tracking the
    pointer costs no layout.
  - The line is drawn in the MIDDLE of the gap between two blocks, not on the target's border-box
    edge. A paragraph's `w:spacing` becomes a CSS margin, which sits outside the box, so an
    edge-drawn line underlined the block's last line instead of reading as a gap between blocks.
- **The editor's continuous view lays the document out at its own page width, and zooms to fit a
  narrow window instead of letting content run off the page.** Enlarging a block's font size on a
  phone pushed text and tables past the sheet, where the window clipped them. Three layers were
  missing Word's actual layout model, and all three are now supplied:
  - **Page geometry is emitted in every render mode.** `WmlToHtmlConverter` stamped a section's
    page setup (`data-page-width`/`data-content-width`/margins) only when `RenderPagination` was
    `Paginated`. A section's page setup is a property of the DOCUMENT, so it is now stamped
    always — the continuous view had no way to know the text column its line breaking was
    authored against, and sized the column from the device instead.
  - **Word's table layout is mapped to the CSS one that behaves the same way.** A table with
    `w:tblLayout` fixed (or a `dxa` width plus a `w:tblGrid`) now renders `table-layout: fixed`
    with a `<colgroup>` carrying the grid's column widths, so content wraps INSIDE a column
    instead of widening it. Every table also gets `max-width: 100%` and cells
    `overflow-wrap: anywhere`, because a Word table never exceeds the text column while CSS's
    auto layout grows one to fit its widest unbreakable word.
  - **A word wider than its column breaks rather than overflows** (`overflow-wrap: break-word`
    in the converter's new always-emitted document-layout stylesheet), as Word and LibreOffice
    do and CSS's `overflow-wrap: normal` default does not.

### Added
- **`DocumentViewport` (npm `docxodus`) — the page-view zoom the editor was missing.** It stamps
  each section wrapper with its `w:sectPr` geometry and applies a fit-to-width zoom, so a phone
  shows a whole smaller page rather than a narrower one that breaks its lines somewhere Word
  never would. New `DocxEditorOptions.fitToWidth` (default true) and `columnWidth`
  (`"section"` default | `"fluid"` for the previous host-width behavior), both also accepted by
  `mountRibbon`/`createRibbonEditor`; `DocxEditor.zoom` reports the applied scale. The shared
  geometry primitives live in the new `page-geometry` module (`parseSectionDimensions`,
  `fitScale`, `ptToPx`/`pxToPt`), which `pagination.ts` now also consumes instead of its own
  private copies.

### Changed
- The editor's continuous view renders a sheet exactly one page wide (`.docx-body-flow`, always
  present now — previously created only when header/footer bands were docked) with the
  document's own `w:sectPr` margins as its gutters, replacing the fixed 920px/72px chrome
  padding. Hosts that want the old reflow-to-the-container behavior can pass
  `columnWidth: "fluid"`.

## [9.4.0] - 2026-08-07

### Added
- **Table cell merge/unmerge — `DocxSession.MergeCells`/`UnmergeCells` (issue #340, the
  deferred Stage B of post-insert table editing).** `MergeCells(cellAnchor, rowSpan, colSpan,
  TableMergeOptions?)` writes native `w:gridSpan` for the horizontal extent and
  `w:vMerge` restart/continue for the vertical one, so a title cell spanning a header row or a
  label cell spanning several data rows is finally expressible; `UnmergeCells(cellAnchor)` is the
  exact inverse, restoring one cell per grid column at its `w:tblGrid` width (addressing a
  continuation cell unmerges the whole run). The rectangle must tile the same whole grid columns
  in every row it covers and must not clip a vertical merge entering from above or continuing
  below — a partial overlap is rejected with the new `EditErrorCode.InvalidTableMerge` before any
  snapshot is taken, never silently tearing the grid. Absorbed cells' content is moved into the
  surviving cell by default (`TableMergeContent.Append`, lossless — the paragraphs keep their
  anchors), with `Discard`/`Reject` as explicit alternatives. Wired through
  `DocxSessionOps` → WASM/npm (`mergeCells`/`unmergeCells`) → MCP `docxodus_table`
  (`merge_cells`/`unmerge_cells`); the grid model, anchor semantics for continuation cells and
  the CRUD rules are documented in `docs/architecture/docx_mutation_api.md`.

### Fixed
- **`DocxDiff` edited moves no longer leave replaced text behind after Reject All in Word-compatible
  consumers (issue #359).** `MoveModifyBlock` destinations were fragmented into `w:moveTo` spans around
  nested `w:ins`/`w:del` edits. Docxodus's own revision processor accepted and rejected that shape, but
  Word-class consumers treat the nested revisions independently: rejecting the move restored its nested
  deletion at the destination, stranding the old word after the source paragraph was restored. Produced
  OOXML now represents an edited move as one complete left `w:moveFrom` and one complete right `w:moveTo`,
  so the destination is accepted or rejected atomically. The edit script and revisions APIs still expose
  the `MoveModifyBlock` token-level change; `RenderMoves=false` and `SimplifyMoveMarkup` retain their
  deliberate whole-block delete/insert projections.
- **Table row/column CRUD is now grid-aware instead of cell-index-aware.** `InsertTableRow`,
  `InsertTableColumn`, `DeleteTableColumn` and `SetColumnWidths` addressed cells by their position
  in `w:tr`, which silently corrupted any table containing a merge. They now resolve cells through
  the grid: inserting a row inside a vertical merge extends the run (and never clones someone
  else's merge markup into a fresh row), deleting a merge's lead row promotes the next
  continuation to the new restart, inserting a column through a horizontal span widens it rather
  than splitting the row, deleting a column through one narrows it rather than dropping the cell,
  and a merged cell's `w:tcW` is the sum of the grid columns it spans. `InsertTableRow` also
  carries the reference row's `w:gridBefore`/`w:gridAfter` shape so the new row still lines up
  with `w:tblGrid`.
- **The opaque-table projection reported the wrong column count for merged tables.** A merge is
  precisely what forces the ` ```table ` fallback, and its `cols:` line read the first row's cell
  count — so a 3-column table with a merged header row told the reader `cols: 2`. It is now the
  table's grid extent (the widest row's summed `w:gridSpan`), mirrored identically in
  `WmlToMarkdownConverter` and `IrMarkdownEmitter` so the equivalence contract holds.
- **The block-move surface no longer does document-scale work on every mouse move.**
  Measured on NVCA's published Model Certificate of Incorporation (234 body blocks, 392
  bookmarks, 94 footnotes, 3 section breaks), hovering across a paragraph boundary cost
  **624 ms → 8 ms**, starting a drag (`ValidMoveTargets`) **4.3 s → 20–30 ms**, opening the
  move menu **4.4 s → 12 ms**, and a review-mode tracked move **8.4 s → ~390–640 ms**. A
  direct move is ~150 ms and undoing one ~165 ms, both from ~780/410 ms. New standing
  instrument: `npm/tests/editor-move-latency-bench.spec.ts`.
  - `DocxSession.ValidMoveTargets` rebuilt the block sequence and re-scanned the story for
    every marker pair **per candidate and per side**, materializing a member set per
    cross-block range — quadratic in the block count and linear again in the marker count.
    It now precomputes what is a property of the container (block order, section-break prefix
    sums, each cross-block range as an index pair) once per sweep and decides each candidate
    with index arithmetic. A range survives a move iff its endpoints still bound a window of
    the same width, which is equivalent to the old set comparison because only one element
    moves. `DocxSessionMoveBlockTests` pins that equivalence against the original
    set-membership definition for every (source, target, side) triple.
  - The editor asked the engine for a block's legal destinations, and re-registered one drop
    listener per body block, every time the pointer crossed a block boundary. Hovering now
    consumes a memoized answer and schedules the query for idle time (an immovable block's
    handle withdraws a beat later instead of every hover paying for the query up front);
    drop targets are registered when a drag actually begins. `blockUnitOf` climbs from the
    pointer's element instead of listing every anchored node in the document.
  - A tracked (review-mode) move forced a full remount. It now reconciles: the render plan
    signs every unit with a content hash, so the move source — which keeps its unid while
    gaining the `w:moveFrom` wrapper — diffs as an ordinary in-place substitution. The bench
    and `editor-block-drag.spec.ts` both assert no remount fallback, since a silent return to
    remounting would only show up as a number.
- **Undo/redo no longer re-serializes the whole document.** `RestoreSnapshot` wrote every
  snapshot-scoped part back to its package stream, an XML serialization of the entire document
  per undo; the session reads parts through their cached `XDocument` and both `Save` paths
  flush every projected part before serializing, so restoring the cache is enough. Undo on the
  charter above went from ~270 ms to ~5 ms of engine time. The comment-threading parts
  (`commentsExtended`/`commentsIds`) are snapshot-scoped but *not* in `Save`'s flush scope —
  their ops persist them directly — so they keep the flushing restore; the split is derived
  from the two part enumerations rather than hard-coded, so it stays true if either changes.
  New `OpenXmlPart.SetXDocumentCache` extension is the in-memory half of `PutXDocument`.
- `UnidHelper.ShortHash` no longer allocates a `SHA256` instance, an input byte array, and a
  `StringBuilder` per call. It runs once per block for every render plan and once per element
  for every Unid derivation — hundreds of calls per interactive operation — so the per-call
  garbage was worth removing, though it does not dominate on a desktop runtime. Same bytes in,
  same digest out.

## [9.3.0] - 2026-08-06

### Added
- **Draggable editor blocks with native tracked moves.** A floating drag handle reorders
  top-level body blocks (paragraphs, headings, list items, and a whole table as one unit) in
  the editor's continuous view, via Atlassian's Pragmatic drag-and-drop. The handle doubles as
  a button opening an accessible move menu (up / down / to top / to bottom) with an
  `aria-live` announcement, so pointer dragging is never the only way to move a block. Enabled
  by default in `mountRibbon`; opt-in as `blockDrag` on `DocxEditor`. Paginated dragging is
  deferred — the handle is not mounted there.
- **`DocxSession.MoveBlock(sourceAnchorId, targetAnchorId, position)`** — one atomic,
  anchor-addressed OOXML mutation per move, one undo snapshot. With tracking off it relocates
  the existing element (descendant Unids and relationship ids untouched); with
  `TrackedChangeMode.RenderInline` it emits a named `w:moveFrom`/`w:moveTo` pair for a
  paragraph and Word's row insert/delete vocabulary for a table. `accept ≡ the moved order`
  and `reject ≡ the original order` hold for both, and `ListRevisions` reports a paragraph
  move as ONE selectable `move` revision resolving both sides. Rippled through
  `DocxSessionOps`, the WASM bridge, npm, the stdio host, docx-scalpel and MCP.
- **`DocxSession.ValidMoveTargets(sourceAnchorId)`** — the blocks a block may legally move next
  to **and on which side** (`MoveTarget(AnchorId, Before, After)`), sharing `MoveBlock`'s own
  guards. The sides are separate because landing *into* a cross-block bookmark/comment range
  changes its membership while landing outside it does not, so a target is routinely legal on
  one side and refused on the other. The drag UI registers only valid targets, snaps a drop to
  the legal side when only one is, disables move-menu items with no destination, withholds the
  handle from a block that can move nowhere, and resolves "move to top/bottom" **within the
  source's own region** — a document with section breaks is partitioned into move regions,
  where targeting the document ends could never succeed. WASM/npm only, like the other
  editor-support endpoints.

### Fixed
- **A tracked move no longer duplicates identity-bearing markers.** The destination clone's
  bookmarks get fresh document-unique ids (both copies keep the NAME, so each survives its own
  resolution and every `REF`/`PAGEREF` still resolves), and the move source takes a fresh
  comment id plus a cloned definition — with fresh `w14:paraId`, entries in both threading
  parts, and cloned replies re-pointed at cloned parents. Previously a pending redline over a
  document with bookmarks or comments was schema-invalid (`Sem_UniqueAttributeValue`) and the
  comment appeared twice in Word's Reviewing pane. Mirrors `IrMarkupRenderer`'s
  `NormalizeBookmarks`/`NormalizeComments` step (B) rather than inventing a second dialect.
- **Whole-block revisions mark content, not just structure.** `w:moveFrom` runs now carry
  `w:delText`/`w:delInstrText` (Word's spelling, and what reject swaps back), and a tracked
  table move marks every cell paragraph's content and mark instead of only `w:trPr` — row
  marks alone left the moved-away table's text rendering as ordinary body text.
- **`RevisionProcessor` keeps a coalesced paragraph's identity.** The paragraph built when a
  deleted or moved-away paragraph mark merges two blocks was constructed without its
  attributes, dropping `pt:Unid`; an anchor-stamped render of the resolved document then
  emitted a block with no `data-anchor` — unaddressable, and invisible to the editor's
  incremental reconciler, which diffs the rendered DOM against `ListBlocks`. Identity now
  comes from the same member the paragraph properties do.
- **Drag autoscroll is registered on the element that actually scrolls** rather than on the
  document flow, which did nothing and logged "Auto scrolling has been attached to an element
  that appears not to be scrollable" on every document open. A drop released outside any valid
  target is now announced instead of failing silently, and a refused move announces an
  outcome rather than the engine's OOXML-worded reason.
- **The CDN embed scopes every whole-document render, not just the first.**
  `createScopedEditorExports` wrapped `ConvertDocxToHtmlComplete` and `RenderHtml` but not the
  new `RenderHtmlForReview`, which is what a remount prefers — so a remount inside
  `createEditor`/`createViewer` inserted the converter's unscoped stylesheet and its `body`
  rules restyled the host page (a host `body { margin: 7px }` became `0px`).

## [9.2.0] - 2026-08-06

### Added
- **`mountRibbon` / `createRibbonEditor` — the editor's UI surface is now shipped, not
  hand-written per page.** `DocxEditor` remains the chrome-less engine; the tabbed ribbon,
  anchor rail, table picker and loading overlay that the demo page carried are now
  `npm/src/ribbon.ts` (+ `ribbon-chrome.ts`, which holds the markup and stylesheet), exported
  from `docxodus` as `mountRibbon(container, options)` and from `docxodus/embed` as the
  one-call `createRibbonEditor(container, source?, options?)` — which also boots WASM, scopes
  the converter's document CSS to the surface, and narrates the boot. The three demo pages had
  each grown their own toolbar and drifted; the GitHub Pages demo was advertising a smaller
  editor than the one that ships.
- **Responsive, container-measured chrome.** Density comes from a `ResizeObserver` on the
  surface's own root (`compactBreakpoint`, default 720 px), not a viewport media query — a
  narrow embed in a wide page is narrow. `compact` turns the ribbon into one horizontally
  scrolling strip, hides the rail and hint, docks the table picker to the bottom edge within
  thumb reach, and grows touch targets to 40 px on coarse pointers. **No command is dropped in
  compact.** `chrome: "full" | "compact" | "auto"` pins or measures it.
- **The loading overlay is part of the surface.** It paints before any runtime exists, so a host
  that boots its own .NET runtime can narrate the gap through `ribbon.loader`
  (`stage`/`progress`/`done`/`fail`); `createRibbonEditor` drives its four stages itself.
  `fail()` shows the error with Retry instead of leaving a dead surface. `loader: false`
  removes it.
- **`docs/demo/app.html`** — the full-bleed editor on GitHub Pages, mobile-first, alongside the
  existing landing page.
- **`DocxEditor.sessionHandle`** — the live `DocxSession` handle as a public getter, so chrome
  can report engine state without reaching into a private field.

### Changed
- `npm/examples/editor.html`, `docs/demo/index.html` and `docs/demo/player.html` are now thin
  hosts of the shared surface. The landing page keeps its hero, capability cards, embed dialog
  and Open Graph metadata but its workspace is the real editor; `player.html` pins the compact
  layout and retires its hand-rolled overflow palette (every command is reachable by scrolling
  the strip). `dist/editor.bundle.js` is now built from `ribbon.ts`, so `window.DocxodusEditor`
  exposes `mountRibbon` alongside `DocxEditor`.
- Controls are addressable as `data-dxr="<name>"` and *also* get `id = idPrefix + name`. With no
  explicit `idPrefix` the surface uses bare ids when they are free and generates `dxr<N>-` when
  they are not, so existing selectors keep working and two ribbons on one page cannot collide.

### Fixed
- **Cross-block drag selection now highlights continuously in Firefox.** The browser reapplied its
  native per-`contenteditable` selection after each `mousemove`, visually fencing the highlight to
  the first block until mouseup even though the final range and formatting were correct. The core
  editor now coalesces bridged range updates into the pre-paint animation frame, preserving native
  live highlighting without overlays or surface-specific patches. The physical-drag regression is
  exercised in both Chromium and a dedicated Firefox Playwright project while the button is held.
## [9.1.1] - 2026-08-05

### Changed
- **Editor interactive latency cut 52–98 % per operation.** Measured end-to-end on a real
  document (`HC031`, Chromium/WASM, warm) by the new standing benchmark
  `npm/tests/editor-latency-bench.spec.ts`: text commit 102 → 30 ms, Enter-split
  137 → 41 ms, bold 108 → 52 ms, font size 88 → 34 ms, Backspace-merge 79 → 25 ms,
  insert table 1.49 s → 130 ms, insert row 1.25 s → 85 ms, delete block 1.2 s → 24 ms,
  undo/redo ~1.2 s → 124/54 ms. Five compounding fixes to the editor's hot paths:
  - *The per-block render shell now stays open across renders.* The session-attached
    single-block render (`HtmlConversionOps.RenderTargetsFromShell`) used to re-open its
    cached shell bytes — package open + styles/numbering XML parse — on every keystroke
    commit; the shell `WordprocessingDocument` is now kept open on the session and each
    render replaces only the main part's body document, so the parse and the converter's
    style/numbering resolution caches persist (guarded by the existing formatting
    signature; `FormattingAssembler` additionally caches its style indexes on the styles
    `XDocument` with a style-count guard, and `MarkupSimplifier` skips the settings-part
    rewrite when there are no rsids to strip).
  - *`ListBlocks` now mirrors the renderer for block-level content controls.* A body-level
    `w:sdt` (a TOC is the everyday case) contributes its `w:sdtContent` blocks as top-level
    render units, matching the flattening the HTML converter performs — before this, any
    document containing one diffed as 100 % churn and the incremental structural reconcile
    permanently fell back to the multi-second full remount (insert row / delete block /
    undo / redo each cost a whole-document convert on such documents).
  - *Reconcile substitutions pair by unid first* (positional pairing only as fallback), and
    a same-unid leaf render may swap into a border-wrapped slot when the old node is
    provably from a single-block swap (no `data-render-sig`) — border-changing ops still
    remount.
  - *Per-op fixed costs removed:* `SetParagraphFormat` no longer eagerly rebuilds the whole
    anchor index (pPr writes can't change an anchor's kind/scope/unid), `SplitParagraph`
    derives the new paragraph's kind locally via the projector's `KindFor` instead of a
    whole-document index rebuild, the block-render path resolves body anchors by a direct
    unid walk when the index cache is cold, and the editor's Enter path renders both split
    halves in one batched `RenderBlocksHtml` call.
  - *The session's per-op index-only rebuild no longer flushes parts.* After every mutation
    the anchor-index rebuild re-serialized any part where a Unid was assigned — a whole
    main-part XML write per keystroke that nothing in the session flow reads (ops and
    renders read the cached XDocuments; both `DocxSession.Save` paths flush every projected
    part themselves). The full projection path keeps the flush for external callers.
- **Mouse selection now crosses editable document blocks.** Browsers normally fence a drag at the
  edge of each independent `contenteditable` paragraph. `DocxEditor` now bridges that gesture into
  one DOM range within the same OOXML story while retaining per-anchor edit/commit boundaries.
  Inline and paragraph formatting therefore work across a physical multi-paragraph drag in both
  the comprehensive editor and the embeddable demo. A stable anchor/offset bookmark also preserves
  that selection when a native toolbar field takes focus. Real-pointer Playwright coverage drives
  both surfaces and verifies multi-block formatting through their actual controls.
- **The live demo now opens a purpose-built Docxodus product guide instead of a generic test
  fixture.** The four-page, branded DOCX opens with the literal “Edit this document. It’s real.”
  and teaches through precise editable exercises, control walkthroughs, a preservation matrix, and
  architecture/privacy callouts plus copyable embed
  examples while showcasing styles, shaded tables, lists, links, headers, footers, and pagination.
  `tools/generate-demo-guide.py` keeps the committed binary reproducible, and both demo surfaces
  load the guide from their own GitHub Pages origin.
- **The social demo now looks and behaves like a product surface.** The main GitHub Pages demo
  has a responsive launch page, animated feature-driven WebAssembly loading state, explicit
  local-only privacy messaging, and a full compact ribbon for inline formatting, font size,
  alignment, lists, indentation, table/rule/footnote insertion, history, pagination, and lossless
  DOCX download. A copy-ready embed dialog documents both hosted iframe and native module usage.
  The compact social-player target gets the same animated feature pitch plus a 4×3 overflow
  palette for advanced controls within its 480×480 layout. Loading failures remain visible and
  retryable, motion respects `prefers-reduced-motion`, and `social-demo.spec.ts` now drives
  formatting, history, layout, insertion, pagination, and downloads against the real editor
  rather than checking decorative UI alone.

## [9.1.0] - 2026-08-04

### Changed
- **WASM payload trimmed and precompressed: ~3.2 MB over the wire (was 16.7 MB raw with
  no compression story).** The browser build now IL-trims `Docxodus` and
  `DocumentFormat.OpenXml` to the `[JSExport]` bridge surface (`TrimMode=full`; modules
  never exported to the browser — HtmlToWml, DocumentBuilder, PresentationBuilder,
  SpreadsheetWriter, … — are removed; no exported API changes), drops the timezone
  database and debug maps/symbols, and ships a brotli-11 `.br` sibling for every
  framework asset so negotiation-capable hosts serve ~3.2 MB instead of ~12.9 MB.
  The one reflective path (`PtOpenXmlUtil.GetPackage()`) is pinned by an ILLink
  descriptor and canaried, with `OpenXmlValidator`, in the new browser spec
  `trim-validation.spec.ts`; `build-wasm.sh` now fails the build if the brotli wire
  total exceeds a 4 MB budget. Measured cold open at 50 Mbps: 1.0 s vs 3.3 s before
  (3.2× faster); native `Content-Encoding: br` decode is noise (~45 ms). Full suite
  (312 browser tests + .NET tests) green against the trimmed artifacts. See
  `docs/architecture/wasm-packaging.md`.

### Fixed
- **MCP HTML preserves native tracked changes.** Full and anchored MCP HTML reads now emit pending
  revisions as `<ins>`/`<del>` instead of implicitly accepting them on the renderer's throwaway
  document. The shared full, single-block, and batched session render paths expose one consistent
  review-mode option, without changing the browser editor's clean editing profile. Every
  newly-authored native revision also enables Word's schema-ordered
  `w:trackRevisions` setting (creating `settings.xml` when needed), so Word continues tracking
  subsequent interactive edits. Coverage: DS408 and MCP140.

### Added
- **Inline document preview in MCP Apps hosts (Claude, ChatGPT) — `docxodus-mcp` now speaks the
  MCP Apps extension** (`io.modelcontextprotocol/ui`, spec 2026-01-26). The server advertises the
  extension capability plus `resources` support, serves a self-contained `ui://docxodus/viewer.html`
  widget template (`text/html;profile=mcp-app`; renders under the spec's default no-network CSP),
  and stamps `_meta.ui.resourceUri` — with the documented `openai/*` compatibility aliases for
  ChatGPT's Apps SDK — onto `docxodus_open` and the new **`docxodus_preview`** tool. `docxodus_preview`
  renders a session (or a single block, by anchor id) through the same converter profile as
  `docxodus_get_content format:"html"`, but routes the markup to the widget via result `_meta`
  (`docxodus/html`) so the model-visible result stays a short summary; `docxodus_open` mirrors its
  `{sessionId, path}` result as `structuredContent` so the widget can fetch its first render and
  refresh after edits via widget-initiated `tools/call`. The viewer is dual-host: MCP Apps
  JSON-RPC-over-postMessage (`ui/initialize` handshake, `ui/notifications/tool-*`) and ChatGPT's
  `window.openai` bridge. Also new: a minimal streamable-HTTP transport (`docxodus-mcp --http PORT`,
  single-response `application/json` shape) so the stdio server can sit behind a tunnel for
  remote-MCP / ChatGPT Apps development. Smoke coverage: `tools/mcp-server/smoke/apps_probe.py`
  (both transports, 23 checks) plus a Chromium harness validating the widget against a
  spec-faithful fake host. See `docs/architecture/docx_agent_server.md` ("Inline preview").
- **Embeddable viewer/editor via CDN — `docxodus/embed` (npm).** The published package was
  already CDN-servable (jsDelivr/unpkg expose `dist/` with CORS `*`, `application/wasm` MIME, and
  the `credentials:"omit"` loader patch), but embedding the *editor* still required hand-booting
  `dotnet.js` and assembling the exports object. The new `embed` entry closes that gap with two
  one-call factories: `createViewer(container, source, options?)` (read-only render, converter
  stylesheet + body injected into the container, footnotes on by default) and
  `createEditor(container, source?, options?)` (full `DocxEditor`; no source opens a blank
  document). `source` is a URL string, `Uint8Array`, `ArrayBuffer`, `Blob`, or `File`. Ships in
  three shapes: `dist/embed.js` (plain ESM for bundler users via `docxodus/embed`, shares module
  state with the main entry), `dist/embed.bundle.js` (self-contained ESM for one-tag CDN
  `<script type="module">`, re-exports the entire main API), and `dist/embed.iife.js` (classic
  script, global `Docxodus`). WASM assets auto-resolve from the bundle's own location —
  `import.meta.url` for the ESM shapes, `document.currentScript` for the IIFE — probing
  `<dir>/wasm/` (package/CDN layout) then `<dir>/` (wasm-webroot layout), with an explicit
  `wasmBasePath` option override. Supporting fixes: `initialize()` now clears its cached promise
  on failure so a retry with a different base path is possible (a rejected first attempt used to
  be permanent), and the raw bridge exports are available via `getWasmExports()` (what
  `DocxEditor.open` needs). Each factory mounts into a private inner root and CSSOM-scopes the
  converter stylesheet to that root, so broad document selectors (`body`, `span`, document
  classes) cannot restyle the host page; unscopable document-global `@import`/`@page` rules are
  omitted. New `examples/embed.html` demo; Playwright coverage in
  `tests/cdn-embed.spec.ts` drives every shape from a second CORS-enabled origin
  (`tests/cors-server.py`) so the cross-origin module import, `_framework` asset fetches, and
  auto-detection are exercised in the exact jsDelivr shape — including a real edit/save/reopen
  round trip and host-style isolation from a page that serves no wasm assets at all.
- **Design stub: browser-LLM redlining demo** — `docs/architecture/browser_llm_demo.md`
  captures the "AI redlines a contract entirely in the browser" demo design (markdown-projection
  anchors as the LLM's edit contract, `TrackedChanges: RenderInline` so edits land as native
  revisions, Chrome Prompt API / BYO-key Claude / WebLLM transport options). Stub only — nothing
  implemented, no engine changes expected.
- **Social-embed demo pages — `docs/demo/`** (GitHub Pages-ready: Settings → Pages → main,
  folder `/docs`). `index.html` is the shareable landing page carrying `og:*` meta (LinkedIn's
  card) plus `twitter:card=player` meta pointing at `player.html`, a ~480×480 boot-on-tap editor
  designed for the historical X/Twitter Player Card iframe — engine from jsDelivr
  (`docxodus@9.1.0`), sample document from `raw.githubusercontent.com`,
  nothing self-hosted. Both accept `?engine=`/`?doc=` overrides, which
  `tests/social-demo.spec.ts` uses to drive them fully locally (also exercising the embed
  bundle's wasm-webroot fallback layout). X no longer documents Player Cards, so iframe rendering
  is explicitly experimental and the ordinary landing-page link is the supported path; local
  tests validate metadata and page behavior, not X's external rendering. LinkedIn never renders
  third-party JS, so the card + one click is its ceiling.
- **Comments can target tracked revisions by id (issue #341).**
  `DocxSession.AddCommentToRevision(revisionId, author, markdown, ...)` brackets the exact live
  insertion, deletion, move-destination, or formatting extent returned by `ListRevisions()`.
  Comment markers sit outside revision wrappers, so selectively accepting or rejecting the
  change preserves the comment: its range stays on surviving content or collapses to a point
  when the content disappears, including a selectively removed table row. Unknown and
  already-resolved ids return the existing `RevisionNotFound` error. The mutually-exclusive
  `anchorId`/`revisionId` target is available
  through WASM/npm (`addCommentToRevision`), stdio/docx-scalpel
  (`add_comment_to_revision`), and MCP `docxodus_comment add`. Coverage: DS410–DS417,
  MCP138, browser revision-session tests, and Python comment tests.

## [9.0.0] - 2026-08-03

### Changed
- **DOCX/OPC outputs no longer lose ZIP compression when Word-authored entries carry misleading
  `superfast` deflate hints (#331).** .NET 10 maps those source hint bits to
  `CompressionLevel.Fastest` when update-mode entries are rewritten, so a 25-part fixture grew
  42,336 → 51,491 bytes even though its uncompressed XML grew by less than 1 KB. The shared final
  output boundary now copies package payloads into a fresh archive with an explicit policy:
  `Optimal` for package markup, `SmallestSize` for compressible binary assets, and `NoCompression`
  for entries already stored because deflate provided no benefit. Entry payloads and OPC structure
  remain byte-identical; names, order, timestamps, comments, and external attributes are
  preserved. The pass also owns the Unix permission normalization from #302, avoiding two archive
  rewrites. Byte-exact clone/no-op comparison paths bypass finalization, so they retain their
  existing exact-package contract without paying the recompression cost. Coverage PKG331–PKG334
  reports per-part compressed/uncompressed deltas, pins a material size reduction (42,336 → 37,101
  bytes on the regression fixture), preserves stored media, verifies every uncompressed part byte,
  and covers both `DocxSession.Save()` and the shared `OpenXmlMemoryStreamDocument` output path. The
  policy and CPU/storage tradeoff are documented in `docs/ooxml_corner_cases.md`.
- **README, package metadata and architecture docs rewritten around what the library actually does.** The repo described itself as an "Office XML Redline Engine" — accurate about one of five capabilities, and containing none of the words anyone searches for. Comparison is now one of four peer sections (render / project / edit / compare), each opening with a **real screenshot** captured from the [NVCA model financing documents](https://nvca.org/model-legal-documents/): a redlined voting agreement showing deletion, insertion, token-level substitution and a paired move in one frame; the model charter rendered to HTML with justification, legal numbering and back-referenced footnotes intact; the markdown projection beside the same document's rendered DOM, with one anchor highlighted in both panes to show they share an addressing system. The redline image is genuine engine output end to end — a round of realistic counsel edits applied through `DocxSession`, compared with `DocxDiff.Compare`, rendered with `RenderTrackedChanges`. Also: a "where it runs" table covering the NuGet / npm / PyPI / CLI surfaces (three CLI tools ship, not two), quick starts trimmed to ≤6 lines each so they can't silently rot, and the OpenXmlPowerTools lineage kept but moved out of the lede. Screenshots also land in `ir_diff_engine.md`, `markdown_projection.md`, `docx_converter.md` and the npm package README (which had no mention of the editor, the session API or the projection). Package metadata is aligned for search across all three registries — `Docxodus.csproj` `Description`/`PackageTags`, `npm/package.json` `description`/`keywords`, `pyproject.toml` `keywords`. New `docs/repo-positioning.md` records the GitHub description and topic list a maintainer still has to apply by hand, the keyword→surface map behind them, and how to regenerate the screenshots.

### Added
- **Tracked surgical text replacements (issue #330).** `ReplaceTextRange`,
  `ReplaceTextAtSpan`, `ReplaceMatch`, and `ReplaceInner` now honor
  `TrackedChangeMode.RenderInline` instead of silently mutating run text directly. A
  selection is split at its exact boundaries: untouched prefix/suffix text remains in
  ordinary runs, removed slices retain each source run's formatting under native
  `w:del/w:r/w:delText`, and the replacement inherits the first affected `w:rPr` under
  `w:ins/w:r/w:t`. Envelopes carry the session author, one operation timestamp, and fresh
  revision ids; adjacent selected runs coalesce into the Word-authored multi-run deletion
  shape. Hyperlink and run-level SDT content stays inside its container, while bookmarks,
  comment/permission/proofing markers, and note-reference runs stay live outside the
  revisions. Reverse-offset repeated matches, selective and whole-document accept/reject,
  undo/redo, and direct mode are preserved. Coverage: DS400–DS407, including Office 2019
  schema validation of single-run, multi-format, hyperlink/SDT, and semantic-marker output.
- **List numbering restart — Word's "Set Numbering Value…" (issue #314).** Nothing on any
  surface could write a `w:lvlOverride`/`w:startOverride`, so "restart at 1", "continue
  from the previous list", and "start this exhibit list at 5" were unaddressable — exactly
  the knobs Word's *Set Numbering Value…* dialog exposes, and constant needs in legal
  drafting. `DocxSession` gains `SetListStartOverride(anchor, value)` and
  `ClearListStartOverride(anchor)`. Set resolves the item's `(numId, ilvl)` via the same
  direct-`w:numPr`-then-style-chain logic `SetListLevel` uses, clones the item's `w:num`
  into a DEDICATED instance carrying `w:lvlOverride[@w:ilvl]/w:startOverride[@w:val]`
  (`NumberingFactory.CloneNumWithStartOverride` — the source num is never mutated: it may
  be shared, and the numbering part is not snapshotted, so additive-only is what keeps undo
  correct), and repoints the anchored item plus every FOLLOWING member of its instance —
  a mid-list anchor therefore splits the sequence exactly like Word (earlier items keep
  their numbers, the tail continues from the new value). Clear repoints EVERY member at a
  clone without the override (they move together, so relative continuation is preserved);
  clearing a sequence with no override at the item's level is a successful no-op that
  consumes no undo history. Style-derived members get a direct `w:numPr` materialized
  (ilvl preserved). Negative values → new `EditErrorCode.InvalidListStartValue`. Numbering
  mutations now also strip the `ListItemRetriever` annotations a previous projection
  stamped on live paragraphs (new `ListItemRetriever.ClearAnnotations`), so the restarted
  numbers are visible in the SAME session's projection/labels, not just after save/reopen.
  Rippled through every surface: WASM/npm (`setListStartOverride`/`clearListStartOverride`),
  stdio host + docx-scalpel (`set_list_start_override`/`clear_list_start_override`), and
  `docxodus_list`, which gains `set_start` (with `startValue`) and `clear_start` actions.
  Coverage: DS350–DS356, including a ListItemRetriever-rendered check that the visible
  numbers actually restart and a save/reopen round-trip.
- **Post-insert table styling — column widths, borders, shading, repeat-header row (issue #315
  Stage A).** Once a table existed its shape and content were editable but its presentation was
  frozen at insert time: `columnWidths` lived only on `InsertTable`, borders only as the boolean
  `borderless`, and shading / repeat-header had no op at all — "shade the header row, border the
  table, repeat the header, widen column 1" all dead-ended. `DocxSession` gains four ops, all
  addressed by the same cell-paragraph anchor the row/column CRUD takes and all localized
  `w:tblPr`/`w:trPr`/`w:tcPr` writes with no model implications (the DocxDiff side already
  digests these shells): `SetColumnWidths(cellAnchor, widthsTwips)` rewrites `w:tblGrid` +
  every row's `w:tcW`, sizes the table to the sum (dxa) and pins `w:tblLayout` fixed — exactly
  what inserting with explicit widths produces; `SetTableBorders(cellAnchor, TableBorderSpec?)`
  writes `w:tblBorders` for only the spec's scope (`All`/`Outside`/`Inside`; style/size/color,
  `"none"` writes explicit none edges à la `borderless`), leaving untargeted edges untouched;
  `SetCellShading(cellAnchor, fill, TableShadingScope Cell|Row)` writes `w:tcPr/w:shd`
  (`val="clear"`, Word's plain-fill idiom) on the one cell or the whole row (header-row
  banding), null fill clearing it; `SetRepeatHeaderRow(cellAnchor, bool)` toggles
  `w:trPr/w:tblHeader` (an emptied `trPr` is dropped). New `EditErrorCode.InvalidTableStyling`
  rejects a width list not matching the column count (or non-positive widths), a fill that is
  neither hex RRGGBB nor `auto`, and a negative border size. Rippled through WASM/npm
  (`setColumnWidths`/`setTableBorders`/`setCellShading`/`setRepeatHeaderRow`, `TableBorderSpec`
  + scope types) and `docxodus_table` (`set_column_widths`/`set_borders`/`set_shading`/
  `set_repeat_header_row`) — the stdio host/docx-scalpel never carried the table CRUD surface
  and stays as-is. Cell merge/unmerge (`w:gridSpan`/`w:vMerge`) is Stage B: it breaks the
  rectangular-grid v1 assumption and gets its own design note first. Coverage: DT208–DT216.
- **Letter/roman/parenthesized list formats + whole-range list conversion (issue #313).** The
  list write surface could only produce `bullet` and `decimal`, and only one paragraph per
  call — converting a 3-item list took 3 calls with no guarantee the items landed in the same
  `w:num` instance, and a "(i)/(ii)/(iii)" romanized sub-list (bread and butter in legal
  drafting) was inexpressible. `ListFormat` widens to `LowerLetter`/`UpperLetter`/
  `LowerRoman`/`UpperRoman` plus a `*Parenthesis` variant of every numbered format
  (`DecimalParenthesis` → `(1)`, `LowerRomanParenthesis` → `(i)`, …) — parenthesization is a
  `w:lvlText` concern, not a `w:numFmt` one, so the variants decompose through the new
  `NumberFormats.FromListFormat` into the existing `NumberFormat` vocabulary and
  `NumberingFactory` stays on `NumberFormats.cs` as the single token-mapping owner (each
  format keeps its own stable marker `w:nsid`, so definitions stay find-or-create idempotent).
  New `ApplyListFormatRange(firstAnchor, lastAnchor, format)` applies one format across a
  contiguous sibling run (inclusive, either document order): every member is guaranteed the
  SAME shared `w:num` instance so the sequence numbers stay intact, each member keeps its own
  `w:ilvl`, non-paragraph siblings are skipped, and the whole range is one undo step
  (cross-part or cross-parent anchors → `AnchorsNotAdjacent`). `NumberingFactory.
  EnsureLevelDefined` now synthesizes missing nest levels in the numbering's own format
  (letter/roman/parenthesized carry down) instead of always falling back to decimal. Rippled
  through every surface: WASM/npm (`ListFormat` union widened, `applyListFormatRange`), stdio
  host + docx-scalpel (`apply_list_format`/`apply_list_format_range` — new on that surface —
  with a `ListFormat` enum), and `docxodus_list`, which gains an `apply_format_range` action
  (`firstAnchorId`/`lastAnchorId`) and the widened `listFormat` token set. Coverage:
  DS340–DS347.
- **Selective per-revision accept/reject + markup-native revision listing (issue #318).**
  Tracked-change resolution was all-or-nothing: `RevisionProcessor` over the whole document,
  so "accept the city correction, reject the sentence deletion" — the single most common
  review action — required accept-everything-then-re-apply emulation. `DocxSession` gains
  three ops. `ListRevisions()` enumerates `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`,
  paragraph-mark and table-row markers, and the `*PrChange` format-change family directly
  off the live markup (body, headers, footers, footnotes, endnotes — no accept-all/reject-all
  re-diff, so it is cheap on large documents and reports the markup's TRUE `w:author`/`w:date`
  instead of engine defaults), grouping contiguous same-kind/same-author markup into one
  `RevisionListEntry(Id, Type, Author, Date?, Text, AnchorId?)` per user-visible change — an
  inserted paragraph is one revision (runs + mark), a deleted table row absorbs its cell
  markup, a named move pair is one `move` entry covering both sides. `Id` derives from the
  markup's own `w:id` attributes, so it is stable across calls and across resolution of other
  revisions. `AcceptRevision(id)`/`RejectRevision(id)` resolve exactly one group in place as
  ordinary undoable session mutations (no whole-document `RevisionProcessor` round-trip, no
  session rebind; `EditResult.Modified`/`Removed` name the touched blocks), mirroring
  `RevisionProcessor`'s per-element semantics: unwrap vs. remove, `w:delText` → `w:t`
  restore, paragraph-mark coalescing into the following paragraph, row/table removal, and
  `CT_*Base`-aware stored-property restore for format changes. New
  `EditErrorCode.RevisionNotFound`. Mechanics in `Docxodus/Internal/RevisionOps.cs`; not
  enumerated in v1 (whole-document accept/reject still covers them):
  `cellIns`/`cellDel`/`cellMerge`, content-control ins/del ranges, `numPr` numbering-ins.
  Rippled through every surface: WASM/npm (`listRevisions`/`acceptRevision`/`rejectRevision`
  + `RevisionListEntry` type), stdio host + docx-scalpel
  (`list_revisions`/`accept_revision`/`reject_revision`), and `docxodus-mcp`, where
  `docxodus_track_changes` gains `accept`/`reject` actions taking `revisionId` and its `list`
  action switches to the markup-native listing (stable ids, real attribution, ~3s → ~ms on a
  49-page document). Coverage: `DocxSessionRevisionTests` (DS370–DS389).
- **Tracked run-format mutations in `DocxSession` (issue #319).** `ApplyFormat` now
  emits native `w:rPrChange` markup when the session is in
  `TrackedChangeMode.RenderInline`, including calls routed through
  `ApplyFormatToSubstring` and the `TextMatch` overload. Every changed run keeps the
  requested current properties and archives its own previous `w:rPr` in a schema-last
  change marker stamped with the session revision author, one high-resolution operation
  timestamp, and a fresh revision id; adjacent runs from one call therefore list as one
  `format` revision, while separate adjacent calls remain independently selectable.
  Accept keeps the new formatting and removes the marker; reject (selective or
  whole-document) restores the original per-run properties. Semantic no-ops (including
  equivalent OOXML ordering/boolean/underline/color/size spellings) produce no revision,
  and a failed operation restores the complete pre-call snapshot. Applying another
  tracked format to a run that already has an `rPrChange` preserves that single marker's
  original baseline and attribution instead of nesting/replacing it, so reject-all still
  reaches the true pre-review formatting; formatting back to that archived baseline
  removes the now-empty pending change, including for partial spans. Paragraph-property/
  style changes remain a separate `pPrChange` follow-up. No wire changes. Coverage:
  `DocxSessionRevisionTests` (DS390–DS399).
- **First-line/hanging indent and paragraph spacing on `SetParagraphFormat` (issue #312).**
  `ParagraphFormatOp` could express alignment, a left-indent delta, page-break-before, and
  paragraph borders — but not the other two workhorses of paragraph layout, so the intent
  "give this paragraph a first-line indent and space-after" was unreachable from every surface
  (the nearest op, `indentDelta`, shifts the whole left edge — a visibly different result).
  New tri-state fields (null = leave unchanged, all values twips, documented per-field since
  unit ambiguity here is a real caller trap): `FirstLineIndent` (`w:ind/@w:firstLine`) and
  `HangingIndent` (`@w:hanging`) — one either/or slot the way Word stores them, so setting one
  evicts the other and an op carrying both is rejected; `SpacingBefore`/`SpacingAfter`
  (`w:spacing/@w:before`/`@w:after`); `LineSpacing` + `LineSpacingRule` (`w:line`/`@w:lineRule`
  — `auto` (default) measures in 240ths of a line, 240 = single/360 = 1.5×/480 = double;
  `exact`/`atLeast` in twips). Existing `w:spacing`/`w:ind` attributes not named by the op are
  preserved. New `EditErrorCode.InvalidParagraphFormat` covers the unrepresentable
  combinations (both indents, negatives — the attributes are unsigned, a rule without a
  value). Rippled through every surface off the shared `DocxSessionJson.ParseParagraphFormatOp`
  wire parser: WASM/npm (`ParagraphFormatOp`/`LineSpacingRule` TS types, now exported from the
  package root), stdio host + docx-scalpel (`first_line_indent`/`hanging_indent`/
  `spacing_before`/`spacing_after`/`line_spacing`/`line_spacing_rule` + `LineSpacingRule`
  enum), and `docxodus_format set_paragraph_format`'s `paragraphFormat` schema (units spelled
  out in every field description). Coverage: DS335–DS339, MCP043, python
  `tests/test_paragraph_format.py`. Follow-up: writing an explicit `SpacingBefore`/`SpacingAfter`
  also clears a direct `w:beforeAutospacing`/`w:afterAutospacing` flag (only the matching one) —
  a set flag makes Word ignore the explicit value, so the write would otherwise be a silent
  visual no-op; Word's own Paragraph dialog switches "Auto" off the same way when a typed value
  replaces it (DS340).
- **Mid-session tracked-changes mode switching (issue #304).** `DocxSessionSettings.TrackedChanges`/`RevisionAuthor` were init-only — recording mode was nailed down for a session's whole lifetime, and the only way to flip it was save→close→reopen (losing undo history and, without `persistAnchorIds`, every anchor id). `DocxSession.SetTrackedChanges(mode)` / `SetRevisionAuthor(author)` now switch how *subsequent* mutations are recorded, with read-back getters. Session configuration, not a document mutation: not undoable, and already-applied markup is never touched. Rippled through every surface: WASM/npm (`setTrackedChanges`/`setRevisionAuthor`), stdio host + docx-scalpel (`set_tracked_changes`/`set_revision_author`), and `docxodus-mcp`, where `docxodus_track_changes` gains a `set_mode` action (absent `revisionAuthor` = keep current, `""` = reset to default) that echoes the now-current state. Coverage: DS329–DS334, MCP133–135, python `tests/test_tracked_changes_mode.py`, npm `docx-session.spec.ts`.
- **`persistAnchorIds` reaches the agent server + stdio host (issue #303).** `DocxSessionSettings.PersistAnchorIds` existed in the core but `docxodus-mcp` never wired it through, so every MCP session permanently stripped the `PtOpenXml:Unid` anchor bookkeeping on save and no anchor id could survive a close+reopen (the reopen-to-switch-`trackedChanges`-mode workflow had no escape hatch). `docxodus_open` now takes `persistAnchorIds` (default `false` — today's behavior unchanged), and `docxodus_save` takes a per-call tri-state override: absent → the session's open-time setting, `true` → a single anchor-stable checkpoint from a default session, `false` → a clean deliverable from a session opened anchor-stable. Backed by a new `DocxSessionOps.Save(handle, persistAnchorIds)` facade overload; the same per-call override lands on the stdio host's `save` op and `docx-scalpel`'s `DocxSession.save(persist_anchor_ids=...)` (whose open-time setting already flowed through the settings wire). Narrative in `docs/architecture/docx_agent_server.md` § "Anchor stability across save→reopen". Coverage: MCP004–007, python `tests/test_persist_anchor_ids.py`.
- **Native Word comment authoring on `DocxSession` (issue #300).** `AddComment(anchor, span?,
  author, markdown, initials?, date?)` writes real `w:comment` markup — the
  `WordprocessingCommentsPart` plus the `CommentText`/`CommentReference` styles are
  find-or-created on first use, the character span is bracketed with
  `w:commentRangeStart`/`w:commentRangeEnd` (the same `SplitRunsForSpan` mechanism annotations
  use), and the `CommentReference`-styled `w:commentReference` run lands directly after the
  rangeEnd, so authored comments show up in Word/Google Docs/LibreOffice's Reviewing pane.
  `w:date` is written only when the caller provides one, keeping output deterministic by
  default. `UpdateComment` replaces a comment's body while preserving its identity attributes
  and the last paragraph's `w14:paraId` (Word's `commentsExtended` threading key);
  `RemoveComment` delegates to `DeleteBlock`'s `cmt` teardown, which now also prunes
  `commentsExtended`/`commentsIds` entries keyed by the removed definition and drops
  `w15:paraIdParent` references to it (a surviving reply becomes top-level, never dangling —
  both threading parts joined the content-only undo snapshot scope so the pruning is
  undoable); `ListComments` returns `(DefAnchorId, Author, Initials?, Date?, Text)` in part
  order. Comment-part create/delete is undo/redo-reconciled by `ReconcileCommentsPart` (the
  `ReconcileNoteParts` twin). New `EditErrorCode.EmptyCommentSpan`;
  `CommentMarkerNotSupported`'s message now names the real op. Rippled through every layer:
  WASM/npm (`addComment`/`updateComment`/`removeComment`/`listComments`, plus the TS
  `EditResult` gaining the `annotationId` field the wire already carried and the error-code
  union gaining the annotation trio it was missing), stdio/`docx-scalpel`
  (`add_comment`/`update_comment`/`remove_comment`/`list_comments`), and the MCP server —
  **`docxodus_comment` now authors native comments** (add/update/remove/list, batchable via
  `docxodus_mutations`) while the bookmark + custom-XML annotation overlay moves unchanged to
  a new **`docxodus_annotate`** tool, retiring the "no native Word review-comment threads"
  known-gap note (reply-threading/resolve state is subsequently closed by #317 below). Coverage:
  `DocxSessionCommentAuthoringTests` (DS346–DS364), MCP070–072, browser spec
  `docx-session-comments.spec.ts`, python `tests/test_comments.py`.
- **Native Word comment replies and resolve/reopen state (issue #317).**
  `DocxSession.AddCommentReply(parentCmtAnchor, author, markdown, initials?, date?)` now authors
  a distinct reply definition plus an adjacent `w:commentReference`, preserving Word's native
  shape in which only the thread root owns range markers; `w15:paraIdParent` links each reply to
  its immediate parent, so a nested reply inherits the root range through reference-only parents.
  `SetCommentResolved(cmtAnchor, resolved)` writes `w15:done` (`false` reopens while preserving
  parentage), upgrading legacy flat comments on demand. The engine find-or-creates
  `commentsExtended.xml` and `commentsIds.xml`, stamps deterministic collision-free uppercase
  eight-hex para/durable ids (max+1), and snapshots the parts' relationship topology so undo
  removes first-created parts and redo recreates them with stable relationship ids.
  `ListComments` adds nullable init properties `ParentAnchorId`/`Resolved`, omitted on the wire
  when no extension entry exists while retaining the five-argument constructor/Deconstruct ABI.
  Rippled through WASM/npm (`addCommentReply`/`setCommentResolved`), the stdio host
  and docx-scalpel (`add_comment_reply`/`set_comment_resolved`), and MCP
  `docxodus_comment` (`reply`/`resolve`, where `resolved:false` reopens). Coverage: DS400–DS404,
  MCP073, browser `docx-session-comments.spec.ts`, and Python `tests/test_comments.py`.
- **Incremental structural repaint in the browser `DocxEditor` — structural operations drop from seconds to modern-web latencies.** Every structural op (insert table / row / column, insert footnote or endnote, delete block, undo, redo) used to pay a full remount: a whole-document `RenderHtml` plus a complete DOM rebuild, ~5–6 s on a 346-block, 94-footnote filing template while `save()` was ~300 ms — the re-render was the bottleneck, and separately the engine ops themselves carried ~1 s of per-op projection overhead. Both are gone. Measured on that same document (warm, continuous mode): insert table **225 ms** (was ~6.1 s), insert row **243 ms**, insert footnote **210 ms** (was ~6.2 s), delete block **93 ms**, undo **201–349 ms** / redo **359 ms** (were ~5.6 s), `setPageNumbering` **287 ms** (was ~650 ms); text commit and `save()` unregressed. The mechanism is one op-agnostic reconciler, `DocxEditor.reconcile()`: diff the DOM's top-level unit sequence against the session's render plan (LCS over `unid|contentHash` tokens), keep every unchanged unit's DOM node, render changed/created units in ONE batched WASM call, drop removed units (with their generated wrappers), and renumber footnote/endnote marker chrome positionally. A full remount remains the universal fallback — unsupported bridge, paginated mode, pure list-item insert/remove (sibling numbers shift without sibling XML changing), a substituted list item whose rendered marker drifts, border-`div` regrouping, order violations, or any error — so correctness never depends on the diff being right, and the reconciled DOM is pinned equal to a remounted DOM by spec. New `npm/src/editor-reconcile.ts` (pure diff functions), 19 new Playwright specs (`editor-reconcile*.spec.ts`) including node-identity proofs, remount-equivalence, chrome renumbering and save/reopen fidelity.
- **Session endpoints powering the reconciler** (WASM `DocxSessionBridge` + npm types; all optional on older bundles, the editor degrades to remounts): **`ListBlocks`** — the ordered top-level render units per scope container (each body `w:p` under its projected kind, each table as ONE `tbl` unit, note definitions mirroring exactly what the renderer's notes section shows), each unit carrying a content signature (`UnidHelper.ContentHash`) because in-session an element keeps its unid across edits, so unid alone cannot reveal an undone text edit or a row insert; note ids (reference AND definition) are excluded from the hash since a footnote insert shifts every later note's `w:id` without changing rendered content. **`ListNotes`** — footnotes/endnotes in citation order with their `w:id`, the id↔ordinal authority for renumbering rendered note chrome client-side instead of re-rendering every citing block. **`ListAnchors`** — the projection's anchor index without the markdown payload, replacing a couple-hundred-KB `Project` marshal per repaint. **`RenderBlocksHtml`** — batch block render: N anchors, one throwaway document, one converter run, with each target cloned alongside its real siblings (so `w:contextualSpacing` margins resolve exactly as in the full render) and the live document's `ListItemRetriever` annotations transplanted onto the clones, so **a list item deep in a list renders its true number in isolation** — closing the numbering-continuation gap (M9) for every incremental swap; a table returns with its generated alignment-`div` wrapper, and `fn:`/`en:` anchors return their note paragraphs. The session-attached single-block `RenderBlockHtml` is now a one-element batch, so those fixes apply to every existing swap path too — including a fixed latent bug where a re-rendered citing paragraph silently *lost its citation marker* because the block render profile left `RenderFootnotesAndEndnotes` off.
- **`DocxSessionSettings.EmitMarkdownPatch` (wire key `emitMarkdownPatch`, default `true`).** When `false`, mutation ops return `Patch = null` and skip the per-op scope re-projection that builds it — `ProjectScope` re-projects the whole document per mutation, pure dead weight for clients that re-render from HTML. The browser editor opens its session with it off.
- **Agent-facing DOCX editing server (`tools/mcp-server`, `docxodus-mcp`).** A [Model Context
  Protocol](https://modelcontextprotocol.io) server that lets an AI agent open a `.docx` file
  into an in-memory `DocxSession`, read/search/edit/format it through ten grouped-intent tools
  (`docxodus_get_content`, `docxodus_search`, `docxodus_edit`, `docxodus_format`,
  `docxodus_create`, `docxodus_list`, `docxodus_comment`, `docxodus_track_changes`,
  `docxodus_mutations`, `docxodus_table`) plus three lifecycle tools
  (`docxodus_open`/`docxodus_save`/`docxodus_close`), and save it back — over stdio only, no
  network calls, no telemetry. Every tool routes through the existing
  `Docxodus.Internal.DocxSessionOps`/`DocxDiffOps` facade (the same one the WASM bridge and
  `tools/python-host` use); no new editing logic was added to the core library. New
  `SessionStore` layer (external string `session_id` → handle/path/settings) supports
  `docxodus_save`'s "write back to the opened path" default and `docxodus_track_changes`'s
  `accept_all`/`reject_all` (which swap a session's underlying document for a whole-document
  `RevisionProcessor` transform while keeping the same external session id). `docxodus_mutations`
  batches any of the mutating grouped tools into one call, with a `preview` mode that applies
  every step and then undoes them via the session's bounded undo ring. Ships as the `dotnet tool`
  `docxodus-mcp` (`Docxodus.csproj` gains an `InternalsVisibleTo` grant for it, mirroring the
  existing grants to `docxodus-pyhost` and `DocxodusWasm`). Several Docxodus capability gaps are
  documented rather than papered over: no selective per-author/per-type tracked-change
  resolution (only whole-document accept/reject), and a pre-existing anchor-kind asymmetry
  between `ReplaceCellContent` (needs a `"tc"` anchor) and the table row/column ops (need a
  `"p"` anchor) that the new test suite surfaced. See
  `docs/architecture/docx_agent_server.md` for the full tool contract and gap list. Coverage:
  `Docxodus.Tests/McpServerDispatcherTests.cs` (`MCP###`).

- **Scoped document storage for the agent server (`IDocumentStore`).** Everything the server reads
  or writes goes through one three-method interface — `Resolve` a caller-supplied location to a
  canonical in-scope identifier, then byte-level `Read`/`Write` — with `DocumentStores` as the
  single owner of "which backend, rooted where," built once at startup from the environment
  (`DOCXODUS_STORAGE_BACKEND`, `DOCXODUS_STORAGE_ROOT`, `DOCXODUS_STORAGE_SCOPE`). The point of
  the seam is that a future backend (object storage, a content repository) is a new
  implementation plus one case in `DocumentStores.Create`, with no change to the dispatcher, the
  tool schemas, or any session logic; the interface takes opaque location *strings* rather than a
  filesystem-shaped API so a backend without directories needn't pretend to have them. Today the
  one implementation is `LocalFileDocumentStore`. **Isolation is structural rather than a check:**
  a store is constructed already rooted at its scope and there is no read/write path that skips
  `Resolve`, so a session cannot name a document outside it — relative locations resolve under the
  root, an absolute path is accepted only if the root contains it (which is what keeps
  `~/Downloads/contract.docx` working under the permissive default root while a narrower root
  confines the identical tool surface), and containment is checked against symlink-resolved paths
  with every component resolved from the volume root down. That last part is load-bearing and was
  wrong in the first cut: resolving only the leaf missed `{root}/link → /elsewhere` followed by
  `{root}/link/secret.docx`, whose own leaf is not a link — the regression test for it is
  `MCP125`. Dangling links are detected by reading rather than following the link, so one can't be
  used as a write escape, and segment boundaries are respected so `/srv/base-2` is not inside
  `/srv/base`. Backend and root are operator configuration and never tool arguments — if the agent
  could name a root per call, the scope would be chosen by the thing it exists to contain — and
  `DOCXODUS_STORAGE_SCOPE` comes from whoever launches the process, which is the only trust
  boundary stdio MCP actually has (there is no in-band authentication; the spawn *is* the
  authorization). Because a scope is just a stable path segment, passing the same value next
  session reaches the same documents with no registry, token format, or revocation list; one
  process serves exactly one scope. Session ids became 16 random bytes rather than a counter, since
  the id is the capability. Coverage: `MCP120`–`MCP131` plus `MCP003`.

### Fixed
- **`DocxDiff` now makes automatic list-number cascades visible in redlines.** The IR reader already
  resolved each aligned paragraph's marker on both sides, but the markup renderer discarded that
  difference; HTML therefore showed only plain black final labels after an inserted, deleted, or
  moved list item. Changed markers are now preserved as native
  `w:numPr/w:numberingChange[@w:original]` metadata (with effective numbering materialized for
  style-inherited lists), and the tracked-changes renderer emits a struck old marker followed by an
  inserted new marker. Deleted and moved-from items retain their source-side label instead of a
  counter value recomputed in the merged redline. Coverage: `TrackedChangesNumberingTests`
  TCN005–TCN006 and the deterministic NVCA screenshot fixture, which asserts both marker values and
  revision states before capture.

### Changed
- **`DocxSession` mutation ops no longer rebuild the full markdown projection per edit.** Anchor resolution (`FindAnchor` and the post-apply re-resolution inside every mutation op) now goes through a cached index-only build — same walk, same keys, same Unid assignment as the full projection, minus markdown emission and per-entry `TextPreview`/`AutoNumberPrefix` numbering resolution. `ProjectScope` caches the projection it builds (it runs post-invalidation, so it IS the post-op state). The deterministic Unid pass prunes to subtrees that actually contain an unassigned element (it recursed the whole tree computing content signatures even when fully assigned — 36 ms of every 74 ms rebuild on a 15k-element document), and the per-scope part flush is skipped when nothing was assigned, with `Save(persistAnchorIds: true)` now flushing projected parts itself so neither save path depends on rebuild side effects. Net (NVCA, native, editor profile): per-op index rebuild 74 ms → 2 ms; `ReplaceText` ~130 → ~40 ms, `InsertFootnote` ~190 → ~40 ms, `InsertTable` ~245 → ~28 ms, `SetPageNumbering` ~90 → ~8 ms.
- **`ListItemRetriever.InitListItemInfo` is idempotent.** Re-initializing a partially annotated live document — any session that gained a paragraph since its first initialization — used to hit `SetParagraphLevel`'s double-set guard and throw `"should never set ilvl more than once"`; every caller but the projector's catch-all swallowed it, and the editor's Enter-split silently dropped its DOM update when the follow-up block render errored on list-bearing documents. Already-annotated paragraphs are now skipped (first-added annotations win on read, preserving exactly the values earlier passes computed).
- **Footnote/endnote authoring on `DocxSession` (#276).** New `InsertFootnote(anchorId, characterOffset, markdown)` / `InsertEndnote(...)`: create a note definition and cite it from a body paragraph at a character offset. The projection could already *read* notes (`fn`/`en` scopes, `EditSummary.FootnoteCount`) and the existing ops could already edit one (`ReplaceText` on the note's `p:fn`/`p:en` paragraph) and delete one (`DeleteBlock` on the `fn`/`en` definition, which also strips every reference to it) — creation was the only missing verb, so no separate edit/delete note op was added. On a document with no notes yet the op writes the whole scaffold Word writes: the `FootnotesPart`/`EndnotesPart` with the two reserved notes (`w:type="separator"` at id `-1`, `w:type="continuationSeparator"` at id `0`), the `w:footnotePr`/`w:endnotePr` settings declaration (inserted at its CT_Settings schema slot via the shared `EnsureSettingsChildInOrder` — the settings part is never wholesale reordered), and find-or-create `FootnoteText`/`EndnoteText` + superscript `FootnoteReference`/`EndnoteReference` styles so the citation isn't rendered as full-size body text. Markup is Word-faithful on both sides, including the `w:footnoteRef`/`w:endnoteRef` auto-number mark on the note's first paragraph. Note ids are kept **ascending in reference order** — the invariant every Word-authored document holds, and one renderers depend on: LibreOffice numbers body markers by citation position but pairs them against the id-sorted definition list, so a first-cited note holding the highest id silently renders the *wrong note text*. A citation that follows every existing one takes `max(id)+1`, scanning references as well as definitions so a document with non-contiguous ids can't alias an existing note; a citation that lands earlier takes the smallest id cited after it, shifting that note and every higher one up by one across all parts. Word-reserved notes (`separator`/`continuationSeparator`/`continuationNotice`) are never renumbered. The offset resolves through the same `SplitRunsAtOffset`/`SplitInlineContainersAtOffset` pair `SplitParagraph` and `ApplyFormat` use, so a citation lands cleanly mid-run and inside a hyperlink. `EditResult.Created` returns the note definition anchor plus its paragraph anchors; `Modified` returns the citing paragraph. Body paragraphs only — Word does not allow a note reference in a header/footer story or inside another note, so a non-body anchor is `AnchorWrongKind` rather than a document Word offers to repair. Rippled through `DocxSessionOps` → WASM `DocxSessionBridge` + npm `DocxSession.insertFootnote`/`insertEndnote` → stdio `Dispatcher` (`insert_footnote`/`insert_endnote`) + `docx-scalpel` `insert_footnote`/`insert_endnote`. Coverage: `Docxodus.Tests/DocxSessionNoteAuthoringTests.cs` (DS320–DS337, including OOXML schema validation), `npm/tests/docx-session-notes.spec.ts`, `python/tests/test_notes.py`.

- **The browser `DocxEditor` renders and edits footnotes/endnotes (#276).** An editor that can author a note has to be able to show it — previously an authored footnote was correct in the file and invisible on screen, as were a document's own existing notes. Three parts: (1) the editor's render profile emits the notes section and the numbered citation markers — `DocxSessionOps.RenderHtml` and the first-paint `completeArgs` both turn `RenderFootnotesAndEndnotes` on, and must stay in step since the remount output has to match the first paint byte-for-byte; notes are document *content*, so unlike the header/footer bands this is not opt-in. (2) `HtmlConversionOps.AssignAnchorUnids` stamps the deterministic anchor Unids on the footnotes/endnotes parts as well as the main part, so note paragraphs carry `data-anchor` and the editor wires them as ordinary editable blocks — the whole ribbon works inside a note with no new command code — and `FindByUnid` searches those parts so a single note can be re-rendered after an edit. Header/footer parts stay unstamped by design (paginated output clones one header node per page, so its anchor would not be unique; each note renders exactly once). (3) The citation marker and note backref are converter-generated chrome that the session's run text does not contain, so the editor excludes them from its content-offset space through one shared `GENERATED_CHROME_SELECTOR` and marks them non-editable — without which offsets drift, or the rendered display number gets committed as literal text and destroys the citation run. Both modes show notes: continuous renders the converter's `section.footnotes` at the end of the body flow, while paginated activates `pagination.ts`'s existing footnote engine — notes land at the bottom of the page that cites them, above a separator rule, with continuation onto the next page when one doesn't fit. (Endnotes render as a section after the page stack rather than on their own final page — a layout imperfection; they are visible and editable.) Coverage: `Docxodus.Tests/DocxSessionNoteRenderTests.cs` (DS340–DS345), `npm/tests/editor-footnotes.spec.ts`; three existing specs that sampled "the first/last block" were narrowed to body blocks, since note content now also lives in the body flow.

- **Header/footer authoring reaches `docxodus-mcp` (issue #316).** The engine, WASM bridge,
  stdio host, and `docx-scalpel` could already create running stories, but an MCP agent could
  neither reach those ops nor search the resulting story text. `docxodus_create` now exposes
  `set_header_text`, `set_footer_text`, and `ensure_header_footer_visible` as thin
  `DocxSessionOps` routes (`bodyAnchorId`, `kind: default|first|even`, and `markdown` for the two
  setters). The setters return the real `p:hdr*`/`p:ftr*` paragraph anchor, which composes with
  the existing `insert_page_number_field` action and reads back through anchor-scoped markdown,
  text, or HTML. `docxodus_search` gains an optional
  `scope: body|headers|footers|header_footer|all` for text/regex searches; omission remains
  body-only, preserving existing result sets. Tool schemas and the agent-server contract document
  the flow. Coverage: MCP137–MCP139.
- **Page-number formatting on `DocxSession` (#277).** Follow-up to #236/#274, which emitted only a plain Arabic, continuous-numbering `PAGE`/`NUMPAGES` field. Two independent layers land, deliberately kept apart. **The section:** new `SetPageNumbering(bodyAnchor, PageNumberingOp { Start, Format })` writes `w:pgNumType` — exactly what Word's *Format Page Numbers…* dialog writes — so a section can restart at a number (`w:start`) and number its pages in a chosen format (`w:fmt`, e.g. `lowerRoman` front matter). Both fields are tri-state: null leaves that *attribute* alone, so the start is settable without disturbing the format and vice versa. New `ClearPageNumbering(bodyAnchor)` removes the two attributes, preserving the chapter-numbering ones this surface never writes (`w:chapStyle`/`w:chapSep`) and dropping the element only once nothing is left on it — a "clear" that discarded the rest of the element would be a silent data loss. Both are addressed by any body block, resolve the governing `w:sectPr` exactly as `GetSectionInfo` does (synthesizing a trailing one if the body has none), and are no-ops — *without consuming undo history* — when the document already says what was asked, because a format dropdown fires on every selection and would otherwise evict the user's real edits from the bounded snapshot ring. **The field:** `InsertPageNumberField` takes an optional `NumberFormat` writing the field's own `\*` general-formatting switch (`PAGE \* roman`). Omitting it emits the plain field byte-for-byte as before; a switch *overrides* the section for that one field and keeps overriding it if the section later changes, so it is the escape hatch rather than the default route (the editor band deliberately inserts plain fields). The cached field result is seeded with page 1 rendered in the requested format (`i`, `A`, `1`) instead of a hardcoded `"1"`, so a renderer that does not recompute fields agrees with the switch. **Read-back:** `SectionInfo.PageNumberStart`/`PageNumberFormat`, *omitted* rather than defaulted when the attribute is absent — "continues the previous section in the default format" is a different claim from "starts at 1 in decimal", and a UI that cannot tell them apart writes attributes the document never had. The existing public `NumberFormat` is reused rather than duplicated (it is already this library's name for `ST_NumberFormat`, the type of both `w:numFmt` and `w:pgNumType/@w:fmt`); `NumberFormat.Bullet` and a negative start are rejected with the new `EditErrorCode.InvalidPageNumbering` rather than silently degrading to decimal. Rippled through `DocxSessionOps` → WASM `DocxSessionBridge` + npm `DocxSession.setPageNumbering`/`clearPageNumbering` → stdio `Dispatcher` (`set_page_numbering`/`clear_page_numbering`) + `docx-scalpel` `set_page_numbering`/`clear_page_numbering`. The browser editor's header/footer band gains **format** and **start-at** controls (`DocxEditor.setPageNumbering`/`clearPageNumbering`/`pageNumbering`), shown on both bands with the same values because they describe the section rather than either story. Coverage: `DocxSessionTests` DS317–DS328 (including OOXML schema validation), `npm/tests/page-numbering.spec.ts`, `python/tests/test_page_numbering.py`.

- **Paginated preview renders real per-page numbers.** A header/footer is authored once and cloned onto every page, so a page-number field's single cached result read the *same number on every page* — the paginated view showed "1" throughout a 48-page document. The converter now marks `PAGE`/`NUMPAGES` complex-field results with `data-field` (paginated mode only, so every other mode's HTML is untouched) and stamps the section's `w:pgNumType` on its wrapper as `data-page-num-start`/`data-page-num-fmt`; `pagination.ts` substitutes each page's real number after layout, formatted by the section's format or, when the field carries its own `\*` switch, by that. `NUMPAGES` gets the true total, which is only knowable once the last page exists. Header/footer parts are now annotated with field info during the paginated render (only the main part was), without which a running head's `PAGE` field is five unrelated runs that could never be identified — a side effect worth knowing is that a HYPERLINK field in a running head now renders as a real `<a href>` in paginated output, where it previously rendered as plain runs. New `npm/src/page-number-format.ts` owns the browser-side number rendering (both `ST_NumberFormat` tokens and `\*` switch arguments, whose case is load-bearing — `roman` is `i, ii, iii` and `ROMAN` is `I, II, III`). Coverage: `npm/tests/page-numbering.spec.ts`.

- **`DocxEditor.insertFootnote()` / `insertEndnote()` (#276).** `DocxSession.InsertFootnote` shipped through .NET, WASM and npm, and the editor learned to *render and edit* notes, but no editor command ever created one — a note could be read and rewritten in the browser and not authored there. The commands cite a new note at the caret in the active body block: the caret's DOM offset is captured *before* `syncBlock` (which re-renders the block and would drop the live selection) and mapped through the same `trimmedSplitOffset` that `SplitParagraph` uses, so the citation lands where the caret actually is rather than at a stale offset. Body blocks only, matching the session's own rule — a header/footer band block or a paragraph inside an existing note is rejected client-side rather than round-tripping to an `AnchorWrongKind`. A new note renumbers every citation after it and can add a whole part, so the commands remount rather than swapping one block. The bridge members are declared optional, so an older WASM bundle degrades to a no-op instead of a `TypeError`. Verified against a 95-footnote filing template: a mid-document insert renumbered all citations 1…95 with note ids still **ascending in reference order** (the invariant LibreOffice depends on), and the saved file renders in LibreOffice with the new note in place.

### Changed
- **The browser editor demo is a tabbed ribbon (`npm/examples/editor.html`).** The demo had grown to ~25 controls in a single unlabeled strip that wrapped onto two rows, with New/Open/Save styled identically to the formatting buttons and a font-size box whose placeholder truncated to "Siz". It is now organized the way a document editor is: an always-visible document strip (New / Open / Save / Undo / Redo — never behind a tab, because they are used constantly), a **Home** / **Insert** / **Layout** tab set with labeled groups separated by hairlines, and a **contextual Table tab** that appears only while the caret is inside a table. That last one replaces the floating table toolbar, whose absolute positioning had to be corrected twice before (it covered the first row, then the content below the table); a docked tab cannot overlap the cell being edited, so the whole class of bug is gone by construction. Alignment and indent controls are inline SVG drawn from the text lines they act on, rather than arrow glyphs that collided visually with undo/redo. The header/footer band chrome is restyled onto the same tokens so the page reads as one instrument instead of two. New in the ribbon: an **anchor rail** under the tabs reporting live engine state — the focused block's `kind:scope:unid` anchor, the block count, the session handle, and the last command with its real duration. It is the addressing spine made visible, and it makes the cost of an operation legible instead of absorbed (a full-document remount on a 346-block document reads `6.20 s` where a single-block edit reads `< 100 ms`). Every command routes through one `run()` wrapper that records it. Tab activation is exposed as `window.__selectTab(name)` so specs can reach a control without depending on pointer geometry.

- **`npm test` now type-checks the Playwright specs.** `tsconfig.json` compiles `src/**` to `dist/` and Playwright strips types without checking them, so nothing verified the specs — a spec could bind a value to a public type that no longer existed and still pass. That is exactly how `EditErrorCode` lost a union member while a runtime assertion on the string literal kept passing. New `tsconfig.tests.json` + `npm run typecheck`, wired into `pretest`. Four legacy specs with pre-existing errors (all the same `new Uint8Array(page.evaluate(...))` union-inference shape) are excluded by name with a note, rather than weakening `strict` for everything.

### Fixed
- **Tracked-changes renders now renumber lists the way Word does (the NVCA marquee-redline numbering bug).** A redline of a numbered list rendered one continuous number sequence across every paragraph — deleted and moved-away paragraphs consumed numbers — so every number after a deletion disagreed with the final document (the README's NVCA voting-agreement screenshot showed `(h)`–`(p)` where Word shows `(g)`–`(o)`). Word ties numbering to the paragraph mark and renumbers as if all changes were accepted: a paragraph whose pilcrow is marked deleted (`w:pPr/w:rPr/w:del`, as both `WmlComparer` and `DocxDiff` emit for a fully deleted paragraph, or `w:moveFrom` for a move source) still *displays* the value the counter holds at its position but does not advance it, so the next live paragraph shows the same number — the duplicate struck/live pairs Word renders in All Markup view. `ListItemRetriever` now implements exactly that: such a paragraph gets its `LevelNumbers` annotation (it still renders a struck number) but contributes nothing to the forward-carried numbering state (`previous`, start-override consumption, continuation tracking). Unconditional rather than gated on a setting, because the default render path accepts revisions before numbering ever runs — only tracked-changes-bearing documents can hit it, and as-if-accepted is the correct reading of the format (Word and LibreOffice agree). A second, compounding defect in `FormattingAssembler`: the "previous paragraph's pilcrow is inserted" heuristic — meant for the split-paragraph markup where pressing Enter inside existing content leaves the ins-marked mark on the first half, making the *follower* the new list item — also fired when the predecessor was a **wholly inserted** paragraph (own ins pilcrow + all content inserted, the comparer's shape for a new or moved-in paragraph), wrapping the *unchanged* next paragraph's number in `w:ins` with the reviewer's author attribution. It now checks that the predecessor carries pre-existing content before treating the follower's number as inserted. `docs/images/redline.png` regenerated from the same NVCA voting-agreement edit round, now with the inserted definition and the move destination as real numbered `Heading3` list members. Word's behavior documented in `docs/ooxml_corner_cases.md`. Coverage: `Docxodus.Tests/TrackedChangesNumberingTests.cs` (TCN001–TCN005: deletion, move source, wholly-inserted predecessor, genuine split, and an end-to-end `DocxDiff` redline).
- **Header/footer text before a tab is no longer dropped from the render.** `ConvertParagraph` splits a paragraph at its first tab run and filtered the preceding runs to those carrying a `PtOpenXml:TabWidth` annotation — but that annotation comes from `CalculateSpanWidthForTabs`, which walks the MAIN document part only. A header/footer run therefore never had it, so every run before a tab vanished from the output: excluded from the preceding-tab path, and not part of the succeeding-tab range either. The common Word running foot `Last Updated October 2025 [tab] PAGE` rendered as just the page number. The filter conflated "which runs contribute to the computed tab width?" with "which runs get rendered?" — only the former should be filtered, and an unannotated run now contributes zero width while keeping its text. Coverage: `Docxodus.Tests/PaginatedHeaderFooterContentTests.cs`.
- **Header/footer content cloned into a paginated page box is no longer editable.** A running story is authored once and cloned onto every page, so every clone carries the same `data-anchor` — 42 page boxes claiming one footer paragraph on a real filing template. Committing any one of them wrote back through that single shared anchor. The per-page number substitution above turned that latent duplicate into a corruption vector, because each clone now shows a *different* number: committing one would write that page's number into the shared story as literal text and destroy the `PAGE` field. Page-box story content is presentation, and the docked editing bands are the addressable affordance — they exist precisely because a cloned node cannot be uniquely addressed — so the clones now have their block addressing stripped. Coverage: `npm/tests/page-numbering.spec.ts`.
- **`DocxEditor.save()` no longer writes the projector's anchor bookkeeping into the user's file.** The editor opened its session with `persistAnchorIds: true`, which suppresses the `PtOpenXml:Unid` strip on *every* save — so a document saved from the editor without a single edit came back at roughly 6x its original size (147KB → 928KB on a real 234-paragraph filing template). It went unnoticed because the attributes live in a custom namespace Word and LibreOffice both ignore, so the output renders identically; only the byte count betrayed it. The setting was there for a real reason — a session render serializes the document and re-renders it, and a Unid is content-hashed, so stripping it made the converter re-derive a fresh id for any block edited since the session assigned one, leaving that block's `data-anchor` unresolvable and the block silently un-editable. The requirement belongs to the **render**, not to the session: `HtmlConversionOps.ConvertToHtml(DocxSession, …)` now serializes with anchor ids unconditionally (those bytes are an internal hop that is rendered and discarded), the editor opens its session with the default settings, and the one remaining JS-side re-render path asks for them explicitly via the new `SaveWithAnchorIds` bridge export. New public `DocxSession.Save(bool persistAnchorIds)` overload makes the choice per call; `Save()` is unchanged. Editor saves are now byte-for-byte what the .NET API produces (220,513 vs 216,269 on that template — the remaining delta is zip framing). Coverage: `npm/tests/editor-save-clean.spec.ts` (zero-edit save carries no `Unid=`, and a list edit still remounts with blocks wired — the invariant a size assertion cannot see).
- **Paginated header/footer rendering honors `w:evenAndOddHeaders`.** The paginated registry gated the *first-page* stories on the section's `w:titlePg` but emitted the *even-page* stories whenever a `w:type="even"` reference existed — and a reference of that type is inert on its own. Word removes only the document-global `w:evenAndOddHeaders` flag when "Different odd & even pages" is switched back off, leaving the part and its reference in the package, so a document in that state rendered a running foot it does not have. Found by smoke-testing the NVCA model certificate of incorporation, a real filing template carrying three `w:type="even"` footer references with the flag absent: its leftover even footer reads "DRAFT" and has no `PAGE` field, so the paginated view showed "DRAFT" — and therefore no page number — on every even page, where both Word and LibreOffice show the Default footer and its roman-numeral page number. The even stories now follow their own flag in the same place the first stories follow theirs. Coverage: `Docxodus.Tests/PaginatedHeaderFooterGatingTests.cs`; documented in `docs/ooxml_corner_cases.md`.
- **Paginated footnote layout (#276).** Enabling footnote rendering in the editor exercised `pagination.ts`'s footnote engine against a dense real document for the first time and surfaced four defects, all fixed: (1) **notes were silently dropped** — an unfitted note was carried in a *single* continuation slot that the code assigned once per note in a page's citation list, so when two notes on one page both failed to fit the second overwrote the first and it rendered nowhere (four notes vanished from a 94-footnote document); unfitted notes now queue and are merged, not overwritten, into the next page's note list. (2) **The stylesheet's child combinators were XML-escaped.** The generated CSS is the *value* of an `h:style` element, so serializing the XHTML turned `.footnote-content > p:first-of-type` into `… &gt; p…` — not a valid selector, so browsers dropped the rule and every note rendered its number alone on one line with the text below it. The affected rules now use descendant selectors, and `GeneratedCssEscapingTests` fails if any generated CSS reaches the browser XML-escaped. (3) **A note that could not be split was re-wrapped inside itself**, nesting a complete `.footnote-item` (number span and all) inside another item's content and breaking the same line. (4) **Body text could be drawn underneath the note block**, which is bottom-anchored and grows upward while the content area spanned the full text height — ~134pt of superimposed, illegible glyphs on one page. The content area is now shrunk to the space the notes actually leave, making the collision impossible by construction (worst case is a clean clip, not corruption), and the note-height measurement runs in the same `.page-footnotes` styling context the notes render in, so the reserve matches what is drawn. Net effect on the reference document: 56 pages → **48, exactly matching the LibreOffice reference**; zero overlapping pages; zero dropped notes; worst-case body clipping 144px → 16px. Coverage: `npm/tests/pagination-footnote-layout.spec.ts`.
- **Word-reserved `continuationNotice` notes are no longer projected as user notes (#276).** `IsBoilerplateNote` filtered only `separator`/`continuationSeparator`, so a `continuationNotice` — which real documents carry — reached the anchor index and the markdown projection as if it were user content, and once the editor began rendering notes it appeared as a stray empty footnote with no citation. Any *typed* note is now treated as reserved: ECMA-376 §17.11.17 defines the type as `normal` | `separator` | `continuationSeparator` | `continuationNotice`, and only `normal` (which Word omits rather than writes) is user content. `DocxSession.DeleteBlock`'s reserved-note guard now shares that one predicate so the two can't drift.
- **A `DocxSession`-saved `.docx` no longer extracts as mode `000` (unreadable) under a Unix `unzip` (#302).** `DocxSession.Save()` keeps the underlying `WordprocessingDocument`/`System.IO.Packaging.Package` open across the session and flushes it in place (`_doc.Save()`) rather than disposing it — and on a non-Windows host, that flush path stamps the zip central directory's "version made by" host byte as Unix while leaving `ExternalAttributes` (the Unix permission bits) at `0` for a genuine Word-authored input, whose entries already carry `0` from the original DOS-hosted archive. `unzip` takes a Unix-hosted, zero-permission entry literally and extracts it unreadable even to its owner; Word, LibreOffice and Python's `zipfile` don't consult these bits at all, so the corruption was invisible outside a pipeline step that shells out to `unzip`. The dispose-based `OpenXmlMemoryStreamDocument.GetModified*Document()` path used by `DocumentBuilder`/`HtmlToWmlConverter`/etc. does not reproduce this on the currently-targeted SDK, but gets the same mechanism-agnostic fix for defense in depth. New internal `ZipUnixPermissionFixer`: after a save, any zip entry whose `ExternalAttributes` is still `0` is assigned a sane default (`0644` files / `0755` directories); a no-op on Windows and a no-op wherever attributes are already set. `DocxSession.Save()` applies it to a copy of the output bytes only, never to the session's own live backing stream, so repeated edit/save cycles on one session can't be corrupted by the fixup itself. Coverage: `Docxodus.Tests/ZipUnixPermissionFixerTests.cs` (DS346–DS349), reproduced against a real Word-authored fixture (`TestFiles/Blank-wml.docx`) whose entries start at the exact zero-attribute baseline the bug needs. **Follow-up:** `WmlComparer` — the default/blessed comparison engine — turned out to manage its own open/dispose/save pipeline independent of `OpenXmlMemoryStreamDocument`, so it had no code-level defense at all despite being the most-used save path in the library; an adversarial review of the original fix caught the gap. `Compare()`'s and `Consolidate()`'s single final output points now route through the same `ZipUnixPermissionFixer`. Doesn't reproduce the bug on the currently-targeted SDK either, but is no longer relying on that happening to stay true. Coverage: DS350–DS351.

### Changed
- **`EditErrorCode.FootnoteRefNotSupported` narrowed rather than retired (#276).** The code stays (it is public surface clients switch on) but now means only "a `[^label]` reference in a markdown *payload* can't be resolved to a note this payload doesn't define"; its message names `InsertFootnote`/`InsertEndnote` as the op to use instead of the old "planned for v2".

### Fixed
- **`docxodus-mcp`'s `set_paragraph_format` can now add a paragraph border, not just clear one; `set_paragraph_format` was entirely missing from `docx-scalpel` (issue #301).** `DocxSession.SetParagraphFormat` already wrote `w:pBdr/w:top`/`w:bottom` via `ParagraphFormatOp.TopBorder`/`BottomBorder` — it's what `InsertHorizontalRule` is built on — but neither client could reach it. In `tools/mcp-server`, `docxodus_format`'s `paragraphFormat` JSON Schema advertised only `alignment`/`indentDelta`/`pageBreakBefore`/`clearBorders`, so an agent had no argument shape to add a border to an *existing* paragraph (e.g. a rule under a document title) even though `Dispatcher.ParseParagraphFormatOp` already forwarded the raw JSON to the shared `DocxSessionJson.ParseParagraphFormatOp`, which already parsed `topBorder`/`bottomBorder`. `ToolCatalog.cs` now declares both border fields (`{style, size, color, space}`, mirroring `ParagraphBorderEdge`); no dispatcher change was needed. In `tools/python-host` + `docx-scalpel`, the gap was total: `Dispatcher.cs` had no `set_paragraph_format` case at all, so alignment, indent, page-break, and borders were all unreachable from Python, not just the border fields. New `Dispatcher.cs` case, `docx_scalpel.ParagraphAlignment` enum, `ParagraphBorderEdge`/`ParagraphFormatOp` dataclasses, and `DocxSession.set_paragraph_format(anchor_id, op)`. Coverage: `Docxodus.Tests/McpServerDispatcherTests.cs` (`MCP042`), `python/tests/test_paragraph_format.py`.

## [8.0.0] - 2026-07-29

### Added
- **Visual header/footer editing region in the browser `DocxEditor` (#275).** Opt in with `headerFooter: true` and the editor docks a **Header** band above the body flow and a **Footer** band below it. Header/footer stories live in their own OOXML parts outside the body, so the bands are composed per story paragraph via the session-attached `RenderBlockHtml` (which already resolves `hdr`/`ftr` anchors) rather than being part of the body render — the same renderer used for the post-edit incremental swap, so there is no fidelity drift between first paint and repaint, and there is exactly one addressable DOM node per story paragraph in both continuous and paginated mode. Story paragraphs are wired by the same `wireBlock` the body uses, so they are ordinary editable blocks: the whole existing ribbon (bold/italic, alignment, font family and size, paragraph style) works inside a band with no new command code, and only the edited band repaints — the body is never remounted. Band chrome adds a **kind selector** (`default` / `first page` / `even pages`; selecting a kind with no part seeds an empty story so there is always something to type into) and a **page-number control** (`currentPage` → `PAGE`, `totalPages` → `NUMPAGES`). Choosing **even** surfaces the caveat that `w:evenAndOddHeaders` is document-global and governs footers too — even pages stop inheriting the Default stories — with a one-click "also create an even footer". When a body block is focused the bands follow *its* section, so a cover-page-plus-body document shows the stories that actually apply. New `DocxEditorOptions.headerFooter` (default **false**, so the editor's DOM is unchanged for existing consumers), new `DocxEditor.setHeaderFooterKind`/`headerFooterKind`/`insertPageNumber`, new module `npm/src/editor-headerfooter.ts`. Coverage: `npm/tests/editor-headerfooter.spec.ts`.
- **`DocxSession.EnsureHeaderFooterVisible(anchorId, HeaderFooterKind)` — make a section's first/even stories actually render.** Sets `w:titlePg` for `First` and the document-global `w:evenAndOddHeaders` for `Even`; `Default` needs no flag and is a successful no-op. Idempotent. `SetHeaderText`/`SetFooterText` already set these flags *while writing content*, which covers authoring a story from scratch — but not a document that already carries a first/even reference with the flag absent, which is exactly what Word leaves behind when "Different first page" / "Different odd & even pages" is switched back off. Editing such a pre-existing story through the anchor-addressed text ops otherwise produces a file whose header content is present but never rendered (confirmed against `TestFiles/HC031-Complicated-Document.docx`, whose six stories all carry references with neither flag). The flags belong to the *section*, not to a content write, so this is its own operation. Wired through `DocxSessionOps` → WASM `DocxSessionBridge.EnsureHeaderFooterVisible` → npm `DocxSession.ensureHeaderFooterVisible` and the stdio host (`ensure_header_footer_visible`) → `docx-scalpel` `ensure_header_footer_visible`. Coverage: `DocxSessionTests.DS268`.
- **`SectionInfo.HeaderRefs`/`FooterRefs` — each `w:headerReference`/`w:footerReference` with its `w:type`.** `HeaderPartUris`/`FooterPartUris` report *which* parts a section references but not which story kind each supplies, and the projection's `hdr{N}`/`ftr{N}` numbering is by part-collection order, which carries no kind information — so a client could not tell an existing document's Default story from its First or Even one. The new lists pair each reference's `HeaderFooterKind` with its part URI (an absent `w:type` reads as `Default`, per ECMA-376 §17.6.10). They report the stories that **effectively apply**: a section declaring no reference of a kind continues the previous section's (§17.6.17), and such entries come back with `Inherited = true`. That matters because a multi-section document typically defines its headers once in the first section — `HC031-Complicated-Document.docx` has four sections and only the first declares any — so own-references-only would report "no header" for most of the document and a caller acting on that would mint a redundant part and break the inheritance. `HeaderPartUris`/`FooterPartUris` keep their original meaning (own references only). Wired through `DocxSessionJson` → npm `SectionInfo.headerRefs`/`footerRefs` (`HeaderFooterRef`) and `docx-scalpel` `SectionInfo.header_refs`/`footer_refs`. Coverage: `DocxSessionTests.DS265`/`DS266`.
- **`DocxDiffSettings.CrossParagraphTokenDiff` (default true) — the cross-paragraph word+pilcrow token stream decoded from Word's compare output.** A run of adjacent word-matched paragraph pairs diffs as ONE token stream: retained words may cross pilcrow boundaries, paragraph marks are ¶INS/¶DEL stream tokens, and the output paragraph count follows the token-level interleave. The stream also spans a story-final one-sided tail, and the story-final pair retains only its common unit prefix plus adjacent-bigram recoveries. Markup-only (`GetRevisions`/`GetEditScriptJson`/`Consolidate` are unaffected); accept ≡ right / reject ≡ left holds pilcrow-exactly. Wired through the WASM bridge, npm, the stdio host, and `docx-scalpel`.
- **Gap-region streams — the same stream generalized to replace regions without a full pair run.** A zero-pair story-final region streams on ≥2 shared word units or on a single count-equal boundary construct (at most one per region, the matching phase ends there; a lone same-position function-word match never forms one); leading one-sided members join a following pair's stream; interior regions ship only on a construct (or an in-slot match plus a deleted member); one-sided section-break entries are transparent story-end metadata; fragments and whole-document rewrites stay with the replace-gap grammar. Includes the story-tail MERGE-SLOT law: a base surplus merges at the FIRST pair whose next-side tail residue shares an unmatched content word with the following base paragraph, displaced pairs re-slotting forward.
- **`PaginationOptions.fragmentParagraphs` — ordinary paragraphs may now split across a page boundary in the read-only pagination viewer.** Previously `PaginationEngine` always kept a paragraph whole, so a paragraph taller than the remaining page space fell entirely onto the next page, leaving a large blank gap at the bottom of the prior page — a visible divergence from how Word/PDF viewers flow text. `PaginationEngine.tryFragmentParagraph` now, for a simple text-only `<p>` (no `keepWithNext`/`keepLines`/page-break flags, no non-inline or pagination-sensitive descendants — objects, breaks, lists, tables, footnote markers, nested anchors, etc. all still fall back to the existing whole-block behavior), binary-searches DOM Range endpoints at whitespace/run boundaries to find the largest head fragment that fits the current page, clones it as the page-ending fragment (keeping the paragraph's `id`/`data-anchor`), and continues the remainder as a normal block on the next page (identity-free, so it can fragment again). Off by default on the low-level `PaginationEngine`/`usePagination`/`PaginatedDocument` (npm `pagination.ts`/`react.ts`) API; the convenience `paginateHtml()` read-only-viewer entry point turns it on by default. The in-browser `DocxEditor`'s paginated mount explicitly opts out (`fragmentParagraphs: false`) since a continuation fragment has no addressable anchor of its own and the editor's block-editing model requires one. New npm/TS option `fragmentParagraphs` on `PaginationOptions`/`PaginatedDocumentProps`/`usePagination`. Coverage: `npm/tests/pagination-paragraph-fragments.spec.ts`.
- **`DocxDiffSettings.NormalizeRevisionAuthors` — collapse every tracked-revision author in the output to a single author (default false).** When `PreserveInputRevisions` is on, the inputs' own tracked changes ride through under their ORIGINAL author. Renderers that color tracked changes BY AUTHOR (LibreOffice) then show that preserved content in a SECOND color, while Word's compare output is single-author / one color — so a document whose source carried foreign-authored suggestions ("Online User" Google-Docs suggestions, etc.) renders in two colors versus Word's one. With this flag on, a byte→byte post-pass stamps `settings.AuthorForRevisions` onto the `w:author` of every tracked-revision element across all story parts (document, headers/footers, footnotes/endnotes, comments-part revisions — comment authors themselves are untouched), collapsing the output to one revision author. This normalizes the RENDER, not revision semantics: the accept ≡ right / reject ≡ left contract is unaffected (author is presentation metadata), and it does not touch Consolidate/N-way output (per-reviewer authors are intended there). Wired through the public `DocxDiffSettings`, the internal `IrDiffSettings`, and the stdio host (`normalizeRevisionAuthors`) + `docx-scalpel` (`normalize_revision_authors`). Coverage: `DocxDiffAuthorNormalizationTests`; documented in `docs/ooxml_corner_cases.md`.
- **`DocxDiffSettings.PreserveInputRevisions` — Word-parity preservation of the inputs' own tracked changes (default false).** Word's Compare PRESERVES tracked revisions already present in the input documents: the original author/date markup rides through into the compare output verbatim (verified against Word's compare output — an input carrying revisions by another author keeps them alongside the fresh compare revisions), while the text diff is computed over the accepted view. Our existing `PreAcceptInputRevisions` flattens them instead. With the new flag ON: no byte-level pre-accept runs (the LEFT package's carried-over parts — headers/footers, unchanged notes, styles, comments — keep their markup), and the markup renderer reaches back from its accepted working copy to the ORIGINAL right-body elements via an in-order alignment walk that models the document-level accept's paragraph-mark merges (a fully-deleted paragraph rides along in its group and vanishes again on accept). Content-EQUAL blocks emit the original element(s) verbatim; whole-block INSERTED content wraps only plain runs in this diff's `w:ins`, leaving foreign `w:ins`/`w:del` children, foreign paragraph-mark markers, and foreign row markers as-is — no same-kind wrapper ever nests. Preserved wrappers get fresh `w:id`s (no duplicate-id validator errors) and Word-extension attrs (`w16du:dateUtc`) are dropped. Note-scope preservation rides the same hooks: footnote/endnote definitions pair by id and their equal/inserted blocks preserve too. Char-precise paths (modify/format-only/split/merge/move, changed header/footer stories) still render over the accepted view — foreign markup there is flattened (v1 scope), as is the LEFT side's in deleted blocks. Round trip is one-sided exactly like Word: `accept(output) ≡ accept(right)` at text level; `reject(output) ≠ left` where foreign markup exists (rejecting a foreign deletion restores its text) — documented on the setting. When both flags are set, Preserve wins (pre-accept skipped). The `DocxCompare` engine-selector path (CLI `--engine=docxdiff`, WASM/npm selector) now opts in for Word parity. Foreign revisions ride through into the output with zero validator-error delta and an intact accept round trip. Coverage: `DocxDiffPreserveInputRevisionsTests` (equal-block ins/del preservation, modified-block schema cleanliness, insert-block preservation without nesting, fully-deleted-paragraph ride-along + fresh ids, clean-input byte-identity guard, precedence).

### Changed
- **DocxDiff: the WC-1450 parity case is adjudicated once, in the parity scoreboard, and pinned on engine truth (issue #289).** `IrSplitMergeTests.WC1450_compat_revisions_match_oracle_count` was skipped on the belief that WC023's fully-duplicated prose (the same sample text in BOTH the leading body paragraphs AND the table cells, so no block content is uniquely anchorable) collapsed `IrBlockAligner` into a whole-body delete + re-insert. Investigation showed it does not: the body aligns to 5 Unchanged plus the table as ONE `Modified` pair, and `IrTableDiffer` resolves the table to exactly the authored edit — `DeleteRow` (the removed row), `ModifyRow` (the `"Second "` cell edit) and two `EqualRow`s. The 2 revisions are that edit at the engine's row grain: a whole deleted row is ONE revision, which is what the WmlComparer oracle itself does for every other whole-row case in the corpus (WC-1140/1150/1460/1660/1670/1750/1760 pin it). Its 7 revisions here are not a finer reading of the same edit — it mis-aligns this fixture, reporting content as INSERTED that is unchanged on both sides, plus two null-text revisions. `IrParityScoreboardTests` already adjudicates WC-1450 as a documented coarser-grain deviation and ratchets the floor for it, so the second hard-count copy in `IrSplitMergeTests` was re-litigating a settled decision against a degenerate oracle. It is replaced by two un-skipped tests that pin what the engine actually says — the deleted row + the cell edit, and the block/row alignment shape that would regress if duplicate-content alignment ever *did* collapse. The stale "179/179 genuine passes, catalog EMPTY" claim in `docs/architecture/ir_diff_engine.md` is corrected to the scoreboard's real state (177 genuine + 2 documented deviations).
- **DocxDiff is now the default primary redline path.** Omitted engine selection in the CLI, WASM HTML bridge, npm API/worker, and browser test harness routes to DocxDiff. Explicit `ComparisonEngine.WmlComparer` / `--engine=wmlcomparer` remains available and retains wire value `0`.

### Fixed
- **`DocxSession`: an `EditResult`'s anchors now stay in the part the edit touched.** Unids are content-addressed, so identical content in *different* package parts yields the same unid — a document with empty default/first/even header stories has one unid shared across several header parts, which is what Word writes. Every element→anchor reverse lookup resolved by bare unid and returned whichever part the projection indexed first, so an edit result named a *different* story; a caller that addressed the returned anchor next (as an editor does for its re-render and the following command) then wrote into the wrong part. Symptom: a page-number field aimed at the default footer landed in the even one. Fixed at a single owner — `AnchorForUnid(unid, preferPartUri)` plus `PartUriOf`/`AnchorForElement`, with all sixteen reverse lookups passing the part they touched. `HtmlConversionOps.RenderBlockHtml`'s session-attached path likewise resolves through the anchor index before falling back to the unid scan, so a single-block re-render can no longer *display* a different part's block. Coverage: `DocxSessionTests.DS267`, `HtmlConversionOpsTests.HCO080`.
- **`DocxEditor`: formatting a just-typed paragraph no longer silently does nothing.** A block rendered with edge whitespace — an empty header/footer story renders as a lone NBSP placeholder — produces a selection span longer than the text the commit stores (`serializeInlineMarkdown(...).trim()` drops it, since JS `trim()` treats U+00A0 as whitespace). "Select all, then Bold" therefore asked `ApplyFormat` for one character past the committed end and the op was rejected out of range. The demo's format buttons `preventDefault` on mousedown precisely so the selection survives, which is the path that computes the span *before* `syncBlock` commits — the ordinary case, not an edge one. `selectionSpanIn`/`blockSpanForSelection` now normalize through a new `trimmedSpan()`, the span analogue of the existing `trimmedSplitOffset()`.
- **DocxDiff: rejecting a redline now reproduces the original document's paragraph ORDER exactly, not just its content (issue #288).** `DocxDiff.Compare` guarantees `accept ≡ right` / `reject ≡ left`. Content preservation always held — the byte-level round-trip fuzzer never lost or duplicated a word in either direction — but in a narrow case the **reject** direction restored every paragraph in the WRONG ORDER. The trigger is a document containing **verbatim-duplicate paragraphs** plus a split/merge edit near one of them: duplicates never anchor (`IrBlockAligner.BuildUniqueIndex` keys on content that is unique on each side), so which occurrence pairs with which is settled by the in-gap refinement — and the split/merge containment scan, which runs *after* that refinement, could then form a group that STRADDLES the pairing it had already made. `EmitEntries` walks the right document, so the straddled pair emitted at its right position and the group at theirs, and reject rebuilt the two left blocks permuted. Every in-gap pass already enforces order-preservation against the pairings that existed when *it* ran; rather than teach each pass about the others, the invariant is now enforced once, globally: `IrBlockAligner.EnforceInPlaceOrderMonotonicity` keeps a maximum-**weight** strictly-increasing subsequence of the (left → right position) map over ALL left-owning units — 1:1 pairs *and* split/merge groups — and releases the rest to plain Deleted/Inserted (always reversible), weight = blocks kept paired so an ambiguous cut sacrifices the pairing holding the least content. It runs before cross-gap move detection, so a released pair that is genuinely a relocation comes straight back as a `Moved` — frequently a *better* reading than the in-place pairing it replaced (the fuzz repros all recover their split/merge groups AND gain a correct move). An already-monotone alignment (every corpus document) returns after one linear scan, untouched. Reproduced at seeds 184/760/1714 of 2000; clean at 2000 seeds after the fix. Coverage: `IrAlignmentAsserts.AssertLeftOrderReconstructible` (now asserted on **every** aligner test), `IrBlockAlignerTests.Duplicate_paragraph_pairing_never_crosses_a_merge_group` + `Fully_duplicated_body_still_emits_left_blocks_in_left_order`, and `DocxDiffFuzzRoundTripTests`, which now compares whole-body BLOCK SEQUENCES (its word-multiset checks are order-independent by design and could not see this) with the default sweep raised 50 → 250 seeds so the canonical repro is covered without the env knob.
- **DocxDiff: the replace-gap paragraph arrangement follows Word's full structural grammar** — insert-block-first emission with leading-ins hoists, fusion and live pilcrows only at a story's end, accept-side paragraph-mark coalescing treating an entirely-deleted table as transparent, and a virtual structural pair when the base story ends with a table.
- **DocxDiff: in-gap paragraph pairing follows Word's positional-first matcher** — same-slot pre-pairing with size-parity and competitor-evidence guards, story-tail surplus absorption, a shared content-word requirement for similarity pairs (any-overlap fallback for zero-content gaps), and top-down positional row pairing for wholly rewritten grown tables.
- **DocxDiff: split/merge segmentation follows the merged-stream expansion arrangement** — char-weighted anchors sliced at solid-match segments; ins-residue joins the preceding member, del-residue the following member, and trailing merge-insert residue the final member.
- **DocxDiff: a story-final word-matched pair whose only retention is a single short-word sliver flushes to one whole ins+del rewrite** (count, length, and fraction thresholds all required; interior pairs never flush).
- **DocxDiff: separator lexing and attachment match Word's** — ASCII `:` is always its own token (digit-interior included), whitespace/punctuation attach to retained anchors by the decoded separator laws, separators flanking a match ride with it across cell and slice boundaries, and lone matched punctuation is suppressed unless anchored.
- **DocxDiff: `w:pPrChange` presence follows the surviving-pilcrow and comparison-fold laws** — split/merge members with new pilcrows never carry one, the member owning the surviving pilcrow stamps history against its counterpart, default-equivalent `w:jc` and unresolvable `w:pStyle` fold out of the comparison, and shared pilcrows archive an empty `<w:pPr/>` when the base side had no properties.
- **DocxDiff: carried-through parts are normalized to what Word emits** — known-font declarations in an input `word/fontTable.xml` are rewritten to Word's stock metadata, and an imported right-side list instance gets its own fresh cloned `w:abstractNum` unless it is a surviving left list, so imported lists restart numbering instead of continuing a left list's counter.
- **DocxDiff: the intra-paragraph token diff now resolves anchor ties by matched CHARACTER length, matching Word's compare tie-break.** When two same-length common subsequences of content words exist, the previous Myers alignment kept whichever the greedy walk reached first (a token-count tie). Word instead keeps the subsequence covering the most characters — a distinctive long word (`"strikethrough"`, 13 chars) is retained as the anchor over an incidental short one (`"text"`, 4 chars), and a contiguous phrase is preferred over a scattered pair. `IrTokenDiffer.ContentAnchors` now selects the common subsequence maximizing total matched character length (`CharWeightedLcs`, an O(n·m) weighted-LCS with a deterministic prefer-left back-walk that degenerates to token count when all weights are 1), instead of the greedy Myers LCS. The accept ≡ right / reject ≡ left contract and the token-diff tiling invariants are preserved (verified by `IrDiffFuzzTests`). Coverage: `IrTokenDifferTests.Repeated_words_with_distinct_tail` (the tie-break regression pin).
- **DocxDiff: the output now carries a synthesized `word/fontTable.xml` + `word/webSettings.xml`, matching Word's compare output — the single biggest redline-render-fidelity fix.** Word's compare synthesizes both parts for every output document (the source documents rarely carry them). A fontTable declares each font's `panose1`/`charset`/`family`/`pitch` metrics, and that is exactly what LibreOffice consults to pick a substitute for a font it does not have installed (Aptos, Calibri, a raw CSS stack). When Word's redline carries a fontTable and ours does not, the two documents substitute the *same* absent font *differently*, so even a byte-identical body renders to different glyphs/metrics and a rendered redline diverges from Word's compare output. A new single-owner backfill (`WordCompareFontTableBackfill`, applied alongside the existing stock-theme and stock-docDefaults backfills) emits a fontTable listing Word's stock faces (Times New Roman/Aptos/Calibri/Aptos Display, with Word's exact metrics) plus every font the body or styles part actually references (unlisted fonts get a generic swiss/roman/modern descriptor keyed on the name; a raw CSS font stack such as `"Roboto, sans-serif"` is declared with Word's own `<w:altName>` = its primary family, which is what LibreOffice resolves for substitution), and an empty-bodied webSettings. Because the fontTable now carries the CSS stack's substitution metadata, the renderer no longer rewrites a shared stack to `Arial` (the old "reliable fallback" workaround, which *diverged* from Word — Word keeps the raw stack): the stack rides through verbatim and both sides substitute it identically. The `accept ≡ right` / `reject ≡ left` contract is untouched (these are font-metadata parts, not tracked-changes markup). Coverage: `WordCompareFontTableBackfill`, `DocxDiffCssFontStackCompatibilityTests`.
- **DocxDiff: the output now backfills the canonical `word/settings.xml` children Word's compare synthesizes — chiefly `compat/compatibilityMode`, the second parts-fidelity render fix.** Word's compare output synthesizes a canonical settings part for every document even when the source carries only an empty stub (verified against Word's compare output). The load-bearing element is `compat/compatibilityMode`: it selects LibreOffice's layout-engine emulation (Word 2007/2010/2013+), so when Word's redline carries one and ours does not, the two documents lay out under *different* engines and a rendered redline diverges from Word's compare output even on a byte-identical body — while `characterSpacingControl`/`themeFontLang`/`clrSchemeMapping` are inert against LibreOffice's defaults and are emitted purely for parity with what Word writes. A new single-owner backfill (`WordCompareSettingsBackfill`, applied alongside the stock-theme/docDefaults/fontTable backfills) ensures each canonical child exists, inserting it at its CT_Settings schema slot via a shared ordered-insert primitive (`WordprocessingMLUtil.EnsureSettingsChildInOrder`, the same slot-insert discipline as `EnsureEvenAndOddHeaders` — never a whole-part reorder). `compatibilityMode` follows Word's articulable rule (matches Word's compare output in the common cases): keep the ORIGINAL (left) document's value when present, otherwise the revised (right) document's, otherwise `12` (Word's default for an unmarked .docx) — so a genuine mode-15 document is never downgraded. The `accept ≡ right` / `reject ≡ left` contract is untouched (these are document-settings, not tracked-changes markup; verified by the round-trip fuzzer). Coverage: `DocxDiffSettingsBackfillTests`.
- **DocxDiff: the output no longer carries orphaned comment definitions.** The output's `word/comments.xml` is cloned from the LEFT/original document, and `MergeRightCommentDefinitions` adds the RIGHT's referenced comments — but nothing removed a LEFT comment whose annotated content the diff fully replaced, leaving no `w:commentRangeStart`/`w:commentReference` anywhere (not even inside a preserved `w:del`). Its `w:comment` definition dangled: e.g. a comment-dense document produced **10 comment definitions where Word's compare output emits 6** (4 unreferenced orphans). Word never emits such an orphan. A new pass (`IrMarkupRenderer.PruneOrphanedComments`, run right after `NormalizeComments` — the inverse of that pass's step (C), which drops markers with no definition) removes every `w:comment` referenced by no marker in any story (body + headers + footers), plus its `commentsExtended`/`commentsIds` threading entries (keyed by the comment paragraph's `w14:paraId`). Safe for the round trip: a definition is pruned only when NO marker references it, so neither accept nor reject can resurface the reference — a LEFT comment whose marker survives inside a `w:del` (so reject restores it) is kept. This carries no rendering impact (orphaned definitions carry no anchor and do not render), but it is a genuine content-fidelity correctness fix. Coverage: `DocxDiffCommentPruneTests` (orphan + threading-entry pruned; del-preserved marker kept).
- **DocxDiff: the byte-level accept/reject round-trip content-preservation contract is now guarded by generative and pre-existing-revision tests.** The core guarantee a consumer relies on — `accept(Compare(left, right))` reproduces the RIGHT document's content and `reject(...)` the LEFT document's content, with nothing lost, duplicated, or mangled — was previously spot-checked only on a handful of hand-built pairs. Two additions pin it directly: a generative fuzzer (`DocxDiffFuzzRoundTripTests`) that, for each of N reproducible synthetic pairs (seed-count knob `DOCXODUS_FUZZ_SEEDS`, default 50), runs the full `DocxDiff.Compare` → `RevisionProcessor.AcceptRevisions`/`RejectRevisions` path and hard-asserts the accepted document's whole-body text (paragraphs and tables) equals the right input's and the rejected document's equals the left input's, order-independently (so a move/split reorder cannot mask a genuine drop); and `DocxDiffInputRevisionsRoundTripTests`, which pins the same contract for inputs that ALREADY carry un-accepted tracked changes and moves (the "redline of a redline" case, body + table cells), where the compared and preserved content is each side's accepted view. No external fixtures — all pairs are synthesized. Verified with zero content loss across thousands of seeds.
- **DocxDiff: a residue paragraph pair sharing only ONE incidental word is now full-replaced (separate ins + del paragraph) like Word, not force-interleaved.** When exactly one free left and one free right paragraph survive similarity pairing in a gap, the 1×1-residue rule paired them as `Modified` if they shared *any* word — so a dissimilar tail-paragraph rewrite that happened to share a lone common word (e.g. `"… a standard readable font size …"` vs `"Medium-large font sizes …"` sharing only "font") rendered as one mixed paragraph interleaved around that incidental "font", where Word emits a clean inserted paragraph + deleted paragraph (no anchor on the lone word). `IrBlockSimilarity.ResidueForcePair` now requires **≥ 2 shared content words** to force the interleave (zero-overlap already stayed separate; the one-word-vs-one-word typo/renumber case — `"Nested." → "Nexted."` — is still exempt). This matches Word's interleave-vs-full-replace cutoff decoded from Word's compare output. The accept ≡ right / reject ≡ left contract is unaffected (the pair becomes a plain delete + insert, both round-trip trivially). Coverage: `IrSplitMergeTests`/`IrAlignment*` (residue classification).
- **DocxDiff: a paragraph merge/split whose surviving paragraph ADDS text is reproduced as Word's cross-boundary interleave, not a whole-paragraph delete.** When two base paragraphs fuse into one revised paragraph that also inserts new prose (e.g. `"A."` + `"B spreads across the line."` → `"A which C. B spreads across the line, plus more."`), Word deletes the paragraph mark and keeps each base paragraph's surviving words as retained anchors *in that paragraph's slot*. The IR merge/split containment scan (`IrBlockAligner.DetectOneToManyInGap`) only fired when the run covered ≥ 90 % of the surviving paragraph's content (`SplitCoverageThreshold`), so any merge that added substantial new text was vetoed — the extra base paragraph fell through to a whole-block delete and the shared anchors landed in the wrong paragraph versus Word. A new **added-text acceptance path** (`MergeSplitAllowAddedText`, default on) fires the merge/split when the run is genuinely retained (foreign slack ≤ `SplitAddedTextMaxSlack`) and the surviving paragraph is still substantially explained (coverage ≥ `SplitAddedTextMinCoverage`), waiving only the high singular-coverage requirement; the unchanged ≥2-phrase-member gate keeps it from gluing incidentally-overlapping paragraphs. `FindQualifyingRun` returns the coverage-path window eagerly (shortest-first, unchanged) but the added-text window only at maximum coverage, so a clean multi-member split still fires at its complete member count. Scoped conservatively to low-slack merges it can render faithfully — looser "messy" merges/splits (which need Word's global minimal-edit tie-break and split paragraph-property fidelity) are deliberately left for a follow-up. The accept ≡ right / reject ≡ left contract is preserved (the machinery is the existing `IrSplitSegmenter`/`RenderMergeBlock` path). Coverage: `IrSplitMergeTests`; documented in `docs/ooxml_corner_cases.md`.
- **DocxDiff: a wholly-rewritten table row that gains a column keeps its base cells in their original columns.** When a Modified table row shares no cell body with its counterpart (a full cell-for-cell rewrite) AND grows by a column, the ordinary-grid cell fill had no affinity to anchor on, so its cost-tie resolved to inserting the new cell at the FRONT — sliding every base cell one column right and scattering the deleted content into the wrong columns versus Word (which keeps the base cells in place and clean-inserts the new column). `IrTableDiffer`'s cell-gap alignment now applies a small positional tie-break (a per-column-of-displacement affinity nudge, `CellPositionalTieBreak`, 10% of max affinity — more conservative than the row aligner's own locality prior, so any genuine cell-body affinity still dominates and only otherwise-tied rewritten rows are affected). Byte-identical wherever cell bodies carry real mutual affinity (the retained-cell-edit-plus-insert case is unchanged). Coverage: `IrTableDifferTests.Wholly_rewritten_row_that_gains_a_column_keeps_base_cells_in_their_columns`.
- **DocxDiff: fixed-width tables render where Word's compare puts them — hairline cell-margin/indent inset backfill.** A fixed-width table (`w:tblW w:type="dxa"`) that declares no explicit cell margins used to reach the output unchanged, so a renderer applied its OWN default cell margin (LibreOffice ≈ 108 twips) which overflows the declared column widths and shifts every cell's text horizontally versus Word — the whole table "ghosts" against Word's redline. Word's compare output normalizes this: it materializes a hairline `w:tblCellMar` (left/right) plus a matching `w:tblInd`, the inset equal to the table's border width (a 0.5pt `w:sz="4"` border ⇒ 10 twips; `w:sz` is eighths of a point, 1pt = 20 twips). A new single-owner post-pass (`WordCompareTableNormalizer`, applied over the assembled body blocks so every table on every render path is treated identically) reproduces exactly that: for a fixed-width table with a derivable border width and no declared `w:tblCellMar`, it inserts the border-width inset as `w:tblCellMar` + `w:tblInd` in CT_TblPrBase schema order. AUTO-width tables (`type="auto"`), tables with no border, and tables that already declare cell margins are left untouched. Mirrors the existing docDefaults backfill (`WordStockDocDefaults`/`DocxDiffDocDefaultsBackfillTests`) — the whole engine goal is to reproduce Word's compare output, and a rendered redline is only faithful if its tables land where Word's do. The `accept ≡ right` / `reject ≡ left` contract is unaffected (the inset is a table property, not tracked-changes markup, and the round-trip is verified at body-text level). Coverage: `DocxDiffTableCellMarginBackfillTests`; documented in `docs/ooxml_corner_cases.md`.
- **DocxDiff: content-anchored intra-paragraph token diff — interior shared words are retained instead of dropped.** The Myers token diff previously keyed on ALL tokens including whitespace separators (which all share one match key), so a paragraph with many identical spaces spent its LCS budget on whitespace and DELETED+re-inserted interior shared content words — e.g. a lone "a" between two divergent phrases — scattering ins/del ink where Microsoft Word anchors on the content word and retains it in place. `IrTokenDiffer.MyersSpans` now runs the LCS over the NON-whitespace token subsequence (content anchors), then re-expands whitespace positionally per inter-anchor segment (a common whitespace prefix/suffix stays `Equal`; the middle is a clean `Delete`+`Insert`). This reproduces Word's whole-sentence-replace-with-retained-anchors shape on formatting-change documents. The accept ≡ right / reject ≡ left contract and the token-diff tiling invariants are preserved (verified by `IrDiffFuzzTests`); coverage in `IrTokenDifferTests`. Known limitation: on two low-similarity sentences that happen to share an incidental repeated word, the anchor can be retained where Word full-replaces.
- **DocxDiff: `FormatChanged` runs spanning heterogeneous left formatting now restore each region's format on reject.** When a `FormatChanged` token span covered LEFT source runs with different formatting (e.g. a bold run + an italic run) but a single RIGHT run, `IrMarkupRenderer.BuildTokenOpContent` stamped ONE `w:rPrChange` from the first left char, so rejecting restored the first format across the whole span and lost the second. It now splits each right-run slice at LEFT source-run format boundaries (`SourceRunModel.FormatBoundaries`), emitting one `w:rPrChange` per left-format region — `reject ≡ left` holds at the property-byte level. Byte-identical on single-uniform-run input. Coverage: `DocxDiffWordShapeTests.IntraParagraphReplace_Reanchoring*`.
- **DocxDiff: an edited table bracketed by asymmetric blank paragraphs is no longer torn into a whole deleted + whole inserted table.** When base and next differed in leading/trailing blank (whitespace-only) paragraphs around a table, a fungible blank got matched across the table and `IrBlockAligner.ReleaseCrossingModifiedPairs` released the whole table's `Modified` pairing into `Deleted`+`Inserted` — emitting two tables where Word keeps one table with native per-row `w:trPr/w:ins|w:del` markup. The crossing resolution is now weight-aware: a `Modified` structural pair crossed only by a fungible blank-spacer pair (`IrBlockSimilarity.IsBlankSpacer`) demotes that blank to `Deleted`+`Inserted` instead of releasing the heavier table; genuine reorders (crossing content-bearing blocks) still release as before. Byte-identical when no blank-spacer pair is present. Coverage: `IrTableCrossingReleaseTests`.
- **DocxDiff: Word-shaped redline projection — a fidelity campaign aligned to Microsoft Word's own compare output.** A coordinated set of renderer/aligner/consume-side changes that make DocxDiff's tracked-changes output *render* like Word's redline. All changes preserve the accept ≡ right / reject ≡ left contract; the edit script keeps its token grain — these are projection/consume-side changes:
  - **Replace-gap grammar** (`RenderBlockOpsWordShaped`): inserted blocks render BEFORE deleted ones within a delete+insert gap; the last inserted and first deleted paragraph share one `w:p` (the seam, deleted-side pPr + tracked mark, guarded against inline `w:sectPr` and page-break carriers); the deleted chain ends at a live terminator mark so accept ends the inserted text there.
  - **Token-level coalescing** (`CoalesceTokenOpsWordShaped`): a changed region inside a paragraph renders as ONE inserted region then ONE deleted region (interior whitespace consumed into both sides), never per-word del/ins alternation.
  - **Aligner** (`IrBlockAligner`): in-gap similarity pairing gains a locality prior (`sim ≥ threshold + 0.3·displacement`; kills cross-gap "word salad" while keeping high-similarity swapped edits); `BlockSimilarityThreshold` recalibrated 0.5 → 0.35 against Word's compare output; split/merge groups require ≥ 2 members; leftover gap tables pair positionally (k-th old merges into k-th new via the per-cell table diff).
  - **Style definitions** (`TrackStyleDefinitionChanges`): the output keeps the LEFT styles part (docDefaults/theme byte-identical to the original, as Word does) while styles whose RAW definitions differ get their current payload updated — pPr side at raw payloads, rPr side materialized from the resolved chain — with the old payload archived in a tracked `w:rPrChange`/`w:pPrChange` inside the definition; right-only styles copied, left-only styles survive.
  - **Stories** (`RebindOrStripStoryReferences` + `EnsureStoryReference`): header/footer references carried by right-cloned inline sectPrs are rebound to the output's own story parts (matched → the merged left part, right-only → the inserted part, else a pruned wholesale import); a matched story whose only left reference lived on a collapsed inline sectPr is re-attached instead of orphaned; `w:titlePg`/`w:evenAndOddHeaders` activate only when the revised document activates them; unresolvable references are dropped (OOXML inheritance).
  - **Inputs**: strict-conformance packages (ISO 29500 `purl.oclc.org` namespaces) are normalized to transitional at every entry point (`StrictOoxmlNormalizer`); `DocxCompare`'s DocxDiff branch pre-accepts input revisions (WmlComparer/Word parity); dangling `numId` references get Word's synthesized decimal-multilevel numbering repair.
  - **Robustness**: cross-kind relationship-id collisions remap instead of throwing; `RevisionProcessor` accepts Word's paragraph-mark `w:moveFrom` sentinel like a deleted mark (no more spurious empty paragraph at every Word-authored move source; RP015 baselines regenerated — they had captured the old bug), and drops emptied hyperlink shells on accept/reject.
- **DocxDiff: strict-conformance OOXML inputs (ISO/IEC 29500 `purl.oclc.org` namespaces) no longer fail with "Document has no w:body element".** Word's "Strict Open XML Document" save format keeps every WML element in the strict namespace family, which the XDocument-based IR reader (and everything downstream) does not speak. A new internal `StrictOoxmlNormalizer` detects a strict main part (resolved through `_rels/.rels`, either conformance class's officeDocument relationship type) and rewrites every XML part and `.rels` stream to the transitional namespaces — including the `extendedProperties` → `extended-properties` rename and dropping Word's `w:conformance="strict"` root attribute — before any read, exactly as Word normalizes strict packages on open. Applied at every `DocxDiff` entry point (`Compare`/`GetRevisions`/`GetEditScriptJson`/the consolidate family) via the shared input-preparation hook; transitional inputs pass through untouched (same instance). The accept ≡ right / reject ≡ left contract holds whether one or both sides are strict. Regression coverage: `DocxDiffStrictOoxmlTests` (strict-left, strict-right, both-strict, self-compare identity, edit-script JSON).
- **DocxDiff: a right-document `r:id` colliding with a left relationship of a DIFFERENT KIND no longer throws `XmlException` ("'rIdN' ID conflicts with the ID of an existing relationship").** `IrMarkupRenderer.ImportHyperlinkAndExternalRelationships` treated an id as free when no left *hyperlink* used it — but the id could name a left **part** relationship (comments.xml, an image, …), and recreating the right hyperlink/external relationship under that taken id makes System.IO.Packaging throw an exception type the guard didn't catch. The freeness test now consults ALL left relationship kinds (parts, hyperlinks, externals, data-part references) and the different-kind collision takes the existing fresh-id remap path (`rIdRemap{n}` + rewrite of the cloned `r:id`s), so the inserted content still resolves to the right-side target. Regression coverage: `DocxDiffRelationshipRemapTests`.
- **RevisionProcessor: accepting/rejecting a fully-deleted (or, under reject, fully-inserted) hyperlink no longer leaves an empty `w:hyperlink` shell or a stray empty paragraph.** Two consume-side gaps, both the hyperlink analogue of the existing wholly-deleted-table rule: (1) `AcceptAllOtherRevisionsTransform` now drops a `w:hyperlink` whose content children all sit inside `w:del`/`w:moveFrom` (preserving any bookmark markers inside it); (2) `AllParaContentIsDeleted`'s collapse step now sees THROUGH a hyperlink shell, so a trailing paragraph whose only content was an inserted hyperlink is removed on reject like its plain-run counterpart (previously `reject(Compare(left, right))` left an extra empty paragraph behind — a body-text divergence from `left`). Benefits every producer whose markup nests revisions inside `w:hyperlink` (the IR renderer's convention, and Word files with the same shape).

## [7.1.0] - 2026-07-12

### Changed
- **DocxDiff.Consolidate: a contested relocation now honors the conflict policy on the PLACEMENT (issue #233).** When two (or more) reviewers move the SAME base block to DIFFERENT destinations, the reviewers agree the block leaves its origin but disagree on where it lands. Previously the merger recorded a placement conflict but let **every** reviewer's relocating insert land on accept regardless of policy — so the moved block appeared at **both** destinations under `BaseWins`, `FirstReviewerWins`, and `StackAll` alike. The placement is now resolved by the `ConflictResolution` policy: `BaseWins` keeps the block at its **base position** (neither move applied → `accept ≡ base`); `FirstReviewerWins` applies **only the first reviewer's** (list-order) destination; `StackAll` keeps the both-placements behavior. In every case there is **no content loss**, `reject ≡ base` holds, and the conflict is still recorded for human resolution. Implemented in `IrCompositeMerger`: the lowered move DESTINATION now retains its `(reviewer, MoveGroupId)` marker (symmetric to the source-delete, stripped before emission) so `PlanContestedRelocationSuppression` can suppress the losing destination(s) in lockstep with `MergeOneBaseBlock`'s base-keep-vs-consensus-removal choice (both keyed off the single `IsContestedRelocation` predicate). The consensus removal is emitted at most once, so a delete both reviewers made is never duplicated. Confined to the IR-based DiffDocx engine — WmlComparer is unchanged. `StackAll` output is byte-identical to before. Proof: `IrCompositeMoveTests.Contested_relocation_{BaseWins,FirstReviewerWins,StackAll}_*` plus the unchanged composite fuzzer, the 84/84 consolidate parity scoreboard, and the 179/179 two-way parity floor.

### Added
- **`DocxSession`: header / footer / page-number authoring (issue #236).** The mutation API could *inspect* a section's header/footer parts (`SectionInfo.HeaderPartUris`/`FooterPartUris`) but had **no** surface to create or edit them — blocking faithful reproduction of real filings (the S-1 smoke test's running footer + centered page number could not be authored). Three new methods close the gap, rippled through every layer (core → `DocxSessionOps` → WASM/npm → stdio/`docx-scalpel`):
  - **`SetHeaderText(anchorId, HeaderFooterKind, markdown)` / `SetFooterText(...)`** — set the running header/footer story for the section that owns `anchorId` (any body block in that section; the governing `w:sectPr` is resolved exactly as `GetSectionInfo` resolves it, synthesizing a trailing `w:sectPr` if the body has none). Creates the `HeaderPart`/`FooterPart` + relationship + `w:headerReference`/`w:footerReference` when the story of that `kind` is absent, else replaces its content. The content is the same markdown subset as `InsertParagraph`, styled with the built-in `Header`/`Footer` paragraph style so it inherits Word's centre/right tab stops. `HeaderFooterKind` = `Default`/`First`/`Even`; `First` sets the section's `w:titlePg`, `Even` sets `w:evenAndOddHeaders` in the settings part, so Word actually shows the story. The created header/footer paragraph anchors (scope `hdr{N}`/`ftr{N}`) come back in `EditResult.Created`.
  - **`InsertPageNumberField(anchorId, PageNumberField = CurrentPage)`** — append a native Word complex page-number field (`fldChar`/`instrText` with a cached "1") to the paragraph `anchorId` (typically a footer paragraph from `SetFooterText`). `CurrentPage` → `PAGE`, `TotalPages` → `NUMPAGES`. Returns the paragraph anchor in `EditResult.Modified`. Compose `SetFooterText` + `SetParagraphFormat(center)` + `InsertPageNumberField` to reproduce the S-1's "Last Updated … / page N" footer.
  - **Undo/redo of the part creation.** The session's snapshot/restore was extended to reconcile header/footer part create/delete (previously only the annotations custom-XML part was reconciled — see the resolved TODO in `RestoreSnapshot`): the snapshot records each header/footer part's relationship id, and restore deletes parts the snapshot lacks and re-creates the ones it has *with their original relationship id* so the restored `sectPr` reference resolves. `reject`/content round-trips are proven by `DocxSessionTests` `DS250`–`DS262` (create/reuse/compose/First/Even/undo-redo/no-sectPr/wrong-kind/round-trip). (One documented edge: the `w:evenAndOddHeaders` settings flag — only set for `Even` — is not reverted by undo; it is idempotent and has no visual effect without an even story.)
  - **Surfaces.** WASM `DocxSessionBridge.{SetHeaderText,SetFooterText,InsertPageNumberField}` + npm `DocxSession.{setHeaderText,setFooterText,insertPageNumberField}` (new `HeaderFooterKind`/`PageNumberField` string-union types); stdio `set_header_text`/`set_footer_text`/`insert_page_number_field` + `docx-scalpel` `DocxSession.{set_header_text,set_footer_text,insert_page_number_field}` (new `HeaderFooterKind`/`PageNumberField` enums). The editor's visual header/footer editing region is intentionally deferred (the parts live outside the body); this ships the engine + wire the editor will drive.
- **`docx2html` CLI: tracked-changes and story rendering flags.** The `docx2html` tool previously always accepted revisions before converting, so a redline document rendered as if every change had been applied — there was no way to preview a redline in a browser. New flags map straight onto the existing `WmlToHtmlConverterSettings`: `--track-changes` (render revisions as `ins`/`del`/move markup instead of accepting them — `RenderTrackedChanges`), `--no-render-moves` (lower move markup to plain ins/del — `RenderMoveOperations`), `--render-comments`, `--render-footnotes`, and `--render-headers-footers`. Defaults are unchanged (all off, moves on), so existing invocations behave identically. This is the preview half of the docx-redlines GitHub Action (JSv4/Python-Redlines#12): the action redlines changed `.docx` files with the `redline` CLI and uses `docx2html --track-changes` to publish browser-viewable previews.

### Fixed
- **Setting `w:evenAndOddHeaders` no longer corrupts `word/settings.xml` on documents whose settings part carries children the ordering table doesn't know.** Both the DocxDiff header/footer renderer (shipped with `CompareHeadersFooters`) and the new `DocxSession.SetHeaderText(…, Even, …)` inserted the flag by routing the whole settings root through `WmlOrderElementsPerStandard`, whose `Order_settings` table lacked `w:hdrShapeDefaults`/`w:shapeDefaults` — elements real Word documents (e.g. `TestFiles/DB006-Source2.docx`) carry — so those children were sorted to the end, out of their CT_Settings schema slots (`OpenXmlValidator`: *"unexpected child element 'hdrShapeDefaults'"*). The two missing entries are now in the table, and both call sites share one `WordprocessingMLUtil.EnsureEvenAndOddHeaders` that inserts the flag at its own schema slot and leaves every other settings child untouched (unknown children are never moved). Regression coverage: `DocxSessionTests` `DS263` (synthetic settings with `hdrShapeDefaults`/`shapeDefaults`) + `DS264` (the real Word-authored fixture that caught it).
- **DocxDiff: a single `w:hyperlink` whose entire anchor is replaced no longer fragments into per-token link elements (issue #232).** When a hyperlink-wrapped run was involved in a redline and its anchor text was *fully* rewritten with no shared token (e.g. `our website` → completely different words, same href), the IR markup renderer emitted the base's single `w:hyperlink` as **two** elements sharing the same `r:id` — a pure `w:del`-link of the old text followed by a pure `w:ins`-link of the new text. The visible text, link target, and accept/reject round-trip were all correct (valid OOXML), but the raw markup diverged from the source's link structure (1 → N), churning any tool that diffs the XML or counts `w:hyperlink` elements. This was the residual normalization left by the B1 fix (#228), whose coalescer merged link fragments only when at least one carried a *plain* (Equal) run — a proxy for "same target" that missed the no-shared-token case. The coalescer now stamps each fragment with its **resolved** target (external URI, or `#anchor` for an internal link — resolved exactly as `IrReader` does, via the part annotation on the retained source tree) and additively merges a same-source-link run when every fragment resolves to the same target. This keys on the resolved target, **not** the `r:id` string, so it still keeps the WC019 whole-anchor **retarget** (text *and* href change → the del/ins fragments carry the same `r:id` string at coalesce time but different resolved targets) split into two links for the post-assembly `r:id` remap. Scope: the base two-way `IrMarkupRenderer` (`Docxodus/Ir/Diff/IrMarkupRenderer.cs`); `WmlComparer` is untouched. B1 invariants (`accept ≡ right`, `reject ≡ left`) and the adjacent-distinct-same-target-links case are preserved; base two-way parity stays 179/179. Regression coverage: `IrMarkupRendererTests.Hyperlink_fully_replaced_same_target_renders_as_one_hyperlink`, `Hyperlink_whole_anchor_retarget_stays_two_links`, and an updated `Hyperlink_single_token_anchor_no_equal_run_replaced_round_trips` (now asserts one link).
- **`GetDocumentMetadata` now detects section breaks nested inside tables, and correctly ignores section properties inside text boxes (issue #51).** `CollectSectionData` previously scanned only the body's direct children (`body.Elements()`), so a `w:sectPr` carried by a paragraph inside a table cell was invisible — the document reported one section instead of two, and the per-section page dimensions and paragraph/table index ranges used for lazy-loading pagination were wrong. It now walks the body as a single **main story** in document order, descending into tables (`w:tbl → w:tr → w:tc`) so an in-cell section break is counted and attributed to the section it starts in. Each `w:p` is treated as a leaf (only its direct `w:pPr/w:sectPr` is inspected, never its runs), so a `w:sectPr` inside a **text box** — a separate story that does not paginate the main document — is excluded automatically (as are text-box paragraphs); counting it would have invented a phantom section. Only body-level tables count toward the table total, preserving the previous counts for the common case. Regression coverage: `DM022` (in-cell break detected) and `DM023` (text-box section properties ignored) in `DocumentMetadataTests.cs`; the separate-story rule is documented in `docs/ooxml_corner_cases.md`.

## [7.0.1] - 2026-07-11

### Changed
- **New project logo.** The README banner now uses the new Docxodus hero artwork (`docxodus-logo.png` — the scene + wordmark + tagline cut), replacing the old `docxodus-mono-final.svg` lockup, which has been removed.

### Fixed
- **`ConvertToHtml` no longer crashes when `word/styles.xml` is absent (issue #265 — sibling of the #264 settings fix).** `StyleDefinitionsPart` is also optional in OOXML (Word opens a document-only package without repair), but `FormattingAssembler.AssembleFormatting` dereferenced it unconditionally at many sites — and it runs before the tab-stop code fixed in #264 — so a package with only `word/document.xml` still threw `ArgumentNullException("part")` before producing any HTML. `AssembleFormatting` now synthesizes an empty `w:styles` part when the source has none, so every style lookup falls through to built-in defaults; this covers the full-document, `RenderBlockHtml`, and session-render paths in one place (and any direct `FormattingAssembler` caller). Regression coverage: `HCO062`–`HCO064` (programmatic document-only package, the in-repo `RPR-FivePageTestDoc.docx` document-only fixture, `RenderBlockHtml` on a styles-less source).
- **`ConvertToHtml` no longer crashes when `word/settings.xml` is absent.** `DocumentSettingsPart` is optional in OOXML (ECMA-376 does not require it, and Word opens such packages without repair), but `CalculateSpanWidthForTabs` called `DocumentSettingsPart.GetXDocument()` unconditionally, so any minimal package without a settings part threw `ArgumentNullException("part")` before producing any HTML. Settings are now treated as optional: when the part (or its `w:defaultTabStop`) is missing, the converter keeps Word's implicit 720-twip default tab stop; `w:defaultTabStop` is read as a direct child of `w:settings` (per the schema) instead of a full `Descendants` walk; and `HtmlConversionOps.AddFormattingParts` no longer synthesizes a throwaway settings part for `RenderBlockHtml`. Regression coverage: `HCO057`–`HCO061` (missing part, 720-twip fallback, custom `w:defaultTabStop`, empty settings, `RenderBlockHtml` on a settings-less source). (#264)
- **`scripts/build-wasm.sh`: per-asset `integrity` keys in `dotnet.boot.js` are normalized to `hash`** so the .NET 10 loader — which honors only `hash` (see dotnet/runtime#122391) — applies subresource integrity to framework asset fetches. A no-op on SDKs that already emit `hash`. (#264)

## [7.0.0] - 2026-07-09

### Added
- **Shared comparison-engine selector across the CLI / WASM / npm surfaces (M-B) — seeded to `WmlComparer`, default NOT flipped.** Introduces one selector so a caller can choose which engine redlines two documents, wired identically everywhere, as the seam the eventual D4 cutover flips in one line:
  - **One selector type + one dispatch owner.** New public `ComparisonEngine` enum (`WmlComparer = 0` default, `DocxDiff = 1`) and `DocxCompare.Compare(left, right, engine, settings)` — the single `WmlComparer`-vs-`DocxDiff` branch in the codebase (mirroring the `DocxDiffOps`/`HtmlConversionOps` single-owner pattern). It takes the incumbent `WmlComparerSettings` (the shape all surfaces already build) and, on the `DocxDiff` branch, maps the common option set to `DocxDiffSettings` via `DocxCompare.ToDocxDiffSettings` (author/date/case-insensitive/conflate-spaces/detect-moves/move-thresholds carry; the WmlComparer-only knobs `DetailThreshold`/`SimplifyMoveMarkup`/`DetectFormatChanges` are dropped). Both engines emit native tracked-changes markup, so revision counting stays uniform on the output document.
  - **CLI (`redline`).** New `--engine=wmlcomparer|docxdiff` flag (default `wmlcomparer`; `--detail-threshold` documented as WmlComparer-only). Backed by `DocxCompare.TryParseEngine` (case-insensitive, trims, rejects unknown).
  - **WASM / npm.** The four primary redline paths (`CompareDocuments`, `CompareDocumentsWithOptions`, `CompareDocumentsToHtmlWithOptions`, `CompareDocumentsToHtmlFull`) gain a trailing `engine` int routed through `DocxCompare`; npm adds a `ComparisonEngine` enum and an optional `CompareOptions.engine`, threaded through `index.ts`, the off-main-thread worker, and the export typings. Omitting the selector reproduces today's behavior exactly. The `*WithLog`/`GetRevisionsJson` helper variants stay WmlComparer-backed for now (DocxDiff has its own revisions surface).
  - **No behavior change.** `ComparisonEngine.WmlComparer` (0, the default) delegates straight to `WmlComparer.Compare`. Proven by `DocxCompareTests` (default reproduces WmlComparer revisions; DocxDiff branch equals a direct `DocxDiff.Compare`; settings map carries the common fields; the enum integer contract) and a `npm/tests/engine-selector.spec.ts` Playwright spec. The default is deliberately **not** flipped — that remains gated on D4.
  - **Note:** the WASM `DocumentComparer` JSExport signatures and the `DocxodusWasmExports.DocumentComparer` low-level typings gained the `engine` parameter (a breaking change at that raw boundary); the public npm `compareDocuments*`/`CompareOptions` surface stays back-compatible (`engine` is optional).
- **Editor edit-latency: multi-block format ops now reconcile incrementally, and remount renders session-attached — no fidelity change.** Two performance fixes for the browser editor's interactive latency, both pinned by a new fidelity spec (`npm/tests/editor-perf-incremental.spec.ts`):
  - **Multi-block formatting is now N single-block swaps, not a whole-document re-render.** `DocxEditor`'s multi-block paths (`format`/`setFontSize`/`setFontFamily`/`setAlignment`/`indent`/`pageBreakBefore`/`setParagraphStyle` over a multi-paragraph selection) previously fell back to a full remount — `Save()` + `ConvertDocxToHtmlComplete` + full DOM rebuild, i.e. the full-document convert (~1–2.5 s on a real document) on every ribbon click — where the single-block path swapped just the edited block (~10 ms). The multi-block path now applies each block's op and swaps each edited block in place via the session-attached `RenderBlockHtml` — exactly the single-block path run N times, so rendering is fidelity-identical by construction — and restores the cross-block selection afterward (previously the selection was silently dropped), so consecutive ribbon actions (center, then bold) keep targeting the same range. The full remount is retained where whole-document context is genuinely required: any result touching a list item (numbering continuation), `clearParagraphBorders` (border-div regrouping), and paginated mode (page reflow, until M4 lands a scoped re-paginate).
  - **`DocxSessionBridge.RenderHtml` — session-attached full-document render.** New WASM export (`DocxSessionOps.RenderHtml`) that renders the live session's current state to the complete anchor-stamped HTML document inside the runtime, replacing the editor remount's `Save()` → marshal bytes to JS → marshal bytes back into WASM round trip (two multi-MB copies per remount on a large document). The option profile matches the editor's `ConvertDocxToHtmlComplete` call exactly and the output is byte-identical (asserted by the new spec); `DocxEditor` falls back to the old bytes path when the export is absent (older WASM bundles). Exposed on `DocxodusWasmExports.DocxSessionBridge` in npm `types.ts`. WASM/npm-only, like `RenderBlockHtml` (not part of the stdio/python surface).
- **DocxDiff: `w:gridSpan` / `w:vMerge` property-only table changes — scope closed and pinned (Issue #230).** #230 (an IR-modeling gap where a reviewer's *property-only* table/cell change read as unchanged) was resolved for cell shells by the block-format-change family (native `w:tcPrChange`, `IrCell.ShellDigest` folded into the cell `ContentHash`) and by consolidate cell-shell composition. This change closes the issue's explicit remaining acceptance criterion — **documenting the chosen `gridSpan`/`vMerge` scope** — and adds the direct regression proof that was missing: a `gridSpan`- or `vMerge`-only change (cell count stable, text unchanged) is now proven to be tracked as a native `w:tcPrChange` `TableCell` `FormatChanged` revision in 2-way `Compare` (`accept ≡ right`, `reject ≡ left` at the tcPr-byte level), and to compose in `Consolidate` like a width/shading edit. A `gridSpan` change that alters the cell *count* (column add/remove) is detected (never silently invisible) and composes per-cell in `Consolidate` (`w:cellIns`/`w:cellDel`); the 2-way single-toucher path lowers it to a whole-table del/ins (a pre-existing renderer-granularity limit, not a soundness gap). No engine change — the behavior already held; this adds the fixtures (`BlockFormatChangeTests.{GridSpanOnly,VMergeOnly}_cell_change_is_tracked_with_native_tcPrChange`, `IrCompositeTableTests.VMerge_only_cell_edit_composes`) and the scope decision in `docs/architecture/ir_diff_engine.md`.
- **DocxDiff.Consolidate: reviewers' table-shell and section format changes now MERGE (sub-project B2) — closing the last "Consolidate ignores block-format" ceiling.** Building on B1 (paragraph `w:pPr` merge), the N-way merge now composes the remaining block-format families with per-reviewer attribution + native Word markup:
  - **Table-shell family** (`w:tcPrChange`/`w:trPrChange`/`w:tblPrChange`/`w:tblGridChange`/`w:tblPrExChange`). Each shell element is attributed **independently** (per-cell `tcPr`, per-row `trPr`+`tblPrEx`, per-table `tblPr`+`tblGrid`) by its digest, mirroring the existing cell-shell composition: 0 changers → base, all reviewers agree → consensus (first reviewer), ≥2 distinct → a recorded `DocxDiffConflict` resolved by policy (a shell cannot stack). Per-element granularity means disjoint edits COMPOSE (one reviewer's `tblPr` + another's `trPr` both land) while only a genuinely-contested element conflicts — and it composes cleanly with #250's column-add/remove and row-move. A single-reviewer shell edit rides the two-way single-source render; a multi-reviewer table routes through the unified `ComposeTableAndRowShells`. This also fixes a #250-era gap where a composed cell's shell was swapped in with **no** `w:tcPrChange` marker (so reject kept the reviewer's shell) — the marker is now stamped with inner = base.
  - **Section family** (`w:sectPrChange`, trailing + inline). The document-final `w:sectPr` (not a body block op — Word compares it at the document level) is composed by `ComposeTrailingSection` across each reviewer's trailing `IrSectionBreak` (modeled page setup + the unmodeled-digest catch-all), and the composite renderer stamps `w:sectPrChange` (inner = base, header/footer references preserved). A mid-document inline (`w:pPr/w:sectPr`) section change rides B1's paragraph FormatOnly path.
  - **Mechanism.** The internal `TrackBlockFormatChanges` is now sliced into paragraph/table/section flags (`TrackParagraphFormatChanges` from B1 + new `TrackTableFormatChanges`/`TrackSectionFormatChanges`); the public `DocxDiffSettings.TrackBlockFormatChanges` opt-out cascades to all three. The composite turns all three slices ON while forcing the umbrella OFF; **two-way `Compare` is byte-identical** (the slices default equal to the umbrella), proven by the full renderer battery + a 470-case two-way/composite regression.
  - **Round-trip.** `reject ≡ base` and `accept ≡ the policy-resolved composite` now hold at the **property-byte** level for every shell/section family — enforced by a strengthened byte-level verifier: a new `Docs.ShellSection` canonical projection (over every body `w:tcPr`/`trPr`/`tblPr`/`tblGrid`/`tblPrEx` + trailing `w:sectPr`) is asserted alongside the previously text-only reject=base checks in `CompositeFuzzTests` (3/4/5-way) and the new `ConsolidateBlockFormatB2Tests` (per-family reject=base/accept=winner/conflict). Without this strengthening a lost shell passed silently.
  - **text + format (v1 decision).** A reviewer editing a paragraph's TEXT *and* its `pPr` is **conflict-routed**, not inline-composed: the existing `ParagraphPropsUnchanged` guard keeps a cross-reviewer text+pPr collision out of token-composition (a recorded conflict, never a silent format drop); a single reviewer's text+pPr edit tracks both. True inline text+format compose is deferred to a follow-up. **No silent drop anywhere** — every reviewer format edit is attributed (per-element winner) or recorded as a conflict.
  - **Proof.** `ConsolidateBlockFormatB2Tests` (25 cases: per-family single/multi-reviewer merge, consensus, conflict-per-policy, disjoint-compose, mixed cell+row+section, inline sectPr, text+format), the strengthened `CompositeFuzzTests` byte-level guard, the 84-case `ConsolidateParityScoreboardTests` (unchanged, still 84), and `OpenXmlValidator`-clean output. Flips the B1 `Consolidate_merges_pPr_but_not_shell_section_v1` ceiling pin.
- **DocxDiff.Consolidate: reviewers' paragraph-property (`w:pPr`) changes now MERGE (sub-project B1).** Previously a reviewer's paragraph-formatting-only edit was ignored by `Consolidate` (the N-way merge forced block-format tracking off). Now a reviewer's pPr-only change (alignment/indent/spacing/style/numbering, text unchanged) is composed across reviewers with per-reviewer attribution: a new `ComposePPr` (mirroring the existing cell-shell composition) attributes the change by a new `IrParagraph.PPrDigest` — 0 changers → base, all reviewers agree → consensus (first reviewer), ≥2 distinct → a recorded `DocxDiffConflict` resolved by the `ConflictResolution` policy (a pPr cannot stack — one paragraph has one pPr, so `StackAll`/`FirstReviewerWins` apply the first changer). The consolidated document carries native `w:pPrChange` authored to the winning reviewer; **reject ≡ base and accept ≡ the policy-winner hold at the property-byte level** (proven by a multi-reviewer byte-level round-trip stress test). Implemented via a paragraph slice of the internal flag (`TrackParagraphFormatChanges`) so two-way `Compare` behavior is byte-identical and the composite fuzzer (3/4/5-way) + 84-case parity scoreboard are unaffected. **Table-shell and section merge followed in sub-project B2 (above); a reviewer who changed BOTH a paragraph's text and its pPr remains conflict-routed (v1 decision).** Flips the former `Consolidate_ignores_block_format_changes_v1_ceiling` pin.
- **DocxDiff block-format-change family — follow-up A: the two-way surface is now complete.** Building on the family that shipped in the same release:
  - **`w:tblPrExChange`** (row-level table property exceptions) is now tracked (was visible-but-untracked). A new `IrRow.TrPrExDigest` (a flattened `tblPrEx`-only projection) drives a native `w:tblPrExChange` marker + a `TableRow`-scope revision with the distinct `"tblPrEx"` changed-name; reject restores the left bytes.
  - **Mid-document `w:sectPrChange`** (an inline `w:sectPr` inside a `w:pPr`) is now tracked (was invisible — the inline sectPr's properties weren't modeled). The reader models it as `IrParagraph.InlineSectionFormat`, folded into the paragraph `FormatFingerprint` **and** `IrModeledFormat.BlockSignature` so a sectPr-only change classifies FormatOnly under `ModeledOnly`; the emit stamps `w:sectPrChange` inside the paragraph's own `w:pPr/w:sectPr` (not the `pPrChange` inner — CT_PPrBase excludes sectPr) with a per-paragraph Section revision. A one-sided add/remove is structural (untracked).
  - **Note-scope and header/footer-scope `w:pPrChange` proven to already work** — they route through the same `RenderBlockOp` dispatch as the body with no per-scope gate, so a changed footnote/header/footer paragraph already emits `w:pPrChange` and reports a Paragraph-scope revision. The former "v1 ceiling" was over-conservative documentation; now pinned by tests.
  - **`TrackBlockFormatChanges` is now a public opt-out** on `DocxDiffSettings` (default `true`; additive wire key `trackBlockFormatChanges` across Ops JSON / npm / docx-scalpel). `Consolidate` still forces it off internally (the N-way merge is a separate follow-up).
  - **Split/merge `w:pPrChange` deliberately declined** — a split's members are brand-new paragraphs already tracked by the inserted pilcrow mark (no per-member left baseline), and a merge's non-final members are deleted; a pPr "change" is not well-defined and would fight the reject-fuse. Pinned as a principled ceiling.
  - Proof: `BlockFormatChangeTests` (per-member), `BlockFormatChangeRealDocTests` (a real corpus doc mutated across the full table + section family), the strengthened `IrMarkupRendererTests` round-trip battery, and additive wire tests (.NET/npm/python).
- **DocxDiff: paragraph-and-above formatting changes are now tracked as native Word markup — closing the last "Word compares Formatting, we don't" scope gap.** Previously only run-level format changes surfaced (`w:rPrChange`); a change to a paragraph's, table's, cell's, row's, or section's properties (with identical text) was either invisible (silently classified unchanged) or applied untracked — the right-side properties won with no revision, so `reject` did not restore the left. Now, gated by the existing `DocxDiffSettings.FormatComparison` policy (default `ModeledOnly`) and produced as the markup Word itself renders:
  - **`w:pPrChange` (paragraph).** An alignment/indent/spacing/style/**numbering** change (direct `w:numPr` numId/ilvl are now modeled in the IR) classifies as a format change and renders a `w:pPrChange` carrying the OLD paragraph properties (and, for a changed paragraph MARK, a `w:pPr/w:rPr/w:rPrChange`). Fires in the format-only, edited, and moved-destination paragraph paths.
  - **`w:tcPrChange`/`w:trPrChange`/`w:tblPrChange`/`w:tblGridChange` (table family).** The pre-existing single lumped table-shell digest is split into per-element `TblPrDigest`/`TblGridDigest`/`TrPrDigest` (flattened so an empty shell ≡ an absent one), so the renderer attributes each change to the exact shell. This makes the cell-shell (`w:tcPr`) edits that became *visible* in #250 actually *tracked* — reject now restores the shell bytes, not just the text.
  - **`w:sectPrChange` (section).** A trailing-section page-size/margin/orientation/columns change stamps `w:sectPrChange` on the body's trailing `w:sectPr` — the right properties applied, the left preserved in the marker, and the header/footer references (owned by the header/footer machinery) untouched. Mid-document section breaks inside a `w:pPr` are a documented v1 ceiling.
  - **Round-trip contract extended.** For every DETECTED change, `accept ≡ right` and `reject ≡ left` now hold at the property-byte level (canonical, reference/rsid-normalized), enforced corpus-wide by the strengthened renderer battery (per-table shell digests + a reference-normalized trailing-sectPr property digest). Under `ModeledOnly`, remaining residual-only paragraph deltas remain untracked right-applies; direct `w:shd` is modeled and tracked on both run and paragraph properties.
  - **Revisions + wire.** `GetRevisions` reports the change via `DocxDiffRevision.FormatChange` carrying a new **`DocxDiffFormatChangeScope`** (`Run` default, plus `Paragraph`/`TableCell`/`TableRow`/`Table`/`Section`); `WmlComparerCompatible` mode excludes every non-`Run` scope by construction (the legacy comparer produces none, keeping the 179-count parity scoreboard meaningful). Additive `scope` field on the revisions wire shape (`DocxDiffOps` → WASM/npm `FormatChangeScope` → docx-scalpel `DocxDiffFormatChange.scope`); the edit-script JSON is unchanged.
  - **Consolidate v1 ceiling (pinned).** `IrCompositeMerger` forces block-format tracking OFF for its per-reviewer diffs (the header/footer precedent) — a reviewer's pPr/shell/section-only edit is ignored by `Consolidate` (explicit, nothing silently dropped); text+pPr edits keep routing to the conflict path.
  - **Proof.** `BlockFormatChangeTests` (per-member detection/markup/revision/round-trip + the Consolidate pin), `BlockFormatChangeRealDocTests` (a real corpus doc mutated across all five family members → all markers present, schema-valid, round-tripping, deterministic, with a headless-LibreOffice load backstop), and the strengthened `IrMarkupRendererTests` battery.
- **DocxDiff: header/footer stories are now compared — the way Word Compare does (its default-on "Headers and footers" granularity).** Previously a header/footer difference was silently ignored: the output carried the LEFT document's parts verbatim, `accept(Compare(l,r))` did NOT reproduce the right's headers, and no revision or edit-script entry reported the change. Now, gated by **`DocxDiffSettings.CompareHeadersFooters` (default `true`)**:
  - **Story pairing per Word's model.** Each section's effective default/first/even header and footer stories pair across the two documents by (section ordinal × occurrence kind) with Word's previous-section inheritance rule — NOT by positional part order, which isn't stable across files. An inherited story referenced by several sections is diffed once.
  - **Native markup inside the parts.** A changed story is rebuilt with fine per-word `w:ins`/`w:del` (fields, tables, images included — media/hyperlink relationships import into the story's own part, since rels are part-scoped); a right-only story becomes a fresh part + `w:headerReference`/`w:footerReference` (with `w:titlePg`/`w:evenAndOddHeaders` ensured for first/even stories) with all-inserted content; a left-only story's content is marked deleted (the part + reference stay — accepting leaves an empty story, Word's own behavior). Revision ids stay unique across body/notes/stories. The round-trip contract — `accept ≡ right`, `reject ≡ left` at the per-block text level, with empty ≡ absent — now extends to header/footer scopes and is enforced corpus-wide by the renderer battery.
  - **Revisions + diff-as-data.** `GetRevisions` reports hdr/ftr-anchored revisions (`p:hdr1:…`) in `Fine` granularity, appended after note revisions; `WmlComparerCompatible` mode excludes them (that mode's contract is the legacy comparer's revision set, which has none). `GetEditScriptJson` carries an additive `headerFooterOps` array (omitted when no story changed — pre-existing scripts serialize byte-identically).
  - **Proof.** `DocxDiffHeaderFooterSmokeTests`: a synthetic pair (edited header, edited PAGE-field footer, untouched first-page header, added even footer) plus the real `WC004-Large`↔`-Mod` corpus pair (whose footer difference was previously silent), verified end-to-end including a headless-LibreOffice render oracle (`tools/diffharness/lo/lo_headerfooter_check.py`) proving an independent renderer surfaces the expected story text on the accept and reject views.
  - **v1 ceilings (documented).** Sections pair by ordinal; `w:sectPr`/settings visibility flags are ensured, not revision-tracked (no `w:sectPrChange`); unreferenced parts aren't compared (invisible in Word); **`Consolidate` does not merge header/footer scopes** (the merger forces the scope off — explicit and pinned, nothing silently dropped). The `DocxDiffCompatibility` catalog gains a `headersFooters` entry (`Covered`). Wire ripple: `compareHeadersFooters` in the Ops settings JSON + npm/python types (default-on needs no client change; the field enables opt-out).

### Changed
- **Upgraded from .NET 8.0 to .NET 10.0** (.NET 8 reaches end-of-support in November 2026). `TargetFramework` bumped to `net10.0` across the core library, tests, all CLI tools (`redline`/`docx2html`/`docx2oc`/`docxodus-pyhost`), the `diffharness` tool, and the WASM bridge; `global.json` now pins SDK `10.0.0`. The library, tests, and WASM/AOT toolchain all built and ran clean on the first pass (2468 passed / 0 failed / 3 skipped — identical to the .NET 8 baseline) with no source changes needed beyond the target framework bump. CI workflows (`ci.yml`, `playwright.yml`, `publish.yml`, `python-publish.yml`) now provision the .NET 10 SDK, and the WASM build script/Python host dev-fallback path were updated for the new `net10.0` output directory.

### Fixed
- **DocxDiff.Consolidate: a reviewer-inserted note that cites another note the same reviewer inserted no longer dangles (cross-kind note-in-note, N-way).** The N-way note-scope merge already renumbered cross-kind nested references (a footnote body citing an endnote, or vice versa) that live in BASE notes — but a reference nested inside a reviewer-INSERTED note's definition body was never rewritten from the reviewer's id space to the output id space (the body-reference rewrite only visits body clones, and the renumber sweep keys on the output-old id, which the reviewer id is not). So when a reviewer inserted a footnote whose text cited an endnote the same reviewer also inserted — whose target id becomes a *fresh* output id, not a base id — the nested reference kept the reviewer id and dangled on merge/accept/reject. `IrCompositeMarkupRenderer.ApplyCompositeNoteDiffs` now rewrites the nested references inside every reviewer-inserted note definition (all-`ins`, single-id-space content) through the same `outputId` map, before the body-order renumber carries them to their final ids. Proven by the new `IrCompositeCrossKindNoteTests` (base cross-kind nesting survives the N-way renumber across all three policies; reviewer-inserted footnote-cites-existing-endnote and endnote-cites-existing-footnote; and the previously-dangling reviewer-inserted-footnote-cites-reviewer-inserted-endnote), each asserting structural reference resolvability on merge/accept/reject with a non-vacuity guard, plus a headless-LibreOffice render backstop on the note-in-note-free note-merge path (note-in-note references are valid OOXML but LibreOffice Writer's DOCX import cannot load any document that contains one — cross-kind or same-kind — so the resolvability oracle is structural + Word, not LibreOffice). Closes the last untested corner (M-A #4) of the note-scope Consolidate merge.
- **DocxDiff: a paragraph move that crosses a table boundary spuriously relocated the TABLE, contesting the whole table block in an N-way `Consolidate` (issue #229).** A reorder can be ambiguous between relocating a light paragraph and a heavy structural block — `[A, table, B]` → `[A, B, table]` reads equally as "paragraph B moved up past the table" OR "the table moved down past B" (both cost one relocation and round-trip identically in a two-way `Compare`). `IrBlockAligner`'s LIS spine arbitrarily picked the reading that relocated the TABLE. In `Consolidate` that spurious table-move then collided with a second reviewer's DISJOINT table-cell edit and collapsed the whole table to a block-level conflict — the cell edit was surfaced in the conflict rather than composed per-cell (no data loss, and `reject ≡ base` always held; this was a precision/quality limitation, not a correctness failure). The aligner now applies a structural-anchor tie-break: among equal-length spines it keeps the most non-paragraph blocks (tables / section breaks / opaque blocks) anchored and relocates the lighter paragraph. It is implemented as a maximum-weight longest-increasing-subsequence (cardinality stays the primary key, so the relocation COUNT is unchanged) that fires ONLY when the plain patience-sort LIS relocated a structural block — the paragraph-only common path is byte-identical (zero churn). The paragraph move and the disjoint cell edit now compose independently (`conflicts == 0`, a native/lowered paragraph move + per-cell table compose, `reject ≡ base`), and two-way markup is cleaner (a paragraph moves, not a whole table). Base two-way parity holds at 179/179; full suite green. Proof: `IrCompositeTableMoveBoundaryTests` (all three conflict policies).
- **npm `DocxEditor`/worker bridge: the .NET 10 WASM runtime hung forever inside a dedicated Web Worker.** Upgrading the WASM bridge from the .NET 8 to the .NET 10 runtime surfaced [dotnet/runtime#114918](https://github.com/dotnet/runtime/issues/114918), an upstream regression where `dotnet.js`'s asset loader treats a truthy `globalThis.onmessage` as a (mis-detected) worker-environment signal and never resolves its asset-loading promises — `dotnet.create()`/`getAssemblyExports()` hang indefinitely with no error. `npm/src/docxodus.worker.ts` set its message handler via the `self.onmessage = ...` property form, which tripped exactly this bug once the WASM runtime moved to net10.0. Switched to `self.addEventListener("message", ...)`, which doesn't set the `onmessage` property and sidesteps the bug; behavior is otherwise identical. Confirmed via a byte-for-byte .NET 8 vs .NET 10 WASM bundle comparison — the .NET 8 build initialized a worker in ~485ms, the unpatched .NET 10 build hung indefinitely, and the patched .NET 10 build initializes in ~520ms. Affects every `npm` consumer of the worker bridge (`createWorkerDocxodus`, the `DocxEditor` worker mode); the main-thread (non-worker) WASM path was never affected.

### Added
- **DocxDiff.Consolidate: N-way merge made structurally complete — note edits, table structural changes, split/merge, and row moves now compose (or conflict loudly); a reviewer's edit is never silently dropped.** Closes the last engineering gate before the default-engine flip:
  - **N-way note-scope merge.** Reviewers' footnote/endnote edits against the shared base now consolidate (was: `NotSupportedException` for any N≥2 note edit; a single-reviewer consolidate silently omitted them). A base-matched note's blocks run the SAME per-block dispatch the body uses — disjoint note edits compose, identical ones reach consensus, contested ones become recorded `DocxDiffConflict`s resolved by the policy (whole-note delete vs edit included) — and reviewer-INSERTED notes land under fresh output ids. The composite renderer applies composed ops inside the footnotes/endnotes parts, rewrites reviewer-sourced body references from each reviewer's id space into the base-anchored output space (`IrCompositeScript.NoteIdMaps`), and runs the same body-order renumber + cross-kind nested-reference sweep as two-way `Compare`. Consolidated revisions cover note edits; the consolidate parity scoreboard's accept ≡ right metric now includes referenced note texts (all 84 corpus cases hold).
  - **Table column add/remove composes.** Per-cell composition pairs each reviewer's cell ops by BASE cell anchor (not position), so one reviewer's column change never shifts another's edits: an added cell renders with native `w:tcPr/w:cellIns` (kept on accept, removed on reject), a removed cell with `w:tcPr/w:cellDel` (removed on accept — absorbed into the preceding cell's `gridSpan`, Word's own semantics — restored on reject); a cell delete-vs-edit is a recorded conflict. (Was: any cell-count change → whole-table conflict fallback.)
  - **Uncontested split/merge/row-moves render natively.** The same sole-toucher eligibility that drives native move composition now covers `SplitBlock`/`MergeBlock` (native inserted/deleted-pilcrow markup, matching two-way `Compare`'s shape) and table `MovedRow` (the two-way del+ins row shape); colliding ones lower to del/ins and resolve through the existing conflict machinery. Also fixes a silent drop: a `MergeBlock` inside a multi-editor table cell reached neither grouping map and vanished.
  - **Cell-shell (`w:tcPr`) edits are visible and composable.** A width/gridSpan/vMerge/shading-only cell edit previously left every cell/row/table hash identical — classified `EqualBlock` and silently dropped from `Compare` AND `Consolidate` with zero conflict recorded. The whole `w:tcPr` now participates in the cell `ContentHash` (`IrCell.ShellDigest`); pairwise `Compare` surfaces such edits, and the composed table sources a changed cell's shell from its editing reviewer, reaches consensus on agreeing shells, and records a conflict for competing ones.

### Fixed
- **RevisionProcessor: rejecting a `w:sectPrChange` silently dropped the section's header/footer references.** The `w:sectPrChange` inner is `CT_SectPrBase` — the old section PROPERTIES only, with no `w:headerReference`/`w:footerReference` (those are outside the tracked property change). The reject transform replaced the whole `w:sectPr` with that reference-less inner, so rejecting a section-property change deleted the section's headers/footers. It now keeps the current references and restores only the properties (Word's own behavior). Analogous to the same-campaign `w:pPrChange` inline-`w:sectPr` fix below. Benefits any consumer that rejects a Word- or DocxDiff-authored `w:sectPrChange`.
- **RevisionProcessor: rejecting a `w:pPrChange` on a section-final paragraph dropped the inline `w:sectPr`.** The reject rebuild of the paragraph properties from the change marker carried the paragraph-mark `w:rPr` but not an inline `w:sectPr` (also outside `CT_PPrBase` / the tracked change), deleting the section break on reject. The inline `w:sectPr` is now carried over. See `docs/ooxml_corner_cases.md`.
- **DocxDiff: a note reference nested in the OPPOSITE note kind's body kept a stale id after renumbering.** An endnote reference inside a footnote definition (or vice versa) dangled after `Compare`'s body-order renumber pass, because each kind's pass swept only its own part for nested references. The old→new id maps are now applied in a single cross-kind sweep over both note parts. Benefits every `Compare()` call.
- **RevisionProcessor: accepting a deleted table cell (`w:cellDel`) NRE'd when the absorbing (preceding) cell had no `w:tcPr`.** The accept-deleted-cells transform now synthesizes the widened-`gridSpan` `w:tcPr` instead of dereferencing null.

### Added
- **DocxDiff: explicit, correct, transparent handling of inputs that already carry tracked changes — `revisionsInInput` lifted `Partial` → `Covered`.** Two parts:
  - **`DocxDiffSettings.PreAcceptInputRevisions` (default `false` → zero behavior change).** When set, every input is run through `RevisionProcessor.AcceptRevisions` BEFORE diffing, so the comparison — and the output package the markup renderer clones from those inputs — is revision-free on both sides. It is, by construction, exactly the `Compare(AcceptRevisions(left), AcceptRevisions(right))` wrapper made a first-class setting (byte-identical to it), applied uniformly to all seven entry points (`Compare`/`GetRevisions`/`GetEditScriptJson` + the four consolidate-family methods, accepting the base and every reviewer). With it on, every `w:ins`/`w:del`/`w:moveFrom` in the result is attributable to THIS diff (no stale input revision passed through), and the round-trip holds against the accepted view of each side in every scope (body, headers/footers, notes). **.NET-only in v1** (WASM/npm/python bridge ripple deferred).
  - **The default is now pinned + documented, not incidental.** Characterization tests (`Docxodus.Tests/Ir/Diff/RevisionsInInputDefaultTests.cs`) pin today's behavior: rule N13 means the engine already diffs the ACCEPTED VIEW of each input (a document's own revisions never surface as their own diff; the produced body is clean and round-trips to the accepted view), BUT the markup renderer clones the output on the LEFT package and only rebuilds the body (+ changed notes), so pre-existing revision markup in carried-over parts (headers/footers, unchanged footnotes/endnotes, styles, comments) leaks through verbatim — which also breaks the round-trip in those scopes (a leaked prior insertion is rejected by reject-all). The contract is documented in `docs/architecture/ir_diff_engine.md` and the two honest costs of accept-all (it flattens prior authorship/change boundaries; "accept all" is itself a policy that overrides a change a prior reviewer left unaccepted) in `docs/ooxml_corner_cases.md` + the inspector note. The `DocxDiffCompatibility` catalog entry `revisionsInInput` is flipped to `Covered`.
- **DocxDiff real-document fidelity guarantees made CI-durable + a client-callable accept/reject surface.** Three pieces, so a regression that still *parses* turns the suite red instead of slipping through:
  - **Vendored dense fixture + headless real-doc oracles.** `TestFiles/DD/DD001-DenseBookmarkXrefFootnote.docx` (built deterministically by `DocxDiffRealDocFixture`, regenerated by the Skip-by-default `__RegenerateVendoredFixture` fact) is a Series-A-style contract dense with bookmarked defined terms/sections, REF/PAGEREF/NOTEREF fields + an internal hyperlink anchor, footnotes (one citing another — a note-in-note reference), endnotes, and non-sequential footnote ids. `DocxDiffBookmarkRealDocTests` no longer depends on machine-local `~/Downloads` NVCA paths: the vendored fixture is a REQUIRED always-run case (the NVCA contracts remain optional local enrichment that soft-skips when absent), and the formerly-`~/Downloads`-only footnote/endnote oracles (unique ids, body+note-in-note resolvability, referenced-note-text round-trip) now run headlessly on every build. The fixture immediately earned its keep — it surfaced that the schema-validity oracle's pre-existing-error dedup mis-counted a *correctly* renumbered note-in-note reference as a new defect (the SDK validator false-positives on note-in-note refs and embeds the reference value in the message, which a legitimate 5→4 id compaction changes); the dedup is now part-scoped + value-normalized (see `docs/ooxml_corner_cases.md`).
  - **Gated semantic ratchet on the head-to-head harness.** `IrVsWmlComparerTests` is no longer totality-only: over the 184-comparison WC corpus it now enforces a Match-rate floor (96), a Divergent ceiling (66), per-cause ceilings on the genuine-fidelity-loss buckets (ScopeGapNewEmpty/OldEmpty/OpaqueGap = 0, Unclassified ≤ 2), an `OldError` ceiling (≤ 2), and a WmlComparer-compatible Match floor (150) — a ratchet that may only tighten. These are bucket-COUNT gates: an over/under-report that shifts a previously-matching pair drops Match below the floor (verified by perturbation). The one gap they leave — a *partial* under-report worsening WITHIN an already-Divergent case while it stays in the unceilinged TokenSpanGranularity bucket — is closed positively by content round-trips: `DocxDiffScenarioTests` (synthetic note edits) and the DD001 real-doc fixture's right side, which now edits footnote AND endnote CONTENT so the referenced-note-text round-trip exercises the note-diff path. The floors are coupled to the WC-corpus snapshot (documented inline for re-baselining).
  - **Accept/reject byte→byte surface + client round-trip tests.** A new `DocxDiff` accept/reject primitive (`DocxDiffOps.AcceptRevisions`/`RejectRevisions`, wrapping `RevisionProcessor`) is rippled to WASM (`DocxDiffBridge.AcceptRevisions`/`RejectRevisions`) + npm (`docxDiffAcceptRevisions`/`docxDiffRejectRevisions`) and the stdio host (`docx_diff_accept_revisions`/`docx_diff_reject_revisions`) + docx-scalpel (`docx_diff_accept_revisions`/`docx_diff_reject_revisions`). It lets clients verify the redline's round-trip contract — `accept(compare(left,right))` ≡ `right`, `reject` ≡ `left` at the per-block text level — not just inspect its shape. New tests assert that contract end-to-end through each client wire (`DocxDiffOpsRoundTripTests` in .NET, `test_docx_diff.py` in docx-scalpel, the accept/reject round-trip in `npm/tests/docx-diff.spec.ts`), so a wire/type-mapping break in a client diff path is caught.
- **DocxDiff compatibility inspector — a rudimentary pre-flight that warns when a document contains a construct DocxDiff has not yet had a fidelity campaign for, and whose status-tagged catalog doubles as the campaign roadmap.** `DocxDiffCompatibility.Inspect(doc|bytes)` (also `DocxDiff.InspectCompatibility(...)`) returns a `DocxDiffCompatibilityReport` of under-tested features present in the document (each with an occurrence count + note); `DocxDiffCompatibility.Catalog` lists every feature with its `DocxDiffCoverage` (Untested/Partial/Covered) independent of any document. By default it changes nothing about `Compare`/`GetRevisions`/`Consolidate`; the caller can run it standalone, or opt into an automatic pre-flight on the diff methods via two new `DocxDiffSettings` knobs — `OnCompatibilityWarning` (an `Action<DocxDiffCompatibilityReport>` callback, fired before diffing with both inputs' warnings combined) and `ThrowOnCompatibilityWarning` (throws `DocxDiffCompatibilityException`, carrying the report) — both default off, so the pre-flight scan runs only when one is engaged (zero cost otherwise). v1 covers content controls, math, DrawingML, textboxes, RTL/complex-script, OLE objects, complex fields, and pre-existing tracked changes; bookmarks/cross-references, footnotes/endnotes, and comments are marked `Covered` (comments lifted from `Partial` by the comment fidelity campaign, see Fixed). .NET-only (no WASM/npm/python ripple yet).
- **Editor — run font family, rule-above placement, a block-delete affordance, and grid-picker cell alignment — closing the affordance gaps the S-1 round-4 smoke test flagged.**
  - **Font family** — new `FormatOp.FontFamily` (maps to `w:rFonts` ascii/hAnsi/cs; `null` leaves it unchanged, `""` clears so the run inherits the style/default). The `w:rFonts` element is inserted in CT_RPr schema order (after an optional `w:rStyle`). Editor: `DocxEditor.setFontFamily(name)` (multi-block + last-selection cached like `setFontSize`); the demo ships a curated font dropdown (Calibri / Times New Roman / Arial / Georgia / Cambria / Courier New / Verdana / Garamond). Lets a run match a serif filing (a blank doc seeds Calibri). Rippled `DocxSession` → `DocxSessionJson` (`fontFamily` wire field) → `npm/src/types.ts` `FormatOp` → `editor.ts`. Tests: C# `DocxSessionS1FeaturesTests` DS220 (sets ascii/hAnsi/cs) / DS221 (schema order + `OpenXmlValidator` clean) / DS222 (`""` clears), browser `editor-fontfamily.spec.ts`.
  - **Rule above or below** — `DocxEditor.insertHorizontalRule(weight, style, position)` gains `position: "above" | "below"` (default `"below"`); `"above"` passes `Position.Before` (the bridge already accepted it). Closes the S-1 heavy top bar that sits between the filing table and "UNITED STATES" — previously rules could only go below. The demo ships an **Above | Below** toggle honored by all three rule buttons. Test `editor-rule-above.spec.ts`.
  - **Block delete** — `DocxEditor.deleteBlock()` removes the active block (e.g. a stray empty paragraph left by a table) via `DocxSession.DeleteBlock` + re-render, focusing the previous block. Inert inside a table (cells are removed via the table toolbar) and inert when it is the only editable block. The demo ships a 🗑 button. Test `editor-delete-block.spec.ts`.
  - **Grid-picker cell alignment** — the demo table grid picker no longer hard-codes centered cells; an **L / C / R** selector (default **left** = the document default) drives `insertTable(..., { cellAlignment })`, so an inserted table can be left-aligned (the S-1 filing line no longer comes out centered). Per-column-width / content seeding stay API-only (`insertTable(opts)` already supports them). Test `editor-demo-grid.spec.ts` (left → cells not centered).
- **Table row/column editing — `DocxSession.InsertTableRow` / `InsertTableColumn` / `DeleteTableRow` / `DeleteTableColumn`, addressed by a cell-paragraph anchor.** Previously a table's shape was fixed at insertion; now rows/columns can be added or removed after the fact. Insert clones the reference row/column's cell widths (and `w:tblGrid/w:gridCol` stays consistent) and starts the new cells empty, returning their cell-paragraph anchors; deleting the last row/column removes the whole table. v1 assumes a rectangular grid (no horizontal `w:gridSpan` merges). Rippled `DocxSession` → `DocxSessionOps` → WASM `DocxSessionBridge` → npm `DocxSession` (`insertTableRow/Column`, `deleteTableRow/Column`) + `DocxEditor` (`insertTableRow("above"|"below")`, `insertTableColumn("left"|"right")`, `deleteTableRow/Column`), and the `editor.html` demo gains a **floating table toolbar** (appears when the caret is in a cell). Tests: C# `DocxSessionTableEditTests` DT201–DT207 (shape + `OpenXmlValidator` schema-valid + whole-table removal + non-cell rejection) and browser `editor-table-edit.spec.ts` (insert/delete row+column reshape and survive save→reopen). _Follow-up:_ drag-to-resize column borders is not yet wired (column proportions are set via `ColumnWidths` at insert); tracked for a later pass.
- **`TableInsertOptions.ColumnWidths` — per-column table widths (twips) for unequal layouts.** `DocxSession.InsertTable` previously split the content width into equal columns; `ColumnWidths` (one positive width per column, left→right) now drives `w:tblGrid/w:gridCol` + each cell's `w:tcW`, sizing the table to their sum. A list whose length ≠ the column count is rejected (no silent equalize). Unblocks the S-1 filing-header row, where the long "As filed…" left line + short "Registration No. 333-" right line need a wide-left / narrow-right split that equal halves would wrap. Rippled `DocxSession` → `DocxSessionJson` (`columnWidths` wire field) → `DocxSessionBridge` (passthrough) → npm `TableInsertOptions` + `DocxEditor.insertTable`. Tests: C# `DocxSessionS1FeaturesTests` DS214 (widths land in grid + cells) / DS215 (wrong count rejected); browser `editor-table-colwidths.spec.ts` (unequal render survives save→reopen).
- **Editor — font size, paragraph borders / horizontal rules, table insertion, and a "New blank document" — enough surface to draft an SEC Form S-1 cover page from scratch.** Smoke-testing the `DocxEditor` against the SpaceX S-1 cover page found four missing capabilities; all are now first-class, lossless, and schema-valid (OOXML validator clean):
  - **Font size** — `FormatOp.FontSizePts` (points → `w:sz`/`w:szCs` half-points; `<= 0` clears). Editor: `DocxEditor.setFontSize(pts)`; demo ships a size dropdown. Drives the cover page's large "FORM S-1" and company-name lines.
  - **Paragraph borders / horizontal rules** — `ParagraphBorderEdge` record + `ParagraphFormatOp.{TopBorder,BottomBorder,ClearBorders}` (`w:pBdr`), and `DocxSession.InsertHorizontalRule(anchor, pos, edge?)` (an empty paragraph carrying a bottom border — what an S-1 section divider is; supports `single`/`double`/`thick` styles + weight). Editor: `DocxEditor.insertHorizontalRule(weight, style)`; demo ships thin + thick rule buttons.
  - **Table insertion** — `DocxSession.InsertTable(anchor, pos, rows, cols, TableInsertOptions{Borderless, CellContents (row-major markdown), CellAlignment})`; returns the created cell-paragraph anchors so each cell is then addressable to fill/format. Borderless emits explicit `w:val="none"` borders (the standard invisible layout table for multi-column rows: the registrant-facts row, the filing-header left/right line, the "With copies to:" counsel columns). Editor: `DocxEditor.insertTable(rows, cols, options?)`; demo ships an "insert table" button. Cell text stays editable via the existing per-block path.
  - **New blank document** — `DocxSession.CreateBlankDocxBytes()` (static) mints a complete blank DOCX (Normal style + doc defaults, settings, US-Letter portrait section) that opens cleanly in Word; surfaced as WASM/npm `createBlankDocx()` + `DocxEditor.openBlank(container, exports, options?)`; demo ships a "New" button. The drafting entry point for editors that start from scratch.
  - Rippled through every layer: `DocxSession` → `DocxSessionOps` → `DocxSessionJson` (hand-written wire parsers gain `fontSizePts`, the border-edge / table-options shapes) → WASM `DocxSessionBridge` (`CreateBlankDocx`/`InsertHorizontalRule`/`InsertTable`) → `npm/src/types.ts` + `session.ts` + `editor.ts` + the `editor.html` demo toolbar. New tests `DocxSessionS1FeaturesTests` (DS201–DS210), including DS210 which builds an S-1-style page with all four features and asserts the OOXML is schema-valid (`OpenXmlValidator`, zero errors). The full cover page was drafted end-to-end through the editing surface and renders faithfully (see `docs/architecture/s1_smoke_test_features.md`).
- **IR-powered DOCX editor foundation — faithful single-block HTML render + an in-browser block editor.** Lets a browser editor render a document, edit a block via `DocxSession`, and re-render *only that block* (~10 ms) instead of a full ~0.7–2.4 s re-conversion.
  - `WmlToHtmlConverterSettings.StampAnchors` — stamps `data-anchor="<unid>"` on block-level HTML elements (`p`, `h1`–`h6`, `li`, `table`) so DOM blocks are addressable by the shared `kind:scope:unid` anchor system.
  - `HtmlConversionOps.RenderBlockHtml(bytes | DocxSession | handle, anchorId, options)` — renders a single block to faithful HTML via a throwaway document that copies the source's styles/numbering/theme/font/settings parts. The session-attached overloads resolve against the live document (no byte re-open / whole-doc Unid pass) — measured 2.55× faster than the stateless path. The full-document render is the faithfulness oracle (per-anchor tag+text parity, test `HCO050`/`HCO052`).
  - WASM/npm surface: `DocumentConverter.RenderBlockHtml` + a `stampAnchors` parameter on `ConvertDocxToHtmlComplete`; npm `renderBlockHtml()`, `DocxSession.renderBlock()`, and `ConversionOptions.stampAnchors`.
  - `DocxEditor` (npm, `import { DocxEditor } from "docxodus"`) — a framework-agnostic, pure-TypeScript block editor: renders a faithful document with `data-anchor` blocks, makes projection-addressable paragraphs/headings `contenteditable`, and on commit edits via `DocxSession` then re-renders only the changed block. `{ paginated: true }` flows blocks into real page boxes (`pagination.ts`). Editing **preserves the block's inline formatting** (bold/italic/links) via `serializeInlineMarkdown`, not plain text. **Structural editing**: Enter splits a paragraph (`SplitParagraph`), Backspace at block start merges with the previous (`MergeParagraphs`). **Formatting controls**: `format(key)` (bold/italic/underline/strike/code/**superscript/subscript** on a selection via `ApplyFormat`), **`setAlignment`/`indent`/`pageBreakBefore`** (paragraph alignment, left-indent, page-break-before via the new `SetParagraphFormat`), `setParagraphStyle(styleId)`, `undo()`/`redo()`, `queryFormatState()`, plus Ctrl+B/I/U and Ctrl+Z/Ctrl+Shift+Z shortcuts; the demo ships a full formatting ribbon. New public DocxSession API: `FormatOp.VertAlign` (w:vertAlign), `SetParagraphFormat(anchor, ParagraphFormatOp{Alignment, IndentDelta, PageBreakBefore})` (w:jc/w:ind/w:pageBreakBefore), and `ApplyListFormat(anchor, ListFormat.None|Bullet|Decimal)` — promotes a plain paragraph to a bullet/numbered list item, synthesizing a reusable numbering definition via the new `Internal.NumberingFactory` (find-or-create by marker `w:nsid`). All surfaced through WASM/npm; editor adds `setAlignment`/`indent`/`pageBreakBefore`/`toggleList`. List items render with their marker glyph + hanging indent in both full and incremental (single-block) renders. (The editor defaults to inline styles so per-block re-renders stay self-contained.) Lossless `save()`. A runnable demo is at `npm/examples/editor.html` (`cd npm && npm run demo`). See `docs/architecture/ir_editor_feasibility.md` and the prioritized `docs/architecture/ir_editor_roadmap.md`.

### Added
- **Editor — `DocxEditor.setPaginated(boolean)` toggles continuous ↔ paginated rendering without losing edits.** Re-renders from the LIVE session (so every committed edit and the undo/redo history survive the toggle), instead of re-opening the original bytes. The `examples/editor.html` demo's Paginated checkbox now routes through it — previously toggling pagination silently discarded all session edits. Browser test `editor-gaps.spec.ts` GAP4.

### Performance
- **Editor — single-block render (`RenderBlockHtml`) is ~6.5× faster on documents with a large style gallery, removing the perceptible delay when editing/leaving table cells.** Profiling a python-docx-authored S-1 (styles.xml = 164 styles / ~434KB) found a keystroke-commit re-render cost ~650ms in WASM, dominated by two pure-overhead steps the incremental per-block path repeats every commit: **(1)** `MarkupSimplifier.SimplifyMarkup` re-walked the copied style-definition parts (≈70ms of ≈73ms of the convert; that pass only strips rsids, which never reach the HTML), and **(2)** `RenderResolvedBlock` deep-cloned + re-serialized the whole style gallery into a throwaway doc every render (≈26ms). Two fixes, both behind the existing `RenderBlockHtml` API (no WASM/npm/bridge change): a new internal `SkipFormattingPartsSimplification` flag (`WmlToHtmlConverterSettings`/`SimplifyMarkupSettings`) skips the rendering-neutral style-part simplification for the single-block path; and `DocxSession` now caches the throwaway "formatting shell" (the formatting parts + an empty body, serialized once) and reuses it across renders, rebuilding only when a cheap content signature of the style/numbering parts changes (i.e. only on a format op that adds a style/numbering/level — never on a text edit, so it survives typing). Measured: `RenderBlockHtml` 149ms→11ms (Debug, 13.5×); browser edit-commit ~650ms→~100ms. New tests: `HtmlConversionOpsTests.HCO053` (full convert with the flag on vs off is byte-identical — covers CSS classes + paginated, not just tag+text), `HCO054` (session shell-cached render is consistent across calls and byte-identical to the stateless path), `HCO055` (a mid-session `ApplyListFormat` rebuilds the shell so the marker renders — invalidation). Output is byte-for-byte unchanged; the full .NET + Playwright suites are green.

### Fixed
- **WmlComparer — redline output preserves the source document's headers and footers.** `Compare` rebuilds the final `w:sectPr` from a whitelist of section properties and previously dropped `w:headerReference`/`w:footerReference`, so the produced redline lost its headers/footers entirely (the parts were still carried in the package, just orphaned). The references are now kept — they always resolve, because the saved `sectPr` and the output package both derive from the left document. Note WmlComparer still does not *compare* header/footer content: the left ("before") document's headers/footers are carried as-is; tracked header/footer comparison is the DocxDiff engine's `CompareHeadersFooters`.
- **DocxDiff (IR diff engine) — comment fidelity campaign: edited commented paragraphs now get fine per-word markup (not a coarse whole-block bail), and comment id↔range↔reference↔definition + threaded-reply integrity is guaranteed end-to-end. Lifts the compatibility-inspector `comments` entry from `Partial` to `Covered`.** Found and fixed under the same layered oracle as the bookmark/footnote campaigns (OpenXmlValidator schema validity → a comment-structure round-trip → WmlComparer parity → headless-LibreOffice load+refresh) over a synthetic corpus AND a vendored comment-dense contract (`TestFiles/DD/DD002-DenseComments.docx`).
  - **Fine per-word markup on edited commented paragraphs.** A commented paragraph previously bailed to a whole-block `del(left)+ins(right)`; comment range markers + the `commentReference` run now ride through the token diff as `AlwaysKeep` zero-width markers (exactly like bookmarks), so editing a commented word yields per-word `w:ins`/`w:del` with the comment anchors intact, accept/reject each resolving to one comment.
  - **`NormalizeComments` — the comment analogue of `NormalizeBookmarks`.** Reconciles the rendered body so every `w:commentReference` resolves to exactly one `w:comment`, every `w:commentRangeStart` is unique and pairs 1:1 with a `w:commentRangeEnd`, and an unchanged comment survives BOTH accept and reject: (A) a common comment with a bare survivor collapses to a single bare range; (A2) a right-added / left-deleted comment's bare markers are wrapped in `w:ins`/`w:del` so they toggle with their side (no leak into the opposite resolution); (B) a wholly-rewritten comment's del/ins copies renumber the deleted copy to a fresh id + a cloned definition — the comment-dedup analogue of the bookmark renumber-collision; (C) orphaned markers are paired/dropped so the output is always schema-valid + fully resolvable.
  - **Right-side definitions + threading merged.** The output is built on the LEFT package, so a comment ADDED in the right document referenced dangling. `MergeRightCommentDefinitions` copies right-only `w:comment` definitions (creating the comments part if the left had none) plus their `commentsExtended` (`w15:commentEx` `paraIdParent` reply links) and `commentsIds` entries; the (B) dedup clone gets its OWN fresh `w14:paraId` + a cloned `commentsExtended` entry so a **reject-side threaded reply keeps its parent link** (the bug that motivated carrying threading through the clone).
  - **`Consolidate` reconciles comments too.** Because comment markers are now `AlwaysKeep`, they ride the N-way composite token diff as well — so `IrCompositeMarkupRenderer` now runs the same `MergeRightCommentDefinitions` (per reviewer) + `NormalizeComments` passes the two-way renderer does, keeping `DocxDiff.Consolidate` output over comment-bearing documents schema-valid + fully resolved (the inspector's `Covered` flag holds for both `Compare` and `Consolidate`). v1 limitation (documented in `docs/ooxml_corner_cases.md`): a cross-document comment-id / `w14:paraId` collision between genuinely INDEPENDENT documents (not two versions of one) is an attribution/threading gap, never a structural corruption — every reference still resolves to exactly one comment.
  - Tests: `DocxDiffCommentStructureTests` (9-shape synthetic corpus: multiple comments on one paragraph, overlapping ranges, a range spanning paragraphs with one end edited, an edited anchored word, a wholly-rewritten commented paragraph, a threaded reply with the anchor edited, a right-added comment, a comment on deleted text) + `DocxDiffCommentRealDocTests` (the vendored dense contract, schema + structure + comment-structure round-trip + a headless-LibreOffice comment oracle). A headless-LibreOffice comment oracle (`tools/diffharness/lo/lo_comment_check.py`) enumerates `Annotation` fields, confirms each anchors + every threaded reply names a loaded parent, and refreshes with zero dropped/orphaned comments. Corner case documented in `docs/ooxml_corner_cases.md`. The 179/0-deviation + 84 parity scoreboards and the full .NET suite stay green.
- **DocxDiff (IR diff engine) — bookmark / internal cross-reference fidelity: editing a bookmark-bearing or bookmark-referencing paragraph could drop a `w:bookmarkStart`/`w:bookmarkEnd`, duplicate a bookmark id, drop a `REF`/`PAGEREF` field, or silently shift body text — dangling every internal cross-reference that targets the bookmark (and, since legal contracts are dense with bookmark-backed references, directly user-visible). Found and fixed by a structural-fidelity campaign (schema validity + a bookmark id↔name↔reference round-trip + WmlComparer parity + a headless-LibreOffice oracle script) over a synthetic corpus AND the real NVCA Model COI (392 bookmarks / 82 REF fields) + SPA (192 bookmarks / 68 REF fields). Both contracts now round-trip fully content-clean and bookmark-structurally sound across every bookmark/reference edit.**
  - **Boundary-dropped bookmarks.** Bookmark range endpoints are zero-width SOURCE markers but NOT diff tokens (rule N3 strips them from the token stream), so the token-driven boundary-ownership flags were blind to them and an edit-boundary bookmark fell through the cracks — the slicer rebuilt the edited paragraph *without* it, orphaning the bookmark and dangling its `w:hyperlink @w:anchor` + `REF`/`PAGEREF`/`NOTEREF`/`HYPERLINK \l` references. Bookmark markers are now `AlwaysKeep` in the source slicer (never dropped at an op boundary).
  - **Duplicate bookmark ids.** A whole-block `del(left)+ins(right)` bail (or a both-sides edit) emitted the same bookmark id (and name) on both the `w:del` and `w:ins` side — schema-invalid (`Sem_UniqueAttributeValue`) and ambiguous to any cross-reference resolver. A new identity-aware `NormalizeBookmarks` pass reconciles every run-level bookmark: one present in BOTH sources collapses to a single BARE pair (survives accept AND reject); a wholly inserted/deleted one keeps its revision context; ids are made unique, every start↔end re-paired (a far endpoint dropped by a dense overlapping-bookmark layout is re-closed), and names preserved — so `reject ≡ left` / `accept ≡ right` at the bookmark-structure level and every reference still resolves. Bookmarks nested in opaque content (math `m:oMath`, drawings) are left untouched (they ride their host's content hash). Tests: `DocxDiffBookmarkStructureTests` (16-shape synthetic corpus: gapped ids, multi-bookmark paragraphs, multi-paragraph ranges, hyperlink+REF on one bookmark, a TOC, the renumber-collision shape) + `DocxDiffBookmarkRealDocTests` (the real NVCA COI/SPA, schema + structure + body-text round-trip).
  - **Dropped REF/PAGEREF fields.** Editing text *before* a `fldChar` field dropped the whole field — its plumbing (`w:fldChar`/`w:instrText`) is zero-width, not a diff token, so a begin/separate/end run clustered at the edit boundary was dropped exactly like a boundary bookmark. Field plumbing is now `AlwaysKeep`, and a new `NormalizeFields` pass re-homes each field's plumbing to the field's own revision context (bare for an unchanged or result-edited field, `w:del`/`w:ins` for a wholly deleted/inserted one). A `w:fldSimple` whole-deleted with its paragraph is also no longer left as a dangling empty field — it is expanded to `fldChar` run form so the entire field toggles (it is not a valid child of `w:del`).
  - **Dropped body text near `w:noBreakHyphen`/`w:softHyphen`/`w:sym`.** The IR reads these as a single text character (so the tokenizer counts them) but the slicer treated them as zero-width — an off-by-one desync that dropped an adjacent character on reject (the reject of a "Company‑Controlled Intellectual" run lost the "I"). The slicer now advances the char counter by one for them, staying aligned with the tokenizer.
  - **Numbering-merge schema order.** Copying a missing `w:num` into the output appended it after a trailing `w:numIdMacAtCleanup`, violating the `w:numbering` child order (`Sch_UnexpectedElementContentExpectingComplex`). `WmlComparer.CopyMissingNumberingFromOneDocToAnother` now inserts numbering children in schema order. (Shared with the blessed `WmlComparer`.)
  - The `tools/diffharness` `diff`/`diffall` reports now include a **bookmark structural column** (`bkmk-struct`: unique ids + 1:1 pairing + every internal reference resolving); a headless-LibreOffice bookmark/cross-reference oracle (`tools/diffharness/lo/lo_bookmark_check.py`) enumerates `com.sun.star.text.Bookmarks` + `GetReference` fields, refreshes, and asserts zero "Reference source not found". Corner cases documented in `docs/ooxml_corner_cases.md`.
- **DocxDiff (IR diff engine) — non-body scope fidelity: a note-in-note reference dangled when its target note was renumbered, and editing a commented paragraph orphaned the comment. Found by a structural-fidelity audit across endnotes/comments/note-in-note (driven by the same id↔reference↔text invariant + headless-LibreOffice backstop as the footnote campaign, on a real NVCA Certificate-of-Incorporation contract). Endnotes were already clean — they share the footnote renumber code, so the footnote fixes (below) carry over (re-verified at scale: the COI contract's 94 footnotes incl. a `continuationNotice` at id 1 round-trip with zero duplicate/unresolved ids).**
  - **A note-in-note reference (a footnote/endnote whose body cites another note) no longer dangles when its target is renumbered.** `RenumberNoteIds` walks only `w:body` references, so when a nested reference's target definition was *also* body-referenced (and thus renumbered, e.g. id 5 → 2), the nested reference kept the stale id and resolved to nothing — surviving accept and reject (the SDK validator does not resolve note-body references, so it never flagged it). The renumber pass now records each definition's old→new id and remaps same-kind references inside note bodies. Test `DocxDiffFootnoteRobustnessTests.NestedNoteReference_ToRenumberedNote_StaysResolvable`.
  - **Editing a commented paragraph no longer orphans the comment.** The IR drops `w:commentRangeStart`/`w:commentRangeEnd`/`w:commentReference` from paragraphs (rule N15 — recorded as `IrCommentStore` char-offset spans for the markdown projection), so the fine token-diff path rebuilt the edited paragraph *without* them. This interim fix bailed a commented edited paragraph to the conservative whole-block `del(left)+ins(right)` fallback; the **comment fidelity campaign** below supersedes it with fine per-word markup, lifting comments from `Partial` to `Covered`. (Test `DocxDiffFootnoteRobustnessTests.EditingCommentedParagraph_PreservesCommentAnchorsOnRoundTrip` still guards the round-trip on that paragraph.)
- **DocxDiff (IR diff engine) — footnote/endnote structural corruption when a note-bearing paragraph is edited: dropped references, duplicate definition ids, and an NVCA-scale reserved-id collision. Three root causes, found and fixed under a layered oracle (OpenXmlValidator schema validity → footnote id↔reference↔text round-trip → WmlComparer parity → headless-LibreOffice load), driven by a new in-process scenario invariant and the extended `tools/diffharness`.**
  - **A trailing zero-width inline (footnote/endnote reference, drawing, tab, …) is no longer dropped from a token-diff `ModifyBlock`.** `SourceRunModel.Slice` rebuilt an edited paragraph's runs by half-open char span `[start,end)`; a zero-width inline at the tail of an Equal/Insert/Delete op sits exactly at `end` and was excluded — so editing a word in a footnote-bearing paragraph silently lost the `w:footnoteReference` (and a *different*, unedited paragraph's reference then absorbed the renumber, producing a duplicate id). Slice now takes the owning op's first/last-token zero-width boundary flags (`ZeroWidthBoundaries`), so a boundary zero-width is claimed by exactly the op whose token range owns it — never dropped, never double-counted (the `alpha⟨tab⟩bravo→charlie` round-trip is the regression guard).
  - **A footnote whose reference was deleted but whose definition lingers is no longer emitted twice.** When an edit removed a body reference while the definition stayed in both stores, the reference-correspondence emitted a `(id,null)` delete *and* `AppendUnreferencedNotes` emitted a `(null,id)` insert → two `w:footnote` definitions sharing one id (one deleted, one inserted). `AppendUnreferencedNotes` now reconciles an orphan with an existing opposite-side surplus of the same id into a single matched pair (a footnote's identity is its `w:id`). The renumber pass also links a `w:del` reference to its *preserved* definition when no deleted-only definition remains, so a reference-deleted-but-preserved note stays resolvable on reject even when its id ≠ its reference ordinal (masked by the `{1,2}`-in-order synthetic corpus; exposed by gapped ids).
  - **A reserved boilerplate note at a POSITIVE id no longer collides with a renumbered real note.** Word's `continuationNotice` rides at `w:id="1"` in the NVCA contract; `RenumberNoteIds` keeps reserved notes verbatim but renumbered real notes from `1`, re-minting id 1 → a duplicate `w:id` on **every** edit (even body/format-only). The real-note counter now starts above the highest positive reserved id (`{-1,0}`-only docs are unchanged). This is why a real, footnote-dense contract — not just the synthetic corpus — is part of the oracle.
  - Tests: in-process `DocxDiffScenarioTests.Scenario_PreservesFootnoteStructure` (unique ids + every body reference resolves + id→ref→text `accept ≡ right`/`reject ≡ left`, over every edit×feature scenario — the former `KnownFootnoteIdBug` characterization test is removed and its five scenarios folded back into the schema-valid theory), `DocxDiffFootnoteRobustnessTests` (gapped ids; positive-id reserved note), `IrMarkupRendererTests.Render_modify_with_zero_width_inline_at_span_boundary_round_trips`. `tools/diffharness` `roundtrip.json` gains footnote/endnote structure assertions (unique ids + refs resolve), surfaced as a `nstruct` column in `diffall`; a headless-LibreOffice backstop (`tools/diffharness/lo/lo_footnote_check.py`) confirms every output loads with footnotes intact, no repair. The 179/39 parity scoreboards and the full .NET suite stay green.
- **DocxDiff (IR diff engine) — two defects found by a LibreOffice-parity verification campaign (methodical edit×feature sweep on a real NVCA contract; harness in `tools/diffharness/`, findings in `docs/architecture/docxdiff_libreoffice_findings.md`).**
  - **Header/footer parts are no longer duplicated into the output.** `Compare` clones content-equal blocks from the RIGHT document; a mid-document section-break paragraph carries an inner `w:sectPr` whose header/footer references then dragged the RIGHT's header/footer parts into the LEFT-based output as `P<guid>.xml` duplicates (26 vs 14 parts on the contract; both sides' header content present). The blessed WmlComparer oracle stays clean (14 parts). Header/footer scopes are deliberately not diffed, so the LEFT package's parts are authoritative and the cloned references already resolve there (same r:ids — both sides derive from one base): `MoveRelatedPartsToDestination` gained an opt-in `skipHeaderFooterReferences` (default false; legacy callers unchanged), passed `true` by `IrMarkupRenderer.ImportRightSourcedMedia`; media (drawings) still import. Test `IrMarkupRendererTests.Render_does_not_duplicate_header_parts_for_equal_section_break_block`.
  - **`GetRevisions` now reports a table column add/remove (was silently 0).** A column-count change bails the markup renderer to a whole-table `del(left)+ins(right)` fallback (round-trips; LibreOffice renders the same whole-table replace), but `IrRevisionRenderer` had no matching fallback, so `GetRevisions` returned 0 — diverging from the WmlComparer oracle (2) and hiding a change the markup tracks. `RenderModifyBlock` now detects the same unpaired-cell condition and emits a Deleted(left)+Inserted(right) pair. Test `DocxDiffTests.GetRevisions_TableColumnChange_ReportsWholeTableReplace`.
  - **Inserted content-control (`w:sdt`) text no longer leaks on reject.** A run inserted inside a `w:sdtContent` was emitted BARE (no `w:ins`), so `RejectRevisions` left it — `reject ≠ left`, a core-contract violation that silently retained content the user rejected. `IrMarkupRenderer.WrapRunLevel` wrapped a container's DIRECT children, but a `w:sdt`'s runs live nested under `w:sdtContent`; a new `WrapContainerChild` descends through it (`w:ins`/`w:del` is a valid `w:sdtContent` child) and wraps the runs. `GetRevisions` was already correct; only the markup renderer leaked. Found by the advanced-feature sweep. Test `DocxDiffTests.Compare_InsertedContentControl_RejectStripsTheInsertedText`.
- **DocxDiff (IR diff engine) — hardening pass from a code-inspection audit: reject-fidelity, fail-fast on silent loss, boundary validation, and a wire-reachability gap.**
  - **Deleted table column now round-trips on reject.** A deleted trailing column produced a left-surplus cell op that the markup renderer's per-cell path dropped (the `ci >= rightCells.Count` cutoff), so `RejectRevisions` did NOT restore the column. `RenderModifyRow` now bails to the whole-table `del(left)+ins(right)` fallback on any column-structure change (a cell op with a missing left or right anchor), which round-trips exactly (`reject ≡ left`, `accept ≡ right`) at the cost of coarser markup — the honest representation, since the per-cell renderer is column-count-stable in v1. Test `DocxDiffTests.Compare_DeletedTableColumn_RejectRestoresTheColumn`.
  - **Composite/Consolidate content-integrity invariants are now enforced at RUNTIME, not by `Debug.Assert`.** The token-span tiling, table tiling, no-un-lowered-structural-op, and StackAll single-anchor invariants in `IrCompositeMerger` (plus the `IrCompositeMarkupRenderer` `NoteOps`-empty tripwire) were `Debug.Assert`s — `[Conditional("DEBUG")]`, stripped from the Release build the library ships in *and* the Release build CI runs tests under (`dotnet test -c Release`). They are now a runtime `IrCompositeMerger.Invariant`/throw, so a tiling/totality violation fails fast instead of silently emitting a corrupt or lossy consolidated document.
  - **Multi-reviewer (N≥2) consolidate fails fast on a reviewer note edit instead of silently dropping it.** The merger does not yet build composite note-scope (footnote/endnote) ops, so a reviewer's note CONTENT edit was silently lost. An N≥2 consolidate where any reviewer edited a note now throws `NotSupportedException` (attributed to the reviewer). A single-reviewer consolidate is unchanged (it degrades to a body-level merge; use `DocxDiff.Compare` for full single-reviewer note fidelity — the documented single-reviewer body-parity corpus exercises exactly this shape). Test `DocxDiffConsolidateApiTests.Consolidate_multireviewer_note_edit_fails_loud_…`.
  - **Consolidate boundary null-hygiene.** The four N-way entry points NRE'd on a null reviewer element, a null reviewer `Document`, or a null `DocxDiffConsolidateSettings.Diff`; they now surface a clear, attributed `ArgumentException` (shared `ValidateConsolidateArgs`). Tests in `DocxDiffConsolidateApiTests`.
  - **`DocxDiffSettings` boundary validation + contract fixes.** An explicit `DateTimeForRevisions` is now validated (a non-parseable value throws `ArgumentException` instead of being stamped verbatim into `w:date`); an explicit (even empty) `WordSeparators` array is honored rather than silently reverting to the default set (only `null` falls back). Tests in `DocxDiffTests`.
  - **`Culture` is now reachable from the non-.NET layers.** `DocxDiffOps.ParseSettings` read no `culture` key, so the case-folding culture was inert from WASM/npm/python; it now parses a `culture` string (rippled to the npm/python settings types). Test `DocxDiffOpsConsolidateTests.ParseSettings_reads_culture_…`.
  - Plus internal-quality cleanups: the aligner's Merge-group lookup throws a descriptive error rather than a bare `KeyNotFoundException`; the hyperlink-suffix non-collision invariant is now explicitly test-pinned (`IrDiffTokenizerTests` — the suffix was already sentinel-framed; an audit "collision" finding was a false positive from an invisible control char); a dead `if (d > 0)` in the Myers backtrace removed; stale `DetectSplitMerge`/`MoveModifyBlock` comments corrected; `docs/architecture/ir_diff_engine.md` updated.
- **Editor — splitting/merging a bordered paragraph no longer leaves the new paragraph rendered inside the rule's border box (the rule's line drawn under the heading text).** A third S-1 smoke test found that even though the OOXML is correct (DS216 already strips the border on split, and LibreOffice renders the saved doc correctly), pressing Enter inside a horizontal rule and typing made the new paragraph *appear* bordered in the editor. Root cause was the **incremental render**, not the engine: the full render wraps a visibly-bordered paragraph in a `border-bottom <div>` (CreateBorderDivs), and `DocxEditor.splitAtCaret`/`mergeWithPrevious` did an in-place node swap that left the new (borderless) paragraph stranded *inside* that stale div. Both now force a whole-document remount when a border wrapper is involved (the treatment list edits already get), so the converter re-groups border boxes — the rule stays its own separator line and the new paragraph renders clean below it. Test: browser `editor-border-bleed.spec.ts` (Enter in a rule + typing → new paragraph outside any border box; rule's line remains).
- **Editor demo — applying a font size with Enter now takes a single Undo to revert (was two), and no longer logs a benign `addRange` console warning.** The size field bound `applyFontSize` to **both** `change` and keydown-`Enter`, so Enter fired it twice → two `ApplyFormat` undo snapshots (one size change needed two Ctrl+Z), and the second call re-selected a block the first had already swapped out (the `addRange(): the given range isn't in document` warning). Enter now commits via `blur()` (single `change`), and `DocxEditor`'s internal `selectRange` skips a detached range defensively. Test: `editor-demo-fontsize.spec.ts` (a size applied via Enter is reverted by one undo).
- **Editor demo — the floating table toolbar no longer overlaps the table's first row for a table near the top of the page.** It was pinned 40px above the table unconditionally; it now measures its own height and the sticky header, sitting just above the table when there's room and dropping just below the table's top edge otherwise, so it never covers the first cell row.
- **Editor — pressing Enter inside a horizontal rule no longer stacks a second rule / borders the next paragraph.** A second S-1 smoke test found that splitting an empty bottom-bordered paragraph (what an HR is) cloned its `w:pBdr` onto the new paragraph, so every Enter produced another rule and any text typed below sat on a border. `DocxSession.SplitParagraph` now drops `w:pBdr` from the new paragraph **only when the split paragraph is empty** (a pure rule); a bordered paragraph that has text still splits with the border on both halves (boxed-block behavior preserved). Tests: C# `DocxSessionS1FeaturesTests` DS216 (Enter on a rule → exactly one bordered paragraph) / DS217 (text-bordered split keeps the border on both halves).
- **Editor — paragraph borders are now clearable from the editor surface (`DocxEditor.clearParagraphBorders()` + a demo "clear rule" button).** The engine/wire already accepted `ParagraphFormatOp.ClearBorders`, but `DocxEditor` exposed no way to reach it — once a paragraph had an HR border it could not be removed through the UI. `clearParagraphBorders()` clears borders on the active block (or every block of a multi-block selection) and re-renders fully (a border change adds/removes the wrapping border `<div>`); `examples/editor.html` gains a `─✗` button. Browser test `editor-clear-borders.spec.ts`.
- **`DocxSession.InsertTable` now keeps a paragraph after the table, so content can follow an end-of-body table.** The smoke test found that inserting a table as the last body block left the document ending `</w:tbl></w:sectPr>` with no trailing paragraph — Word's convention is to keep a `w:p` after every table, and in the editor there was no block below the table to type in. `InsertTable` now appends an empty paragraph after the table whenever what follows it isn't already a paragraph (nothing, a `sectPr`, or another table); when a paragraph already follows, none is added. Tests: C# `DocxSessionS1FeaturesTests` DS218 (table at body end is followed by a `w:p`) / DS219 (no double-insert when a paragraph already follows).
- **Editor — inserting a table on an empty line no longer strands that empty paragraph above the table.** `DocxEditor.insertTable` inserted after the active block, so building a table from a blank line left the blank line stranded above it. When the caret is on an empty paragraph (outside a table), the table is now inserted *before* it, so the empty paragraph becomes the editable line **below** the table — no stray line above, a reachable line below. Non-empty blocks are unchanged (table inserted after). Browser test `editor-table-empty-source.spec.ts`.
- **Editor — the font-size combobox now sizes the selected sub-range, not the whole paragraph.** Because the size field must take focus to be typed in, clicking it blurred the contenteditable block and collapsed the selection, so `setFontSize` always fell back to whole-paragraph sizing. `DocxEditor` now caches the last real (non-collapsed) selection per block — refreshed whenever a selection sits in a block, cleared when a caret is collapsed inside a block so it never goes stale — and `setFontSize` uses it when the live selection has been stolen by a toolbar control. New regression test in `editor-demo-fontsize.spec.ts` (select "BIG" of "BIGsmall" → only "BIG" becomes 28pt).
- **Editor demo — the font-size control is now an editable combobox (any positive value), not a dropdown capped at 48pt.** The engine's `setFontSize` was always unbounded, but the demo `<select>` only offered presets up to 48, so the large display sizes typical of an S-1 cover ("FORM S-1", the company name) weren't reachable from the toolbar. It's now a numeric input with a preset `<datalist>` (8…96) that accepts any value (apply on change/Enter) and reflects the current selection's size. New regression test `editor-demo-fontsize.spec.ts` (typing 72 sets a ~96px / 72pt run).
- **Editor — formatting now applies across a multi-paragraph selection, not just the focused block.** A selection spanning several blocks applies `format` (bold/italic/underline/…), `setFontSize`, `setAlignment`/`indent`/`pageBreakBefore`, and `setParagraphStyle` to **every** block in range; previously only the active block changed, so centering the S-1's ~20 stacked heading/address lines was one-click-per-line. Inline ops apply to each block's slice of the selection (first block from the caret to its end, middle blocks whole, last block to the caret); paragraph ops apply per block. The spanned blocks are resolved with `Range.comparePoint` (robust to a selection boundary that normalizes onto a wrapper element, which `Range.intersectsNode` mishandles). Single-block behaviour is unchanged. New test `editor-multiblock-format.spec.ts` (select three paragraphs → all centered + bold, survives save→reopen).
- **Editor demo — table insertion is now a visual grid picker instead of a freetext `prompt()`.** The table button opened a `prompt("rows x cols")` dialog (typo-prone, no preview); it now opens a hover-to-pick rows×cols grid (Word/Google-Docs style, up to 8×10) with a borderless toggle — clicking a cell inserts that table at the caret. New regression test `editor-demo-grid.spec.ts` drives the real demo (`editor.html`, now served in the Playwright harness) and confirms picking 3×3 inserts a 3×3 table.
- **Editor demo — a "double rule" button now exposes the double-border horizontal rule (the S-1's signature top divider).** The engine + `DocxEditor.insertHorizontalRule(weight, style)` already supported `double`/`thick` border styles, but the demo's two rule buttons both hard-coded `single`, so a true double rule was unreachable from the toolbar. Added a `══` button (`insertHorizontalRule(12, "double")`). New regression test `editor-rule-style.spec.ts` locks the capability (a double rule renders with a `double` bottom border and survives save→reopen).
- **Editor — Enter inside a table cell now stacks a second paragraph WITHIN the cell instead of being inert.** Smoke-testing the S-1 cover page found that a cell could hold only one line: pressing Enter did nothing (`GAP3`), so the value-over-label rows (e.g. "Texas" over "(State or other jurisdiction…)") and the multi-line law-firm address columns were unbuildable from the toolbar. The engine already splits correctly inside a `w:tc` (the new `w:p` is a sibling within the same cell, so the table grid is unchanged) — `DocxEditor` now routes Enter-in-cell to that split and re-renders the two cell paragraphs in place, each independently formattable (bold value, smaller-italic label). Grid-changing keys (cross-cell Backspace-merge, Tab focus-jump) stay inert. New test `editor-cell-multiparagraph.spec.ts` (two stacked paragraphs survive a lossless save→reopen in the same cell); `editor-gaps.spec.ts` GAP3 updated to the new contract.
- **Editor — Shift+Enter now produces a real Word line break (`w:br`) instead of a literal newline that Word renders as a space.** Found while smoke-testing the S-1 cover page: a line break typed in a paragraph or table cell committed as a raw `\n` inside `<w:t>`, which the editor's own converter showed as two lines (via `white-space: pre-wrap`) but Word collapses onto one line — a silent WYSIWYG-vs-Word divergence. Two symmetric fixes: the `DocxSession` markdown parser (`MarkdownPayloadParser`) now maps an intra-paragraph newline (the canonical GFM hard break `"  \n"`, trailing spaces consumed) to a `w:br` run, mirroring `WmlToMarkdownConverter`'s existing read-side `w:br → "  \n"`; and `DocxEditor` handles Shift+Enter deterministically (native `insertLineBreak`) and serializes a `<br>`/embedded newline to `"  \n"`. A *blank* line still splits paragraphs (unchanged). New tests: C# `DocxSessionS1FeaturesTests` DS211–DS213 (hard break → `w:br`, round-trips through the projection as `"  \n"`, blank line still splits) and browser `editor-linebreak.spec.ts` (Shift+Enter survives save→reopen as a `<br>`). The block-content text-diff commit path treats `w:br` as zero-width (matching the session's run-text space), so split/span offsets are unaffected.
- **Converter crashed (`ArgumentNullException`) on a document containing a borderless table (`w:tblBorders`/cell borders with `w:val="none"` and no `w:sz`).** This is the standard way real filings lay out multi-column rows (e.g. an SEC Form S-1 cover page: the registrant-facts row and the "With copies to:" counsel block are borderless tables), so the whole conversion aborted on such documents. Both border-resolution paths — `FormattingAssembler.ResolveInsideBorder` (via `ProcessInnerBordersPerTblBorders`/`RollInDirectFormatting`) and `WmlToHtmlConverter.ResolveCellBorder` (via `AdjustTableBorders`) — special-cased only `w:val="nil"`, so a `"none"` border fell through to a `(int)/(decimal) border.Attribute(w:sz)` cast that threw on the absent (optional) `w:sz`. Both now read `w:sz` null-safe (missing = 0 width). Output for documents that DO carry `w:sz` is unchanged. New test `HtmlConversionOpsTests.HCO056` (a borderless 2-column table renders instead of crashing). Surfaced by drafting the SpaceX S-1 cover page in the editor.
- **Paginated render was BLANK for any document containing a hard page break (`w:br w:type="page"`).** In pagination mode the converter emitted the page-break (and column-break) marker as an EMPTY `<div class="page-break">`, which `XElement` serializes self-closing (`<div .../>`). A browser's HTML parser treats a self-closed non-void `<div/>` as an UNCLOSED tag, so every following sibling — including the visible `#pagination-container` — nested inside the `display:none` `#pagination-staging`; the page boxes rendered at 0×0 and the whole paginated view was invisible (continuous mode was unaffected). `WmlToHtmlConverter.ProcessBreak` now gives both break markers an empty text-node child so they serialize with an explicit `</div>` and stay siblings of staging. New tests: C# `HtmlConverterTests.HC008d_PageBreakMarker_PaginatedMode_NotSelfClosing` (serialized marker is not self-closing) and browser `editor-gaps.spec.ts` GAP1 (paginated render is visible, container not swallowed by staging). The single-block render and continuous mode are unchanged.
- **Editor — table-cell text is editable, but structural keys inside a cell are now inert (no table corruption).** Table-cell paragraphs are indexed by the projection, so they were already `contenteditable` and round-tripped text edits cleanly — but pressing Enter inside a cell split the cell paragraph, and Tab/Backspace could run structural ops the single-block model can't give whole-table context. `DocxEditor.onKeydown` now keeps Enter (split), Tab (list-nest / focus jump), and Backspace-at-start (cross-cell merge) inert inside a `<table>`; cell editing stays text + inline/paragraph formatting only. The demo hint and `editor.ts` comments are corrected (cells are not "read-only"). Browser test `editor-gaps.spec.ts` GAP3. Editor-only; no library/converter API change.
- **Editor — editing a block no longer flattens inline formatting the markdown subset can't express.** Committing a block edit re-serialized the block to the projection's markdown subset (bold/italic/links) and rebuilt every run via `ReplaceText`, so underline, strikethrough, color, font size/family, highlight, and super/subscript were dropped from an edited block. The commit path now computes the minimal changed text span (longest common prefix/suffix between the committed and current content text) and applies it through `DocxSession.ReplaceTextAtSpan`, which rewrites only the runs inside that span — every untouched run keeps its exact `rPr` (themed colors, half-point sizes, language, kerning) and typed text inherits the boundary run's formatting, like Word/contenteditable. Empty/whitespace-only paragraphs (whose rendered placeholder space doesn't line up with the session's empty run text) fall back to a `ReplaceText` rebuild, where there is no formatting to preserve. Editor-only: `ReplaceTextAtSpan` was already exposed on the WASM/npm bridge; `committedText` is standardized to content text (list markers + bidi marks excluded) so the diff offsets align with the session's run-text space. New browser tests `editor.spec.ts` "editing a block preserves underline on an untouched run", "M1: run formatting survives a later text edit", and "text typed adjacent to a formatted run inherits its formatting".
- **Editor — Enter, indent, and bold/italic/strike were dead on Google-Docs-exported documents (non-integer twips, invisible per-paragraph borders, explicit "off" run props).** Smoke-testing the `DocxEditor` against a Google-Docs-exported `.docx` (every paragraph carries floating-point `w:ind`/`w:sz`/`w:spacing`, an all-`nil` `w:pBdr`, and explicit `<w:b w:val="0"/>` run props) surfaced four distinct failures, three at the library layer and one in the editor:
  - **(1) Increase/decrease indent silently failed on every paragraph.** `DocxSession.SetParagraphFormat` read the existing left indent with a bare `(int?)ind.Attribute(W.left)` cast, which threw `FormatException` on a non-integer twip value like `w:left="12.996749877929688"` (as Google Docs emits) — the op returned `InternalError` and the editor swallowed it, so indent did nothing. Fixed by reading the current indent with the converter's tolerant `WordprocessingMLUtil.AttributeToTwips` (decimal → truncate, the same parse that lets the doc render) and writing back a clean integer. Test `DocxSessionTests.DS215_SetParagraphFormat_Indent_ToleratesNonIntegerExistingValue`.
  - **(2) Even after (1), the indent never showed visually.** `WmlToHtmlConverter.CreateBorderDivs` treats *any* paragraph with a `w:pBdr` as bordered: it wraps the paragraph in a `<div>`, moves the left indent onto the div, and forces the paragraph's own `margin-left` to 0. Google Docs stamps an all-`nil` (invisible) `w:pBdr` on every paragraph, so each one was wrapped and its indent relocated to a div the editor's single-block re-render never updates. Fixed with a new `HasVisibleBorder` guard so an all-`nil`/`none` `pBdr` is treated as no border (no wrapping div; the indent stays on the `<p>`). Test `DocxSessionTests.DS217_RenderBlock_InvisiblePBdr_DoesNotEatIndent`.
  - **(3) Bold/italic/strike were no-ops.** `DocxSession.ApplyFormat`'s `Toggle` only *added* a `<w:b>`/`<w:i>`/`<w:strike>` when none existed; a run already carrying an explicit `<w:b w:val="0"/>` (Google Docs stamps these) kept its `w:val="0"`, so turning the property on did nothing. Fixed to normalize an existing element to ON by dropping its `w:val` (bare `<w:b/>` = on). Test `DocxSessionTests.DS216_ApplyFormat_Bold_TurnsOnExplicitlyOffRun`.
  - **(4) Enter at end-of-line was silently dropped on directional paragraphs.** The converter wraps run text in bidi marks (`U+200E`/`U+200F`) that are NOT in the session's run text, so the editor's caret-offset math counted them and a caret at end-of-line mapped past the paragraph's length → `SplitParagraph` returned `OffsetOutOfRange` and the keystroke was lost (mid-paragraph splits were off-by-one). Fixed in `npm/src/editor.ts` by excluding bidi formatting marks from the content-offset space (new `stripBidi`/`domOffsetForContentOffset`, applied in `contentOffsetOf`/`blockContentText`/`placeCaretAtOffset`/`selectRange`), the same way generated list markers are already excluded. New browser test `editor.spec.ts` "Enter at end-of-line works when block text carries bidi marks". Editor-only.
- **`DocxSession.SplitParagraph` — Enter over-propagated the source paragraph's properties to the new paragraph, breaking the editor's "draft from scratch" flow.** Splitting cloned the *entire* `w:pPr` into the new paragraph, so four things leaked across an Enter: **(A) paragraph style** — pressing Enter after a `Title`/`Heading` produced another `Title`/`Heading` instead of the style's linked `w:next` (Normal), so every block under a letterhead inherited the 28pt Title and had to be re-set to Normal by hand; **(B) baked bold** — because the new paragraph stayed a heading, text typed into it then converted to Normal kept the heading style's bold as a direct run property; **(C) `pageBreakBefore`** — a page-broken paragraph propagated its break to every paragraph split off it (e.g. a page-broken "Enclosures" heading pushed each list item onto its own page); **(D) duplicate Unids** — the cloned `w:pPr` children kept the source's `PtOpenXml:Unid`, landing the same id on multiple elements. Fixed in `SplitParagraph`: an empty Enter-at-end split of a *non-list* paragraph now rebases the new paragraph onto the style's `w:next` (new `ResolveNextParagraphStyle` reads `w:style/w:next` from the styles part) with a clean `pPr` — dropping the heading-only direct props and the inherited paragraph-mark `rPr` (so freshly-typed text isn't bold); `pageBreakBefore` is always stripped from the new paragraph; cloned property Unids are re-minted; and the new paragraph's anchor kind is resolved from the fresh projection (a Heading→Normal rebase flips `h`→`p`). List items are exempt so the editor's Enter-continuation keeps the list, and mid-text splits keep the style (a same-paragraph continuation). New tests `DocxSessionTests.DS046_SplitAfterStyledParagraph_AppliesNextStyle`, `DS046b_SplitStyledParagraphMidText_KeepsStyle`, `DS047_SplitDoesNotPropagatePageBreakBefore`, `DS048_SplitListItemContinuesList`, `DS049_SplitReMintsClonedPropertyUnids`. Library-internal; `SplitParagraph`'s signature is unchanged, so the WASM/npm editor picks the fix up transparently.
- **`DocxSession.ApplyFormat` — inline code (`FormatOp.Code`) was a visual no-op on documents that don't define a "Code" style.** The op stamps the run with `w:rStyle w:val="Code"`, but on a document whose style definitions never declared a "Code" style (e.g. most real-world DOCX, including the `HC031` sample) that reference is a phantom — Word and the converter silently render the run as plain text, so the editor's `</>` ribbon button appeared to do nothing even though the run was correctly split out. Fixed by a find-or-create pass: when `op.Code is true`, `ApplyFormat` now ensures a real **character** style with id `Code` exists (new `Internal.StyleFactory.EnsureCodeCharacterStyle`, mirroring `NumberingFactory`), synthesized with a monospace font (`Consolas`) so the run actually renders as code. An existing "Code" style of any kind is left untouched (the document's own definition wins); the styles part is flushed via `PutXDocument` since the session's `Save` only persists projected parts. New test `DocxSessionTests.DS214_ApplyFormat_Code_CreatesMissingCodeCharacterStyle` (no "Code" style → apply code → save/reopen asserts a persisted character style with a Consolas run font). Library-internal; `ApplyFormat`'s signature is unchanged, so the WASM/npm bridge picks the fix up transparently.
- **Editor — nested lists: indent/Tab on a list item now changes the list LEVEL (nesting) instead of just the margin.** Indenting a numbered/bulleted item (the ribbon indent button, or Tab / Shift+Tab) shifted the paragraph's left margin via `SetParagraphFormat` but left `ilvl` unchanged, so numbering stayed flat (1, 2, 3, …) while items merely moved sideways — "nested lists no bueno". `DocxEditor.indent()` now detects a list item and routes to `DocxSession.SetListLevel(±1)` (the op already existed and was exposed in the bridge — the editor just wasn't calling it), and **Tab / Shift+Tab** on a list item nest / un-nest it. Numbering then nests correctly (e.g. `1, 2, [sub-level 1], 3`) with the deeper indent, and Shift+Tab restores the flat sequence. New browser test `editor.spec.ts` Mlists4. Editor-only; no library API change (`SetListLevel` is unchanged).
- **Editor / `DocxSession.SetListLevel` — nesting a SOURCE-document list (numbering inherited from a paragraph style, abstractNum defining only level 0) was a silent no-op.** The previous fix wired Tab→`SetListLevel`, but that worked only on lists with a direct `w:numPr` and a multi-level abstractNum — i.e. lists *created in the editor*. Real-world docs (notably python-docx's default "List Bullet"/"List Number") attach numbering via the **style** (no direct `numPr`) and define **only level 0**, so Tab did nothing: `SetListLevel` rejected the paragraph ("no numPr"), and even past that the converter had no level to render. Three coordinated fixes: **(1)** `SetListLevel` now resolves the effective `numId`/`ilvl` from the pStyle chain and **materializes a direct `w:numPr`** at the new level when the item is style-inherited (what Word does when you Tab a styled list item), flushing the body part immediately (`MainDocumentPart.PutXDocument`) so the edit survives `Save` under WASM (where the typed-DOM/XDocument divergence otherwise drops it). **(2)** new `NumberingFactory.EnsureLevelDefined` synthesizes the missing `w:lvl` definitions (bullet glyph cycle •/o/▪ or decimal, matching the existing format) up to the target level AND upgrades `multiLevelType` off `singleLevel` (which `ListItemRetriever` honors by force-flattening to ilvl 0). **(3)** `ListItemRetriever` no longer treats a nested **bullet** as a numbered-list "continuation" (the heuristic that renders `4.` instead of `3.4` for numbered runs wrongly collapsed a first nested bullet back to its parent's glyph). A source-list bullet now nests visibly (• → o, 0.25in → 1in) and Shift+Tab restores it. New test `DocxSessionTests.DS054b` (materialized `numPr`, synthesized level, `multiLevelType` upgrade, and end-to-end converter render of the level-1 glyph). Full .NET + browser suites green (no numbered-list-continuation regression).
- **Editor — the block-commit / split / merge DOM swap could throw an uncaught `NotFoundError` ("node … no longer a child … moved in a blur event handler").** `commitBlock`/`splitAtCaret`/`mergeWithPrevious`/`swapBlock` guarded their `el.replaceWith(...)` with `el.isConnected`, but a synchronous `blur` fired during a focus transfer can detach `el` between the check and the call — which `isConnected` doesn't catch. Not reproducible via normal click/type/rapid-click (0 errors), but automated drivers that move focus / set selection via script, and the paginated duplicate-anchor case, threw repeatedly. Centralized all four sites through a single `replaceNode(oldEl, …newNodes)` helper that re-checks `parentNode` inside the re-entrancy guard and tolerates the race in a `try/catch` (silent — the session is already updated, so a skipped visual swap leaves correct content and the next commit/remount reconciles). New browser test `editor-gaps.spec.ts` GAP5 (scripted focus/selection churn in both render modes throws nothing). Editor-only.
- **Editor — paginated mode duplicated every `data-anchor` (hidden staging copy + visible page-box copy), leaving `document.querySelector('[data-anchor]')` ambiguous and a stale staging copy that a future reflow could revert edits from.** `pagination.ts` measures a hidden `#pagination-staging` subtree once, then flows *clones* into the page boxes; the editor left staging in the live DOM. `DocxEditor.mountPaginated` now removes the `#pagination-staging` subtree after the (one-shot) measurement pass, so the page-box copies are the single source of truth for `data-anchor` (a remount rebuilds staging fresh from the live session). New browser test `editor-gaps.spec.ts` GAP6 (paginated render has unique anchors, no leftover staging, sample anchor resolves to one page-box element). Editor-only.
- **WASM — `RenderBlockHtml`'s error handler crashed with `JsonSerializerIsReflectionDisabled`, masking the real failure as an uncaught error.** The bridge's catch handler serialized an anonymous type (`new { error = ex.Message }`), but the trimmed WASM build disables reflection-based `System.Text.Json`, so the handler *itself* threw — surfacing as a bare `Uncaught Error: JsonSerializerIsReflectionDisabled` in the console and hiding whatever `RenderBlockHtml` actually failed on. Fixed by building the error JSON reflection-free with `JsonEncodedText.Encode` (`{"error":"…"}`), matching the documented contract (rendered HTML starts with `<`, errors are a JSON object) so the editor degrades gracefully (skips the swap, keeps the typed DOM) instead of crashing. WASM bridge only; no public API change.
- **Editor — couldn't type into a numbered bullet after clicking to it from another bullet.** Committing a list item happens on blur, and the commit re-rendered (replaced) the item's DOM node. When the user clicked straight from one bullet to another, that node replacement ran *during* the blur — which cancels the browser's in-flight focus transfer, so focus fell to `<body>` and typing into the next bullet did nothing (a fresh, empty item was the common case). Fixed by NOT re-rendering a list item on a text commit: a plain text edit never changes the item's number, and the DOM already shows exactly what the user typed with the correct marker, so the commit now only syncs the session (`ReplaceText`) + bookkeeping and leaves the node in place. Plain (non-list) blocks still re-render in place for canonical HTML (verified focus stays on the newly-clicked block). New browser test `editor.spec.ts` Mlists3 drives REAL mouse clicks + REAL keyboard across three numbered items (one with text, two empty) and asserts focus lands on each clicked bullet, typing into each works, numbering stays 1/2/3, and save round-trips. Editor-only; no library/converter API change.
- **Editor — numbered lists didn't continue and Enter at the end of a list item did nothing.** Two issues found in manual testing of the `DocxEditor` block editor. **(1) Numbering didn't continue** (every numbered item rendered "1."): the editor re-rendered an edited list item *in isolation* (single-block render has no whole-document numbering context) and, on remount, failed to re-wire list items because the session's persisted unids diverged from a re-derived scheme. Fixed by opening the editor session with `persistAnchorIds: true` (so anchors are stable across re-render) and routing any list-affecting edit through a **full remount** rather than a single-block swap, so the converter assigns continuing numbers (1., 2., 3.) with real document context. **(2) Enter at end-of-line of a numbered item added nothing**: the generated list marker renders as a number/bullet run *plus a suffix tab*; `ConvertRun` tagged the number run with `data-list-marker` but the **suffix tab** (rendered via the tab-width path in `TransformElementsPrecedingTab`, not `ConvertRun`) was untagged, so its character inflated the caret offset past the paragraph's text length — `SplitParagraph` returned `OffsetOutOfRange` and the keystroke was silently dropped. Fixed by tagging the marker **wrapper span** (which contains both the number/bullet glyph and the suffix tab) with `data-list-marker` when the tab run carries `PtOpenXml.ListItemRun`, so the editor's caret/offset math (`isInMarker`) excludes the whole marker. The editor now makes both number and tab non-editable and excludes them from offsets; Enter at EOL of a numbered item adds a continuing item. New browser test `editor.spec.ts` Mlists2 (numbered continuation + Enter adds a continuing item). Editor-only feature surface; no change to the stable converter/library public API.
- **Consolidate — the lowered move-source `DeleteBlock` leaked `moveGroupId`/`isMoveSource` into the public composite edit-script JSON (B5 follow-up).** *Internal/experimental.* The B5 lowering pass retains `MoveGroupId`/`IsMoveSource` on a lowered move-source `DeleteBlock` so `IrCompositeMerger.MergeOneBaseBlock`'s contested-relocation detection can tell a relocation-delete from a plain removal. But `IrCompositeScriptJson.WriteEditOpBody` emitted `moveGroupId`/`isMoveSource` UNCONDITIONALLY, so `DocxDiff.GetConsolidatedEditScriptJson` produced a `DeleteBlock` carrying move fields — violating the documented `IrEditOp` field-presence contract (a `DeleteBlock`/`InsertBlock` carries NULL move fields) and misleading machine consumers of the edit-script-as-data (the public differentiator). Fixed by keeping the marker ONLY for the merger's internal detection (read off the GROUPED pre-emit ops) and STRIPPING it (`MoveGroupId`/`IsMoveSource` → null) the moment the lowered op is wrapped into the emitted `IrCompositeOp` (a new `IrCompositeMerger.EmitOp` cleaner routed through every block-level emit site — single-reviewer, consensus, contested-relocation, block-conflict FirstReviewerWins/StackAll, and the preceding-anchor insert path). As belt-and-suspenders, `WriteEditOpBody` now gates the move fields to the `MoveBlock`/`MoveModifyBlock` kinds only, so no future lowering can leak them. The contested-relocation conflict's competitors also previously got EMPTY `ResultText` (the lowered move-source `DeleteBlock` has no `RightAnchor`, so `BlockResultText` returned `""`), leaving the recorded conflict unable to say WHICH block is contested; a new `ContestedBlockText` sources the competitor text from the contested block's LEFT (base) content so the conflict identifies the relocated block. New probe `IrCompositeJsonTests.Consolidated_json_delete_insert_ops_carry_no_move_fields` asserts no `moveGroupId`/`isMoveSource` on any `DeleteBlock`/`InsertBlock` in the public JSON over a move scenario; new committed tests `IrCompositeFixTests.Three_reviewers_move_same_block_records_conflict_no_loss` (3-way contested relocation: no loss + placement conflict recorded with LEFT-content competitor text + reject≡base under all policies) and `MoveModify_single_reviewer_lands_moved_and_edited_content` (move-and-edit lands the moved+edited content + removes the original + reject≡base); the existing B5 move/merge/split tests are parameterized over all three policies. New fuzz coverage `CompositeFuzzTests.Composite_with_structural_ops_round_trips` (3/4/5 reviewers × 60 seeds) extends the composite fuzzer with a structural-op reviewer pool (`DiffFuzzer.GenerateCompositeWithStructuralOps` / `keepStructuralOps`) — Relocate/Split/Merge mutations — so the lowering + contested-relocation branch get real round-trip (reject≡base) + apply-verifier (no content loss on accept) coverage a future refactor can't silently break. No public API surface; the 179/179 GetRevisions + markup-corpus parity floors and the consolidate/composite suites are unchanged.
- **`DocxSession.ReplaceText` silently dropped a paragraph's footnote/endnote reference (B3).** Editing a paragraph that contained a `w:footnoteReference`/`w:endnoteReference` discarded the reference run along with the replaced text, orphaning the note definition. `ReplaceText` already preserves zero-width, semantically-significant inline markers (bookmark/comment/permission ranges) across a whole-block replace, but the preserved set covered only *bare-child* markers — a note reference lives inside a `<w:r>`, so it slipped through and was deleted. (This surfaced via the consolidate `footnotes-survive` battery, where two reviewers built with `ReplaceText` lost the base footnote on accept; diagnosis confirmed the base-engine `DocxDiff` 2-way diff round-trips a note-bearing edit cleanly — accept ≡ right, reject ≡ left, footnote survives — so the bug was in `DocxSession.ReplaceText`, not the diff engine.) Fixed by recognizing note-reference-only runs (a `<w:r>` whose only non-`w:rPr` content is a footnote/endnote reference) as preserved markers: the accept path re-attaches them around the replacement, and the tracked path keeps them out of the `w:del`/`w:ins` so they survive on both accept and reject, in their original leading/trailing position. New tests `DocxSessionTests.DS314_ReplaceText_PreservesFootnoteReference_AcceptMode`, `DS315_…_TrackedMode` (accept ≡ edited, reject ≡ base, ref survives both), and `DS316_ReplaceText_PreservesEndnoteReference_AcceptMode`. No public API surface; the 179/179 parity floors and consolidate/composite suites are unchanged.
- **Markup round-trip — adjacent DISTINCT same-target `w:hyperlink`s collapsed into one when one was edited (B1 follow-up; base-engine renderer regression).** *Internal/experimental.* The F2 `CoalesceAdjacentHyperlinks` post-pass rejoined adjacent emitted hyperlink fragments by ATTRIBUTE EQUALITY (`SameHyperlinkShell`) gated on "carries a plain run", which could not distinguish N fragments of ONE source `w:hyperlink` split by an internal edit (MUST rejoin) from N genuinely-DISTINCT adjacent source links that happen to share a target (MUST stay separate). With two authored same-target links `first`/`second` where the right edits the second (`second`→`SECOND`), all three emitted fragments folded into ONE link, so reject yielded ONE `firstsecond` link while the left had TWO — `RejectRevisions ≢ left` at the ContentHash level (IrReader frames each hyperlink boundary). Fixed by coalescing on SOURCE-WRAPPER IDENTITY instead of attribute equality: `SourceRunModel` now numbers each top-level source `w:hyperlink` with a document-order ordinal (the LEFT and RIGHT models number their Nth link identically, so an intra-anchor edit's Equal/del/ins pieces share the ordinal while distinct links get distinct ordinals), `Slice` stamps the ordinal onto each emitted wrapper as a transient `pt:SourceLinkId` marker, and the coalescer groups adjacent wrappers ONLY when they share that ordinal (stripping the marker before output). The pure-`w:del`-link + pure-`w:ins`-link gate is retained so the WC019 whole-anchor RETARGET (same ordinal, different target, no plain run) still stays two links. New tests `IrMarkupRendererTests.Adjacent_distinct_same_target_hyperlinks_one_edited_round_trips`, `…_unchanged_stay_separate`, and `Hyperlink_single_token_anchor_no_equal_run_replaced_round_trips`. No public API surface; WC019, the three F2 hyperlink tests, and the GetRevisions 179/179 + markup-corpus parity floors are unchanged.
- **Markup round-trip — editing text INSIDE a `w:hyperlink` anchor broke accept/reject (B1; base-engine renderer).** *Internal/experimental.* `IrMarkupRenderer.SourceRunModel` modeled a whole `w:hyperlink` as one ATOMIC segment and re-emitted the entire link element for EVERY overlapping token-op slice. An intra-anchor edit (e.g. a multi-word anchor "our website" with one word changed) produces several token ops over the same hyperlink span, so the link was emitted once per op — doubling/tripling the anchor, so `RejectRevisions(Compare(left,right)) ≠ left` and `AcceptRevisions ≠ right` for any hyperlink-internal edit. Fixed by (1) RECURSING into a `w:hyperlink`'s run-level children as sub-segments tagged with their owning container chain, so `Slice` reconstructs the hyperlink wrapper EXACTLY ONCE per contiguous run group it contributes (intra-anchor del/ins now land on the runs INSIDE the link); (2) claim-tracking the remaining atomic containers (`w:sdt`/`w:smartTag`/`w:ins`/`w:del`) so they too emit once across overlapping ops; (3) a `CoalesceAdjacentHyperlinks` post-pass that re-joins the per-op wrapper fragments of ONE source link into a single `w:hyperlink` — GATED so it never merges a pure-`w:del`-link followed by a pure-`w:ins`-link (the WC019 whole-anchor RETARGET shape: text AND href both change → different targets, the new id remapped post-assembly), which must stay two links so the remap + empty-shell-drop restore each side. The FormatChanged path descends into the wrapper to stamp `w:rPrChange` on the inner runs. New tests `IrMarkupRendererTests.Hyperlink_internal_text_edit_round_trips_2way` (multi-run anchor), `…_single_run_anchor_edit_round_trips_2way`, and `…_present_edit_outside_anchor_round_trips_2way` (control). No public API surface; WC019 (single-run whole-anchor retarget) and the GetRevisions 179/179 + markup-corpus parity floors are unchanged.
- **Consolidate — a reviewer's MOVE / MERGE structural ops were silently dropped, losing content on accept (B5).** *Internal/experimental.* The `IrCompositeMerger` (N-way consolidate) groups each reviewer's pairwise edit-script ops by base anchor (`GroupByBaseAnchor`, keyed on `op.LeftAnchor`) and routes right-only inserts by preceding anchor (`GroupInsertsByPrecedingAnchor`, only for `InsertBlock`). A structural op with a NULL left anchor reached NEITHER path and was silently DROPPED: a **move DESTINATION** (`MoveBlock`/`MoveModifyBlock`, `IsMoveSource=false`, `RightAnchor` set, `LeftAnchor` null) → the relocated paragraph VANISHED on accept; a **`MergeBlock`** (`RightAnchor` set, `SplitMergeAnchors` = left anchors, `LeftAnchor` null) → the merge was ignored and its consumed base blocks passed through as `EqualBlock`, so the merge was lost. In both cases `conflictCount` was 0 and `reject ≡ base` still held (no corruption — pure ACCEPT-side content loss). Fixed by a **lowering pass** (`IrCompositeMerger.LowerStructuralOps`) that rewrites every reviewer Move/Split/Merge op to equivalent `Insert`/`Delete`/`Modify` ops — PRESERVING op order (the preceding-anchor insert routing depends on it) — BEFORE grouping, so the existing Equal/Modify/FormatOnly/Insert/Delete composition handles them with no loss: a move SOURCE → `DeleteBlock` (retaining `MoveGroupId`/`IsMoveSource` as a relocation marker), a move DEST → `InsertBlock`, a `SplitBlock` → `DeleteBlock` + N ordered `InsertBlock`s, a `MergeBlock` → N ordered `DeleteBlock`s + an `InsertBlock`. The lowered `InsertBlock`s carry the contributing reviewer's `SourceReviewer`, so the renderer sources runs from the right reviewer. Two reviewers relocating the SAME base block to different places now collide on the lowered source-delete → a recorded **placement conflict** (the consensus removal is emitted once so `reject ≡ base`; each reviewer's relocating insert is routed independently so every destination survives accept — no loss; the conflict is deliberately NOT policy-resolved into the op stream, so a `BaseWins` flip never wrongly restores a block both reviewers removed). **v1 limitation (documented):** in a consolidate, a reviewer's moves/splits/merges are lowered to insert/delete — the consolidated output shows them as `w:del`/`w:ins` rather than native `w:moveFrom`/`w:moveTo` or split/merge markup; content is fully preserved and round-trips. Native cross-reviewer move/split/merge composition is a follow-on. New tests `IrCompositeFixTests` (move-disjoint accept preserves the moved paragraph; two reviewers relocating the same block lose no paragraph + record a conflict under all three policies; merge-vs-edit preserves the merge + records a conflict; split-vs-edit content-preservation regression guard). No public API surface; the 179/179 GetRevisions + markup-corpus parity floors and the consolidate/composite suites are unchanged.
- **Consolidate — StackAll block-conflict duplicated the base block on reject + multi-reviewer table edits were silently dropped (B2/B4).** *Internal/experimental.* Two confirmed `IrCompositeMerger` (N-way consolidate) bugs. **(B2)** Under `ConflictResolution.StackAll`, a delete-vs-edit BLOCK conflict (one reviewer deletes a base paragraph, another edits it in place) stacked BOTH competitors verbatim — a base-anchored `DeleteBlock` AND a base-anchored `ModifyBlock`, both base-restoring on reject — so `RejectRevisions(consolidate)` reproduced the base paragraph TWICE (and doubled the `w:del` wrapper), breaking `reject ≡ base`. The StackAll block-conflict path now emits AT MOST ONE base-anchored op for the contested block (the lowest-reviewer-index competitor, verbatim) and re-emits every other competitor that carries right content as a base-ANCHORLESS `InsertBlock` (its own reviewer's right block, `SourceReviewer` set), which contributes a `w:ins`-wrapped block that vanishes on reject; a pure-delete competitor (no right content) is dropped (already captured in the recorded conflict). A `Debug.Assert` pins that exactly one emitted op consumes the base anchor (the block-level analogue of the token path's `AssertTilesBase`). BaseWins/FirstReviewerWins were already clean and are unchanged. **(B4)** When 2+ reviewers edited the SAME base TABLE — even DIFFERENT cells — `AllOpsIdentical`'s ModifyBlock branch compared `BlockResultText`, which returned `""` for any non-paragraph block, so two table edits mis-compared as `""=="" `and short-circuited to a FALSE consensus: only the first reviewer's table survived, the rest were silently dropped, and `conflictCount` was 0 (data loss). `AllOpsIdentical` now returns FALSE when any touched op is a table ModifyBlock (`TableDiff != null`), so multi-reviewer table edits fall through to the BLOCK-LEVEL conflict branch under the active policy (BaseWins keeps the base table + records the conflict; the other policies surface a recorded conflict too). `BlockResultText` gained an `IrTable` branch that serializes cell text so a table conflict's competitor `ResultText` is meaningful, and a `Debug.Assert` tripwire at the consensus emit trips if a multi-reviewer table edit ever silently reaches it. Under this fix even DISJOINT table-cell edits become a recorded block-level conflict (v1 does not compose table cells per-cell — see the v1 limitations note); `reject ≡ base` holds for the table-conflict output under all policies. New tests `IrCompositeFixTests` (delete-vs-edit reject≡base under all three policies; disjoint + same-cell multi-reviewer table edits record a conflict, reject≡base). No public API surface.
- **Markup round-trip — WC022 adjacent-empty-paragraph reject-order swap (M2.6 Task 2; WC022 closed, allowlist 2→1).** *Internal/experimental.* After the M2.4b bookmark-marker drop, WC022 still failed forward REJECT only: the block aligner's `IrBlockAligner.InOrderRefine` first-to-first matching crossed document order when several content+format-equal empty paragraphs competed. BEFORE had two adjacent empty paragraphs `[efb022 (empty+pPr), c88b (bare empty)]`; AFTER `[5e71 (bare empty), c88b (bare empty)]`. Scanned in right order, AFTER's `5e71` (no identity match) grabbed the only free bare-empty left `c88b`, stranding BEFORE's `efb022` to pair with AFTER's `c88b` — crossing left-document order, so `RejectRevisions` reconstructed blocks [8]/[9] swapped (accept and the reverse direction were already clean). Fixed by giving `InOrderRefine` a **same-unid identity-reservation phase**: a free right block first claims the free left block sharing its persisted unid (the IR's stable per-element identity — an unchanged paragraph keeps it across both documents) BEFORE any first-fit, keeping the pairing monotonic. Pure deterministic tie-break over equal-`(ContentHash, format)` candidates — it never changes WHICH blocks pair, only which identical left fills an identical right — so no other corpus pair shifts (markup corpus 182/184, fuzz 50/50, aligner/edit-script/parity suites green). NOT a 1:N problem (all pairings here are 1×1), so closed now rather than deferred to the 1:N split design. WC022 leaves the markup allowlist (**2→1**: only WC-BodyBookmarks remains). No public API surface.
- **DocxDiff revisions — anchor-presence contract pinned + documented precisely (M2.6 Task 2).** *Internal/experimental.* The `DocxDiffRevision` / `IrRevision` anchor-presence rule was documented as strict (`Inserted → RightAnchor only`, `Deleted → LeftAnchor only`) while the renderer intentionally emits BOTH anchors for a TOKEN-LEVEL insert/delete inside a Modified block (that block exists on both sides) — the shape a Python E2E observed as "a Deleted with both anchors". Surveying the whole WC corpus (both granularities × both directions) confirmed the actual invariant: each type's PRIMARY anchor is ALWAYS present, the opposite anchor is additionally present only for token-level ins/del; `FormatChanged` carries both; `Moved` is exclusive (source = left only, dest = right only). This is intentional and useful (a token-level edit can be located in either document), so the **contract docs were corrected to match the emission** — `DocxDiffRevision`/`IrRevision` XML-doc, `python/src/docx_scalpel/types.py`, and `npm/src/types.ts` now state the block-level-vs-token-level distinction. New public-surface tests (`DocxDiffTests`: token-level both-anchors, block-level primary-only, corpus-pair invariant) and a strengthened IR corpus test (Moved exclusivity) pin it. No behavior or API change — documentation + tests only.
- **DocxDiff `GetRevisions` exceeds the throwing oracle on whole-store note conversion (M2.6 Task 2; WC-BodyBookmarks final verdict).** *Internal/experimental.* WC-BodyBookmarks converts the document's entire endnote store to footnotes (BEFORE 24fn+190en → AFTER 213fn+0en) and carries many body-level bookmarks. On it `WmlComparer.Compare` THROWS `DocxodusException` "Internal error in ProcessFootnoteEndnote" and produces nothing — there is no oracle behavior to match. Our `DocxDiff.GetRevisions` completes without throwing and yields a substantial revision list in both directions: **we exceed the oracle on the consumer surface**, now pinned by a capability test (asserts the oracle throws AND our GetRevisions is total). Only the markup `Compare` round-trip fails — the cross-part 190-endnote→footnote migration the per-scope note diff is not built to reconcile — so the pair is RETAINED in the markup allowlist (now **1** entry) with the oracle-throws ceiling as context; fixing whole-store note-kind conversion is a large, isolated effort with negligible real-world value (the oracle can't do it; real documents don't flip their whole endnote store) and is not pursued.
- **Markup round-trip — note-id renumber output pass (M2.6 Task 1; WC034 foot+end closed, allowlist 4→2).** *Internal/experimental.* The IR markup renderer assembled the produced footnotes/endnotes parts keeping LEFT-package note ids, so a matched note whose body reference renumbered between revisions (WC034-After3: a note inserted ahead of it shifts left-en#1 → right-en#2) left the definition at its left id/part position — the accepted part's note sequence was `[en2,en1,en3]` vs RIGHT's `[en1,en2,en3]` (content correct, id/order wrong). New `IrMarkupRenderer.RenumberNoteIds` post-assembly pass mirrors the oracle's `WmlComparer.ChangeFootnoteEndnoteReferencesToUniqueRange`: it walks the produced body's footnote/endnote references in document order, renumbers each reference to its 1-based ordinal (separator/continuation boilerplate notes ≤ 0 keep their reserved ids and lead the part), and renumbers + reorders each definition to match — a `w:del` reference resolving the next deleted-only definition, an `w:ins`/equal reference the live (matched/inserted) definition with that id. To link an unchanged-but-renumbered matched note, `IrEditScriptBuilder.BuildOneStore` now also emits an all-`EqualBlock` note diff when a matched note's left/right ids differ (zero revisions — `EqualBlock` projects to nothing in `IrRevisionRenderer` — purely so the renderer reconciles the produced definition's id). Both WC034 foot+end pairs now round-trip clean in BOTH directions (ACCEPT==RIGHT notes, REJECT==LEFT notes) and leave the markup allowlist (**4→2** fixtures: WC022, WC-BodyBookmarks remain). The corpus markup invariant's `NoteContentHashes` now orders notes by body-reference document order rather than absolute id (the semantically faithful round-trip the oracle itself satisfies — accept-by-right-order, reject-by-left-order — robust to the legitimate id divergence between the two sides). The pass runs for every render; `GetRevisions` parity, the full IR.Diff suite, and the old-engine `RevisionProcessor`/`WmlComparer` suites stay green. No public API surface.
- **Markup round-trip — reject `w:del`/`w:ins` nested in `w:hyperlink` + empty-hyperlink drop (M2.5 Task 3; WC019 closed, allowlist 5→4).** *Internal/experimental.* A hyperlink-text edit (WC019: `www.ericwhite.com` → `www.ericwhite2.com`, rId4 retargeted) nests the revision markers INSIDE the `w:hyperlink` (the schema forbids a hyperlink inside `w:ins`/`w:del`, so the IR emits del-old-link/ins-new-link with the markers under each link). `RevisionProcessor`'s reject del→ins / ins→del reversal rules were gated to `parent==w:p`, so they never fired under a `w:hyperlink` — REJECT left the deleted link's content unrestored and the inserted link's content unremoved (reject ≠ left). Two minimal, additive `RevisionProcessor` changes close it: (1) the two reject reversal rules now also fire when `parent==w:hyperlink` (WmlComparer never produces del/ins-in-hyperlink — its `PreProcessMarkup` strips hyperlinks via `RemoveHyperlinks` — so no existing case is touched; this only ADDS handling for a previously-unhandled valid shape), and (2) the accept transform drops an EMPTY `w:hyperlink` shell (no surviving run content) — the artifact the del-old-link/ins-new-link shape leaves after accept/reject collapses one link's content. WC019 now round-trips clean in BOTH directions (ACCEPT==RIGHT, REJECT==LEFT, content + format) and leaves the markup allowlist (**5→4** fixtures). The full old-engine `RevisionProcessor`/`WmlComparer` suite (95 tests) and the full suite (1939) stay green. No public API surface.
- **Diff engine — affix-trim word boundary mirrors `GetComparisonUnitList` (M2.5 Task 3; WC-1920 closed).** *Internal/experimental.* The compat-mode common-affix trim's word-boundary back-off (`IrRevisionRenderer.IsWordBoundaryBefore`) treated EVERY non-letter-digit char as a boundary, so `This is a test` → `This is a test!` (a `!` appended in a separate run, which the reader's N5 coalescing already joins to `test!`) trimmed the shared `test` and reported just `ins !` (1 revision) where `WmlComparer` keeps `test!` whole and reports del `test` + ins `test!` (2). The boundary test now mirrors `WmlComparer.GetComparisonUnitList` EXACTLY (`IsOracleSplitChar`): the comparer groups per-character atoms into words where a char is an ISOLATED atom (boundary on both sides) iff it is a `WordSeparators` member, a CJK ideograph, or a NON-digit-adjacent `.`/`,`; every other char — letters, digits, and other punctuation (`!`/`?`/`:`) — JOINS the surrounding word. So `test`/`test!` now share no trimmable affix and the whole-word del+ins survives, while the `.`/`,` isolation (with the `3.14` digit-adjacency carve-out) keeps the `This`/`This.` and `endnote`/`endnote.` trims working. **WC-1920 is a genuine PASS (8==8); genuine-pass ratchet 176→177** — zero scoreboard blast radius (no other row changed), compat-mode differential MATCH 136→150, full IR.Diff suite green. No public API surface.
- **Diff engine — note-store reference-order correspondence (M2.5 Task 3; WC-1710/1720 closed).** *Internal/experimental.* The IR no longer pairs footnotes/endnotes by raw `w:id`. `WmlComparer` does not either: `ChangeFootnoteEndnoteReferencesToUniqueRange` renumbers every note id in BODY-REFERENCE ORDER and `ProcessFootnoteEndnote` pairs a note with another note IFF their body references correlate Equal — so a note's correspondence is its body reference's, never the stored id. WC034-After3 relocates an endnote reference INTO the middle of `Video` (a NEW en#1 inserted, Before's en#1 content pushed to en#2), so by-id pairing cross-matched unrelated notes and over-reported. `IrEditScriptBuilder.BuildOneStore` now collects each side's referenced note ids in body document order (`CollectNoteReferenceOrder`, recursing tables/textboxes/fields/hyperlinks) and aligns the two sequences (`AlignNoteReferences`): an exact-content LCS spine first, then a best-content-similarity residue (Jaccard over the notes' WORD tokens — separators/markers excluded so length doesn't dominate) with a FORCED lone-left/lone-right pair (the 1×1-residue rule). So Before-en#1 ↔ After3-en#2 (content modify) and After3-en#1/en#3 are whole-note inserts, matching the oracle. **Invariant:** when the reference sequences pair in order (no inserted/deleted reference shifted them — the common case) this reduces EXACTLY to the former by-id pairing, so unrelated fixtures (WC-1600/1660/1750/…) are byte-identical — the order-preserving spine + forced-1×1 residue is precisely what keeps the single-edited-note case a modify rather than the del+ins a prior unconstrained-similarity trial produced (which regressed WC-1660/1670/1750/1760). Paired with a **structural-only affix-trim guard** (`IrRevisionRenderer` region flush): the compat common-affix trim no longer cancels a del/ins region whose two sides are byte-IDENTICAL text (the M2.5 Task 1 intra-word note-ref relocation — `Vi`⟨ref⟩`deo` vs `Video`, the relocated ref textless — which `WmlComparer` reports as del `Video` + ins `Video`; the old trim erased both to empty). Together these recover the oracle's exact counts: **WC-1620/1630 (footnotes, 3==3), WC-1710/1720 (endnotes, 7==7) are now GENUINE PASSES** — GetRevisions genuine-pass ratchet **174→176** (179/179 PASS-or-deviation, 0 FAIL); 3 surviving deviations (WC-1450/1830 sub-paragraph 1:N split; WC-1920 cross-run word coalescing). `IrNoteDiff` gains a `LeftNoteId` (a matched pair can carry distinct left/right ids; the verifier resolves the left store by it and the markup renderer locates the source note by it); JSON round-trip carries it. The WC034 markup round-trip stays allowlisted: the note CONTENT the markup produces is now correct (verified), but the produced part's note-ELEMENT order is left-order ([en2,en1,en3]) vs the right's renumbered order ([en1,en2,en3]) — a note renumber/reorder markup pass, distinct from the correspondence (now done). Full suite 1939/1939, full IR.Diff suite green, WC corpus apply-verify + JSON round-trip green. No public API surface.
- **Diff engine — intra-word note-ref interruption tokenization (M2.5 Task 1; WC-1710/1720, WC034).** *Internal/experimental.* `IrDiffTokenizer` now models a note reference (or any zero-width CONTENT atom — `IrNoteRef`/`IrInlineImage`/`IrOpaqueInline`/`IrTextbox`) that splits a word WITHOUT a separator on either side as a genuine word-structure change: the two flanking word tokens carry a sentinel-framed interruption marker (keyed on the interrupting atoms) in their `MatchKey`, so `Vi`⟨ref⟩`deo` is NOT word-equal to a contiguous `Video`. This is the established WC034-After3 case — an endnote/footnote reference RELOCATED INTO THE MIDDLE of the body word `Video` (runs `Vi`[ref]`deo` vs Before's contiguous `Video `[ref], verified run-by-run) — where `WmlComparer` correctly reports a `Video` del+ins and the IR's id-less per-run note-ref tokenization previously read the word as unchanged. The Fine-mode edit script now correctly reports the body change (verified: `ModifyBlock` token diff Delete `Vi` / Insert `Vi`+ref+`deo`). The detection keys on char-offset adjacency (the flanking words touch the zero-width atom with no `Separator` token between), so a ref BETWEEN words (separator-adjacent — the overwhelmingly common case) is untouched: its word keys are byte-identical to today and the ref token's own key stays position-independent. **Zero corpus blast radius** — full GetRevisions scoreboard unchanged at 174 PASS + 5 deviation + 0 FAIL, full IR suite 408/408. Engine truth (runs unconditionally; identical under Fine and WmlComparerCompatible). The WC-1710/1720 GetRevisions COUNT and the WC034 markup round-trip remain documented deviations: their residual is a SEPARATE, broad note-store-correspondence concern (the IR aligns notes by `w:id` while the oracle aligns by content/reference-order, remapping ids — WC034-After3 renumbers a note), plus the WmlComparerCompatible common-affix trim cancelling the text-identical body del/ins; both are deferred (note-store cross-correspondence, the deferred-item-#5 class), with the now-isolated root cause recorded in the scoreboard catalog and the markup allowlist. No public API surface.
- **Diff engine — deviation-burndown close (M2.4b Workstream D; WC-1900, WC019, WC022).** *Internal/experimental.* The final burndown workstream, closing the textbox-duplicate and markup-leftover gaps under the binding method rule. **(1) Non-adjacent Choice/Fallback textbox dedup** (`IrRevisionRenderer.RenderTextboxDiffs`): `WmlComparer.PreProcessMarkup` opens the document with `MarkupCompatibilityProcessMode.ProcessAllParts` (Office2007), MC-RESOLVING each `mc:AlternateContent` to a single branch and discarding the other (evidence: the WC048 Choice declares `Requires="wps"`, an Office2010 namespace the Office2007 processor cannot satisfy, so the SDK keeps the VML Fallback) — so the oracle counts ONE copy of each logical textbox. The IR reader walks BOTH branches for markdown-projection parity, and the old render-time dedup only collapsed ADJACENT duplicate batches (pair-walk on i/i+1). A NESTED textbox interleaves an empty wrapper body between the two branches (`[Textbox3, ⌀, Textbox3, ⌀]`), so the pair-walk never matched and WC-1900 (textbox-in-cell) leaked +2 revisions. The dedup now uses content-signature OCCURRENCE PARITY (emit odd occurrences, drop the even Fallback copy) so the duplicate collapses wherever it lands. **WC-1900 is a genuine PASS (6 == 6); GetRevisions genuine-pass ratchet 173→174** (179/179 PASS-or-deviation). WC-1920's residual −1 is re-diagnosed as a tokenizer punctuation-attachment grain (`test`/`test!`) inside a textbox-nested table, deferred to M2.5. **(2) True hyperlink rId remap** (`IrMarkupRenderer.ImportHyperlinkAndExternalRelationships`): when a right-sourced hyperlink's rId collides with a DIFFERENT left relationship (WC019: Before→`ericwhite.com`, After→`ericwhite2.com`, both `rId4`), the import previously refused to clobber and left the cloned `@r:id` resolving to the LEFT target; it now mints a fresh relationship id, recreates the right target under it, and rewrites the cloned `w:hyperlink/@r:id` (accept resolves the correct target — verified). **(3) Body-level bookmark drop** (`IrReader.AppendBlocks`): a bookmark marker appearing as a DIRECT body/cell child (legal OOXML) fell through to an `IrOpaqueBlock`, so an inserted/orphaned body-level `bookmarkEnd` became a spurious content block the markup round-trip could not toggle. The block-level N3 rule now drops it, mirroring `WmlComparer`'s `PreProcessMarkup` (`MarkupSimplifier RemoveBookmarks=true`); WC022 goes from 0/4 to 3/4 round-trip sub-checks passing. The **round-trip allowlist shrinks 8→5 fixtures** (WS-A removed the 3 SmartArt fixtures earlier in the milestone); the 5 survivors — WC034 ×2 (mid-word note-ref, oracle correct), WC022 (residual adjacent-empty-paragraph alignment ordering), WC019 (residual: rejecting `w:del`/`w:ins` nested in `w:hyperlink`, a shared-`RevisionProcessor` gap the oracle sidesteps via `RemoveHyperlinks`), WC-BodyBookmarks (residual endnote→footnote note-store conversion) — all carry established root-cause evidence (none a renderer-markup gap) and are deferred to M2.5. The reader change does not perturb markdown-projection equivalence (26/26). No public API surface. See the [M2.4b plan Outcome](docs/superpowers/plans/2026-06-11-diff-m24b-deviation-burndown.md#m24b-outcome) for the full per-row verdict table.
- **Diff engine — structural-deviation closures (M2.4b Workstream C; WC-1210/1420/1430/1440/1840/1770/1750/1760).** *Internal/experimental.* Per the binding method rule (`WmlComparer`'s `GetRevisions` count is the presumed-correct oracle; the IR is fixed to match unless an oracle fault is established with evidence), four changes close eight more deviation rows. **(1) Adjacent-block insert/delete coalescing** (`IrRevisionRenderer.RenderBlockOpList`, WmlComparerCompatible mode only): `WmlComparer` groups the produced document's atoms by adjacent correlation status, so a contiguous run of inserted (resp. deleted) blocks is ONE revision; the IR surfaced one per `InsertBlock`/`DeleteBlock` op. A maximal same-direction run now coalesces into one revision (block texts joined with paragraph-mark newlines), splitting sub-regions at table/opaque boundaries (a table STARTS a new region that joins the inserts following it — WC-1210 `Abcde` | empty-table+`fghij`) and gating on text content so a pure math/image/opaque/empty run stays one-per-block (WC-1550 two-maths, WC-1320/1340/1350 standalone image/SmartArt keep their counts). Applied at body, note-scope-excluded, and table-cell op levels. Closes WC-1210/1420/1430/1440/1840. **(2) Textbox-interior compat coarsening** (`RenderTextboxInnerOp`): `WmlComparer` treats `w:drawing` as one opaque comparison atom and never descends `w:txbxContent` (WmlComparer.cs ~L8225/8673), so a changed textbox surfaces as one whole-paragraph del+ins; the IR reader models the interior (for the markdown projection) and could split it finer (WC-1770 `In1`→`In` → a lone Deleted `1`). A textbox-interior `Modified` paragraph now renders as whole-block Deleted+Inserted, validated against WC-1890/2080/2090/2092 (all keep passing). Closes WC-1770. **(3) Table-aware block similarity + unambiguous table-residue pairing** (`IrBlockSimilarity`/`IrBlockAligner.FillOneGap`, an ALIGNMENT capability addition live in BOTH Fine and compat modes): an `IrTable` pair is now scored by Jaccard over concatenated cell-paragraph token multisets, and a gap with exactly one free-left + one free-right table pairs them as `Modified` regardless of score (the table analogue of the 1×1 residue), feeding `IrTableDiffer`'s row/cell diff. The two endnote tables of WC-1750/1760 now produce per-cell edits (`Aaa→Aaa1`, `Eee→Eee1`, deleted `Ggg`/`Hhh`/`Iii` row). **(4) Note-scope empty-mark prune scoping**: the empty-paragraph-mark prune is restricted to BODY paragraphs, since `WmlComparer`'s per-note atom grouping DOES surface an inserted/deleted empty paragraph in a footnote/endnote scope (WC-1750/1760's deleted trailing rows leave an empty-paragraph insert reported as `\n`). With (3) these close WC-1750/1760. **GetRevisions scoreboard 165→173 PASS** (179/179 PASS-or-deviation, 0 FAIL; new PASS-only ratchet at 173). **WC034 'Video' investigation:** the oracle's `Video` del+ins (WC-1710/1720) is **CORRECT, not spurious** — examining the raw OOXML shows an endnote reference (id=1) relocated INTO THE MIDDLE of the word in After3 (runs `Vi`[en-ref]`deo` vs Before's contiguous `Video `[en-ref]), a real run-structure change; the IR's id-less, per-run note-ref tokenization is coarser there and reports the word unchanged (a deferred M2.5 tokenizer item), correcting the prior "oracle spurious" misdiagnosis. The remaining deviations are genuine engine-grain (WC-1450/1830 block-vs-atom paragraph granularity inside a cell, where sub-paragraph content migrates across the paragraph boundary) or out-of-scope (WC-1900/1920, WS-D textbox-duplicate dedup). The table-similarity engine change is verified Fine-safe — full IR suite (400), differential triage (compat MATCH **136**, Fine classification unchanged), and 50-seed fuzz (alignment + apply-verify + JSON round-trip, 0 regressions) all green; full Release suite 1931/1931. No public API surface.
- **Diff engine — low-coverage Modified rendering closes the "coincidental Equal island" deviation family (M2.4b Workstream B; WC-1170/1190/1950).** *Internal/experimental.* The block aligner's 1×1 gap residue (`IrBlockAligner.FillGaps`) pairs a near-rewritten paragraph/cell as `Modified` *regardless of similarity score* (the only sensible reading of a lone in-gap pair); Myers then credits the few COINCIDENTALLY shared words as `Equal` islands, splitting one logical rewrite into more `Inserted`/`Deleted` revisions than `WmlComparer`'s whole-document LCS, which reports the rewrite as one contiguous del + one ins. Per the binding method rule (`WmlComparer`'s count is the presumed-correct oracle), the IR is reconciled at **render time, WmlComparerCompatible mode only** (Fine grain — the engine's truth — is byte-untouched), via two `IrRevisionRenderer` transforms derived empirically from the failing rows: **(1) low-coverage coarsening** — when a Modified pair's `Equal`+`FormatChanged` CONTENT-token coverage (larger-covered side) is below **0.67** AND the larger side carries at least **8** content tokens (a near-rewrite, not a short edit), the word-bearing `Equal` islands are bridged into the open region exactly like separators so the change collapses to ONE del + ins (the existing word-boundary affix trim still keeps any wholly-common edge); **(2) empty-paragraph-mark prune** — a whole-block insert/delete of a paragraph with ZERO content tokens (a bare paragraph mark, e.g. the empty cell paragraph a moved-into-table block leaves behind) surfaces no `WmlComparer` revision, so it is suppressed. **Closures (genuine PASS, removed from `DocumentedDeviations`):** WC-1170 (`Video provides.`→a 36-token paragraph: coincidental `Video` island), WC-1950 (cell rewrite sharing only function words `the`/`of`/`each`), WC-1190 (the spurious 3rd revision was the empty-mark cell paragraph). The **0.67 / ≥8** thresholds were swept over the corpus to a stable plateau (floor 0.55–0.72 × min 6–10, identical result) and are **size-and-coverage-gated so no passing row regresses** — verified: WC-1930's 3-token `designs that compleme` short edit stays fine-grained (max-coverage 0.33 but below the size gate), and the empty-mark prune is strictly the bare-mark case so the standalone math/SmartArt/image paragraph inserts/deletes `WmlComparer` DOES count (WC-1320 deleted SmartArt, WC-1550 two-maths, WC-1340/1350 images) keep their revisions. **GetRevisions scoreboard 161→165 PASS** (179/179 PASS-or-deviation, 0 FAIL; floor 179 held). The 8 residual rows of the family are **NOT the coincidental-island mechanism** and stay catalogued with sharpened per-row evidence: WC-1420/1430/1440 (high-coverage in-place math-run edits above the coarsening floor — Fine engine grain the coarsening deliberately skips), WC-1440/1450/1830/1840 (adjacent-block-insert coalescing + math-only-paragraph-in-region — a Workstream C structural item; naive adjacent-insert coalescing regressed the WC-1550 standalone-math counts), WC-1210 (empty STRUCTURAL-table insert — table-aware prune deferred to WS-C), WC-1770 (textbox-interior UNDER-report: the IR's single-atom `1` deletion is the more precise account, one short of the oracle's whole-textbox-paragraph del+ins). Markup parity scoreboard unchanged (39/39 — render-side change, no produced-markup effect); Fine-mode tests byte-stable; differential, full Release suite, and 50-seed fuzz green. No public API surface.
- **Document IR — relationship-id-stable opaque hashing (M2.4b Workstream A; WC-1940 + SmartArt markup parity).** *Internal/experimental.* `IrHasher.Canonicalize` hashed raw relationship-id attribute *values* (`r:id`/`r:embed`/`r:dm`/`r:lo`/`r:qs`/`r:cs`/…), so an **unchanged** diagram/image whose rel ids renumber between revisions (or get freshly minted on accept by `MoveRelatedPartsToDestination`) read as *changed* — its opaque content hash differed side-to-side, and the aligner paired the paragraph del+ins instead of Equal. Two confirmed mechanisms: (1) reader-side, the SmartArt paragraph's `wp:docPr/@id` (a non-content drawing-object id) renumbered `1`↔`2` and the diagram rel ids differed (WC-1940 / WC052); (2) accept-side, the re-imported diagram's `dgm:relIds` got fresh `R…` ids so the accepted block's hash matched neither input (WC014 ×2 markup round-trip). Fixed at the reader/hasher level, mirroring `WmlComparer.CloneBlockLevelContentForHashing` (the parity oracle): a new part-aware `IrHasher.Canonicalize(element, IrRelResolver)` overload replaces every relationship-namespace attribute *value* with a stable content-identity token — internal media part → `part-sha:<content-hash>` (bytes streamed, cached per part per `Read`), internal **xml** part (diagram data/layout/colors, charts) → dropped to a sentinel (matching the oracle, whose serialized xml-part bytes vary cosmetically across saves even when unchanged), external/hyperlink rel → `ext:<target-uri>`, dangling → sentinel, any resolution failure → `unresolved` (totality — never throws) — and strips the renumber-prone `wp:docPr/@id`. The reader threads an `IrRelResolver` (owning part + per-`Read` byte-hash cache, reusing the `ResolveImagePart` cache pattern) through every opaque-hash call site per scope. **Closures:** WC-1940 → genuine PASS (scoreboard **162 PASS + 17 deviation = 179**, composition shift deviation→PASS); the three SmartArt markup fixtures (WC014 ×2, WC052) round-trip clean and leave the Task-4 allowlist (8→5 entries / 6→5 root causes). WC022 was **re-diagnosed**: its image/math drawing rel ids *do* renumber and now round-trip identically after this fix, but a residual body-level `w:bookmarkEnd` marker (the WC-BodyBookmarks root cause, **not** the rel-id renumber the catalog claimed) still blocks it — corrected catalog text, kept on the allowlist for Workstream D. New headline guards pin both directions: different rel id + same part bytes → **same** opaque hash; same rel id + different part bytes → **different** opaque hash (the M1.2 content-sensitivity guarantee survives); dangling rels tolerate to a stable token. 5 IR golden snapshots regenerated (only `contentHash`/opaque `hash` values changed — anchors stable); markdown-projection equivalence (the emitter never surfaces opaque hash values), differential, full WC corpus round-trip, and 50-seed fuzz all green.
- **Diff engine — NBSP conflation now happens at tokenize-SPLIT time (WC-1970/1980 parity).** *Internal/experimental.* `IrDiffTokenizer` folded a non-breaking space (U+00A0) to an ordinary space only in the post-split *match key*, while U+00A0 was never a `WordSeparators` member — so a pure space→NBSP edit (e.g. `l'article 1` → `l'article` + NBSP + `1`) tokenized to **different word/separator boundaries** on the two sides (the NBSP side glued `l'article 1` into one word; the space side split it three ways), producing a spurious 2-revision diff where `WmlComparer`, which conflates NBSP→space *before* its word split, correctly reports **0**. Fixed: when `ConflateBreakingAndNonbreakingSpaces` is true, U+00A0 is now treated as a separator **equivalent to `' '` at split time** (its Separator token's match key normalizes to `" "`; its raw `Text` preserves the original U+00A0 so output stays byte-faithful); when false, behavior is unchanged. This corrects a misdiagnosis in the parity-deviation catalog — WC-1970/1980 (WC055/WC056) were catalogued as a `WmlComparer` "oracle under-report" but were an IR engine bug; both now genuinely **MATCH** at 0 and are removed from `DocumentedDeviations` (scoreboard **161 PASS + 18 documented deviation = 179/179**, deviation count 20→18). Catalog text for WC-1170/1190 (the real coincidental shared token is `Video`, not `provides`) and WC-1710/1720/1940 (precise per-word/aligner root cause) corrected in place; the M2.3 differential `OldEmpty`-bucket notes carry a dated retraction. New `IrDiffTokenizer` tests pin the space↔NBSP match-key-sequence equivalence under conflation and its divergence when off. Full suite + 50-seed fuzzer + differential green.
- **Document IR — M1.5 pre-Phase-2 hardening (sweep + revision-skip soundness).** *Internal/experimental.* Two principled sweep fixes lifted corpus markdown byte-equivalence **642 → 648/668** (each adds a rule pin; both strictly improve IR fidelity): (1) the EmitBlocks trailing-blank-line rule now keys on the oracle's *structural* `IsListItem` verdict (`w:numPr` present inline or via the `pStyle→basedOn` chain, **numId-agnostic**), captured by the reader as `IrParagraph.IsListItemForLayout` — so a `Subtitle`/`Heading{N}` style whose chain carries a bare `<w:numPr><w:ilvl/></w:numPr>` (no `numId`) gets the same spacing the oracle gives it, while its resolved `List` correctly stays null (closes HC007/HW010 + the HC005/HC048 cascades); (2) a complex field that reached its `separate` but whose closing `end` is implied at paragraph close (a TOC field) now emits a faithful run-based `IrFieldRun` carrying the computed result instead of dropping it to an opaque capture — so the result text reaches both the rendered markdown and `TextPreview`, matching the oracle's raw `Descendants(w:t)` view (closes HC022/HC031; HC031 snapshot regenerated + reviewed). The remaining 20 divergences are all accepted oracle-bug-family (special-char drops, multi-run hyperlink/emphasis splits, customXml-range content-control acceptance — every case where the IR is *more* correct), bundled with the D3 cutover. Separately, the `IrReader.Read` revision-skip scan that avoids the `RevisionProcessor` round-trip on revision-free documents was made **provably sound**: its element set is now a true superset of every name `RevisionProcessor` dispatches on (was missing `w:tblPrExChange`, the unconditionally-rewritten `w:delText`/`w:delInstrText`, and the full move/cell/customXml-RangeEnd set) and it scans every part the reader walks (was `MainDocumentPart`-only, missing header/footer/footnote/endnote/comment-only revisions). New `IrRevisionSkipTests` add behavioral guards (`tblPrExChange`-only and header-only insertion reads match `RevisionProcessor.AcceptRevisions`) plus a set-drift guard pinning the scan set to `RevisionProcessor`'s dispatch so it cannot silently rot. No public API change.
- **Solution builds no longer race the WASM-mode assembly.** `DocxodusWasm` references `Docxodus` with `WASM_BUILD=true`, so every solution build compiled Docxodus twice into the same `bin/<Config>/net8.0/` output — whichever finished last won, and `Docxodus.Tests` intermittently linked the SkiaSharp-free WASM assembly (`error CS1061: 'ImageInfo' ... 'SaveImage'`). WASM-mode output now builds into isolated `bin/wasm/` + `obj/wasm/` paths; the `dotnet clean` workaround documented in CLAUDE.md is no longer needed.

### Added
- **`DocxDiff` consolidate — per-cell table composition (disjoint cross-reviewer cell edits compose inline).** *Internal/experimental.* Previously 2+ reviewers editing the SAME base table — even DIFFERENT cells — routed to a whole-table block conflict (`BaseWins` kept the base table; disjoint cell edits were NOT composed). Now DISJOINT cross-reviewer table-cell edits COMPOSE inline (Alice edits cell(0,0), Bob edits cell(1,2) → both land, each authored to its reviewer); only edits to the SAME cell by ≥2 reviewers become a recorded conflict resolved by the policy. A new TABLE-COMPOSE branch in `IrCompositeMerger.MergeOneBaseBlock` (before the block-conflict branch) fires when every toucher is a table `ModifyBlock` and the gates pass: `ComposeTableDiffs` aligns rows by base row anchor (0 touch → EqualRow; 1 ModifyRow → that reviewer's cells; 1 DeleteRow → authored delete; ≥2 ModifyRow → per-cell compose; delete-vs-modify same row → row conflict; ≥2 deletes → consensus delete; InsertRows by different reviewers route by preceding base row, both appear). Per base cell: 0 changed → base passthrough; 1 → that reviewer's BlockOps authored; ≥2 → **RECURSE** into the SAME body block/token composition over the cell's paragraph mini-body (a new reusable `MergeBlockStream` runs the body `GroupByBaseAnchor` + `MergeOneBaseBlock` loop over the per-reviewer cell `BlockOps` against the base cell's block anchors), so disjoint words inside one cell paragraph fuse and same-word edits become a cell-paragraph-anchored conflict. The op's `Op.TableDiff` is the merged apply/JSON truth; an additive `IrCompositeOp.AuthoredRows` (→ `IrAuthoredRowOp`/`IrAuthoredCellOp`, default null) carries the renderer/revision attribution view. A new `IrMarkupRenderer.RenderComposedTable` (a shared `RenderOneCompositeBlock` helper factored out of `IrCompositeMarkupRenderer.EmitCompositeOp`, reused per cell-block) emits a SINGLE `w:tbl` on the base table's tblPr/tblGrid with per-cell `w:ins`/`w:del` authored to the right reviewers and base cells verbatim; `IrCompositeRevisionRenderer` and `IrCompositeScriptJson` gained additive composed-table branches (`authoredRows` JSON mirrors `authoredTokens`). **Three STOP-boundary fallbacks** to the whole-table block conflict (no silent loss — base table kept under `BaseWins`, disagreement recorded under every policy): a reviewer `MovedRow` (`NoReviewerHasMovedRow`), a column-count change / unpaired cell (`AllColumnStructureStable`), and cell-shell-prop changes the IR does not model (a pure `w:tcPr` width / gridSpan / vMerge change leaves every IR hash identical → the table reads as `EqualBlock`, never reaching the branch — invisible, not dropped). Debug guards: `AssertTilesBaseTable` (rows tile base rows, cells tile base cells) parallels the token path's `AssertTilesBase`. **The base two-way engine is UNCHANGED — the parity scoreboard stays 179/179** (the composite path is separate; the 2-way `RenderModifiedTable`/`RenderModifyRow` are untouched). reject ≡ base for all policies; accept = policy-resolved. New tests: `IrCompositeTableTests` (15 scenarios: disjoint compose, consensus, same-cell conflict, same-cell disjoint-words recursion, disjoint InsertRow, delete-vs-modify rows, MovedRow + column-count fallbacks, all-policy conflict, renderer/revisions/JSON/verifier) + `CompositeFuzzTests.Composite_disjoint_table_cells_round_trip` (different reviewers edit different cells; reject ≡ base + table-aware `IrCompositeVerifier` over 2/3/4-way seeds). See the "N-way composite / Consolidate" section of `docs/architecture/ir_diff_engine.md`.
- **`DocxDiff` consolidate — native move composition (non-colliding reviewer moves render as `w:moveFrom`/`w:moveTo`).** *Internal/experimental.* Previously the N-way merger LOWERED **every** reviewer move to `w:ins`/`w:del` (content preserved + round-tripped, but no native move markup). Now a SINGLE reviewer's `MoveBlock`/`MoveModifyBlock` whose source base block is touched by **only that reviewer** renders as a NATIVE move — `w:moveFromRangeStart`/`End` + `w:moveFrom` (source) and `w:moveToRangeStart`/`End` + `w:moveTo` (destination), authored to the mover, both halves sharing one move-group id — and surfaces as `Moved` source + destination revisions and a `MoveBlock` pair in the composite edit-script JSON (`moveGroupId`/`isMoveSource` on both halves). Move-group ids are **globally namespaced** across reviewers (a single deterministic counter, reviewers in list order then ascending local gid), so two reviewers' independent native moves never collide on a shared `w:name`. COLLIDING moves keep the existing behavior: a move-vs-edit on the same base block, or two reviewers moving the same block, both LOWER to del/ins and record a conflict (contested relocation; reject ≡ base under every policy). Split/Merge still lower (out of scope). Gated on `DetectMoves` (false ⇒ moves still lower, as before) and `SimplifyMoveMarkup` (degrades a native move to del/ins). The base two-way engine is unchanged (parity scoreboard stays 179/179); the work is concentrated in `IrCompositeMerger` (`PlanMoves` + `ApplyMovePlan` + right-positioned move-dest routing) — the markup/revision/JSON renderers needed NO functional change (their native-move handling was already in place). New tests: `IrCompositeMoveTests` plus native-move cases in `IrCompositeMarkupRendererTests`/`IrCompositeRevisionTests`/`IrCompositeJsonTests`.
- **`DocxDiff` — N-way composite / consolidate (closes the last `WmlComparer` gap).** A new public surface on `DocxDiff` (`Docxodus/DocxDiff.cs` + `Docxodus/DocxDiffConsolidate.cs`, `#nullable enable`, fully XML-doc'd) merges **N reviewers** — each an independently revised copy of ONE shared base — into a single multi-author tracked-changes document. Four entry points: **`Consolidate(base, reviewers, settings?) → WmlDocument`** (native `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange` markup, each reviewer's edits stamped with that reviewer's own author name), **`GetConsolidatedRevisions(...) → IReadOnlyList<DocxDiffConsolidatedRevision>`** (the attributed revision list — `DocxDiffRevision` shape + the contributing reviewer's `Author` + a `ConflictId` link), **`GetConsolidatedEditScriptJson(...) → string`** (the composite edit script as data: every op additively carries `author`/`sourceReviewer`/`conflictId`/`authoredTokens`/`sourceRightAnchors`, plus a top-level `conflicts` array), and **`GetConflicts(...) → IReadOnlyList<DocxDiffConflict>`** (the inspect-before-merge view). New public types `DocxDiffReviewer {Document, Author}`, `DocxDiffConsolidateSettings {Diff, ConflictResolution}`, `DocxDiffConsolidatedRevision`, `DocxDiffConflict {Id, BaseAnchor, TokenStart, TokenEnd, AppliedPolicy, Competitors}`, `DocxDiffConflictCompetitor {Author, ResultText}`, and the `ConflictResolution { BaseWins (default), FirstReviewerWins, StackAll }` enum. **Algorithm (`IrCompositeMerger`):** N pairwise edit scripts `Build(base, reviewer_i)` share one base anchor space, so per base block the merge is exact — untouched → passthrough; one reviewer → that reviewer's op (authored); ≥2 identical → consensus (one op); ≥2 all-paragraph token edits with unchanged paragraph properties → **token-span composition** (non-overlapping word edits compose inline, each authored; overlapping spans → conflict per policy); anything else (delete-vs-modify, a pPr+text edit) → a block-level conflict per policy. Block-level inserts never conflict (all appear, ordered by reviewer index). Conflicts are **always recorded in the data** regardless of policy: BaseWins keeps the base text at the conflicted span (all competitors recorded), FirstReviewerWins applies the first reviewer inline, StackAll emits each competing edit in order. **Round-trip:** `reject(consolidate) ≡ base`, `accept ≡ the policy-resolved composite`. Single-document `Compare` output is byte-unchanged (the markup renderer's author override is additive). **Parity:** a scoreboard over all 84 legacy `WmlComparer.Consolidate` corpus cases (WC001 multi-reviewer + WC002 single-reviewer) reports **84/84 reproduce-PASS, 0 deviations, 0 fails** under the SOUND-SEMANTICS metric — legacy Consolidate is a JUXTAPOSITION tool (it keeps the original and APPENDS each reviewer's labeled copy in colored boxes, even for one reviewer; it never inline-merges), so the IR-native true inline merge deliberately supersedes legacy's whole-corpus shape (the per-row deviation catalog is empty because no corpus case has a true token-overlap conflict; that supersession path is exercised by unit tests). Single-reviewer cases reproduce the reviewer document char-exact at accepted-body text (accept ≡ right, all 74); multi-reviewer no-conflict cases compose every reviewer's added tokens. **v1 limitations (documented, honest):** note-scope (footnote/endnote) diffs are NOT merged (a `Debug.Assert` tripwire guards against silent drop — follow-on); multi-reviewer edits to the SAME table are not composed per-cell — they are surfaced as a block-level conflict under the active policy (BaseWins keeps the base table; per-cell table composition is a follow-on, with a `Debug.Assert` tripwire guarding against silent drop); cross-reviewer move-vs-edit and split/merge-vs-edit collisions resolve at BLOCK granularity; a reviewer changing both a paragraph's text AND its pPr routes that block to the conflict path; conflict spans are reported as base TOKEN indices + `BaseAnchor` (not char offsets). Surfaced through every layer: WASM (`DocxDiffBridge.Consolidate`/`GetConflictsJson`/`GetConsolidatedRevisionsJson`/`GetConsolidatedEditScriptJson`), npm (`docxDiffConsolidate`/`docxDiffGetConflicts`/`docxDiffGetConsolidatedRevisions`/`docxDiffGetConsolidatedEditScript`), and docx-scalpel (`docx_diff_consolidate`/`docx_diff_get_conflicts`/`docx_diff_get_consolidated_revisions`/`docx_diff_get_consolidated_edit_script`), all routing through `DocxDiffOps`. See the "N-way composite / Consolidate" section of `docs/architecture/ir_diff_engine.md`.
- **Diff engine — composite (K-reviewer) consolidate fuzzer (T5.2).** *Internal/experimental, test-only.* A deterministic seeded fuzzer for `DocxDiff.Consolidate`, the multi-reviewer analogue of the single-doc `DiffFuzzer`. `DiffFuzzer.GenerateComposite(seed, reviewerCount)` synthesizes ONE shared base (the existing `GenerateBase`) and `reviewerCount` reviewer documents, each a DEEP-COPY of the base (`DocModel.Clone` — every Para/Run/Table-row reconstructed, no aliasing) with a DISJOINT partition of comparable mutations applied: reviewer `i` edits only base paragraphs in its residue class (`index % reviewerCount == i`), 1–3 mutations apiece, with a 1-in-8 deliberate-collision chance of touching a foreign paragraph. The v1 set is paragraph-only (EditWord / InsertParagraph / DeleteParagraph) and the composite base carries NO table or footnote — both project to a text view the apply-verifier's body-only `Docs.PlainText` oracle (direct-child `w:p` / `w:t` only) cannot match: a footnote reference tokenizes to an atomic NoteRef token (MatchKey `fn`) and a non-paragraph block reconstructs as its content hash, neither of which appears in body plaintext. (The composite merger also does not yet build note-scope ops — `IrCompositeScript.NoteOps` is always null — so a reviewer footnote edit would be silently dropped; tracked for a later note-consolidation task.) `CompositeFuzzTests` asserts the consolidate own-oracle across many seeds: **round-trip — reject ≡ base** (50 seeds × 3/4/5-way = 150 cases) and the **composite apply-verifier** (`IrCompositeVerifier`, 30 seeds × 3/4-way = 60 cases). No real consolidate bug surfaced (the two initial apply-verifier failures were the table/footnote text-projection asymmetry above, resolved by scoping the v1 generator; round-trip held on every seed including footnote/table bases). No public API.
- **Diff engine — M2.6 first-class 1:N paragraph SPLIT / N:1 MERGE semantics (GetRevisions scoreboard 179/179 GENUINE — deviation catalog EMPTY).** *Internal/experimental (flows through the public `DocxDiff` surface).* A paragraph split mid-text (the user pressed Enter — one before-paragraph becomes N after-paragraphs) or the reverse merge is now an engine capability instead of an inflated whole-paragraph delete+insert account. **Op model (additive only):** two new `IrEditOpKind` members — `SplitBlock` (one `leftAnchor` → ordered N≥2 `splitMergeAnchors`) and `MergeBlock` (the mirror) — with per-member `segmentDiffs` whose slices tile the singular side's token stream exactly (the partition invariant: boundaries are implicit in the diff ops, never stored); the JSON wire gains the two optional arrays on split/merge ops only, so existing scripts serialize byte-identically. N:M stays un-emitted and is rejected by the test-side pairing assert (a `SplitBlock` must carry a null `rightAnchor` — the load-bearing scope ceiling). **Detection:** a containment scan in the aligner's gap fill (`IrBlockAligner.DetectOneToManyInGap`; after similarity pairing, before the 1×1-residue rule; gated by the new `IrDiffSettings.DetectSplitMerge`, default ON) — in-order LCS coverage ≥ 0.90 of the singular paragraph's content tokens across an adjacent candidate window, foreign slack ≤ 0.34, zero-matched-content edge members trimmed (the false-positive guard: an unrelated neighbor insert or edge empty carrier never rides along; an INTERIOR net-new member like WC-1830's inserted math paragraph is absorbed), window length ≤ 8, O(1) content-count prefilter bounds so a fully-rewritten G×G gap never pays LCS cost (adversarial 200×200 fixture stays under its 5s bound). A same-gap Modified pairing whose partner is one segment of the run is PROMOTED into the split; Unchanged/FormatOnly pairs are never candidates (preserves the WC022 identity-reservation reject-order invariant, regression-tested both directions). Thresholds are corpus-swept with a pinned plateau gate (`IrSplitThresholdSweepTests`: the shipped pair sits at the 104-row grid maximum with ≥1 full grid step of margin on every axis). **Surfaces:** the apply-verifier proves `apply(split, [L]) == [R[a..b]]` member-by-member (count/order/reference-identity on the body path; a new produced-anchor-order proof on the cell/note path, asserted corpus-wide); Fine-granularity revisions report each segment's token diff plus one `"\n"` mark revision per added/removed pilcrow; WmlComparer-compatible granularity reproduces the oracle's account (segment-0 inline edits + ONE coalesced inserted region per split) — **WC-1450 and WC-1830, the last two documented deviations, now genuinely PASS (genuine-pass ratchet 177→179; the catalog is empty and asserted so)**; the markup renderer emits the anchored-split shape (N paragraphs, inserted paragraph marks on all but the last via the existing `MarkParagraphMark` primitive; the merge mirror uses deleted marks) with accept ≡ RIGHT / reject ≡ LEFT round-trips proven on both corpus split fixtures + synthetic split/merge shapes, schema-valid. **Fuzzer:** new `SplitParagraph`/`MergeParagraphs` mutation kinds run the own-oracle battery (apply-verify + JSON round-trip + determinism, green at 1000 seeds); they are excluded from the cross-engine differential class because the engines frame a clean split differently by construction (WmlComparer reports the tail as a deleted+reinserted pair of identical text; the IR keeps it Equal and reports the structural mark). The reshuffled seed stream also exposed and fixed a PRE-EXISTING reject-order bug: a deleted block whose nearest preceding paired left was MOVED AWAY restored at the move's destination on reject — moved lefts no longer anchor the deletion interleave, and the builder now stages move-source ops so they interleave with trailing deletions in left-document order. `DetectSplitMerge = false` restores strict 1:1 op semantics. Full suite 2000/0/1; Release build clean. See the "1:N paragraph split / N:1 merge" section of `docs/architecture/ir_diff_engine.md` for the complete implemented algorithm.
- **`DocxDiff` — the IR diff engine's first PUBLIC surface (M2.5 Task 4).** A structure-aware, anchor-addressed DOCX comparison engine, shipped as the public facade `DocxDiff` (`Docxodus/DocxDiff.cs`, `#nullable enable`, fully XML-doc'd). Three entry points: **`Compare(left, right, settings?) → WmlDocument`** (native tracked-changes markup — `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange` — satisfying the WmlComparer contract: accept ≡ right, reject ≡ left), **`GetRevisions(left, right, settings?) → IReadOnlyList<DocxDiffRevision>`** (consumer revisions rendered off the edit script), and **`GetEditScriptJson(left, right, settings?) → string`** (the edit script as data — the differentiator vs `WmlComparer`, which only produces a document or in-memory list). `DocxDiffSettings` mirrors `WmlComparerSettings` defaults (`AuthorForRevisions`, `CaseInsensitive`/`Culture`, `ConflateBreakingAndNonbreakingSpaces`, `WordSeparators`, `DetectMoves`/`MoveSimilarityThreshold`/`MoveMinimumWordCount`) with two honest, documented deviations: `Deterministic` revision dates default **true** (reproducible output; `WmlComparerSettings` is wall-clock by default) and `FormatComparison` defaults **`ModeledOnly`** (reports only modeled-field deltas — false-negative on unmodeled rPr; `Full` for byte-fidelity). A public `RevisionGranularity { Fine (default), WmlComparerCompatible }` governs revision atomization (Fine = engine-native one-per-token-span, byte-stable; compatible = the legacy comparer's coarser contiguous-region grain). `DocxDiffRevision` mirrors `WmlComparerRevision`'s consumer shape (`Type`/`Text`/`Author`/`Date`/`MoveGroupId`/`IsMoveSource`/`FormatChange`) and **adds `LeftAnchor`/`RightAnchor`** (`kind:scope:unid` block anchors interoperable with `DocxSession` and the markdown projection — locate a revision in the document model, or feed it straight to a session mutation). No static/process-global state: `AuthorForRevisions` flows per call, so the surface is multi-author / consolidate-compatible. `WmlComparer` **remains the default/blessed comparison API**; `DocxDiff` ships as a **production-candidate** pending the Word manual-verification checklist + burn-in (decision D4, open). Wraps the internal `Docxodus/Ir/Diff/` pipeline (`IrReader → IrEditScriptBuilder → IrMarkupRenderer/IrRevisionRenderer/IrEditScriptJson`); the internal `IrDiffSettings` stays internal. 15 public-surface smoke tests (each method over a WC corpus pair + a programmatic pair, settings-mapping spot-checks, JSON parse + determinism). See `docs/architecture/ir_diff_engine.md`.
  - **Cross-layer ripple (M2.5 Task 5).** The three entry points are now exposed through every shipping layer, all routing through one shared core facade (`Docxodus/Internal/DocxDiffOps.cs`) so the settings-in / revisions-out JSON wire shapes live in exactly one place (the same single-owner pattern as `HtmlConversionOps`). **WASM:** `wasm/DocxodusWasm/DocxDiffBridge.cs` (`[JSExport]` `Compare` (bytes→bytes), `GetRevisionsJson`, `GetEditScriptJson`), settings passed as a JSON-string parameter; revision/settings DTOs added to `JsonContext.cs`. **npm:** `DocxDiffSettings`/`DocxDiffRevision` types + `DocxDiffRevisionGranularity`/`DocxDiffFormatComparison` enums + the `DocxDiffBridge` slice on `DocxodusWasmExports` (`npm/src/types.ts`), and `docxDiffCompare`/`docxDiffGetRevisions`/`docxDiffGetEditScript` wrappers (`npm/src/index.ts`) following the existing `compareDocuments`/`getRevisions` naming; a 4-test Playwright spec (`npm/tests/docx-diff.spec.ts`) exercises all three over the WC001 fixtures in-browser. **Python:** the stdio host gains the `docx_diff_compare`/`docx_diff_get_revisions`/`docx_diff_get_edit_script` ops (`tools/python-host/Dispatcher.cs`), and `docx-scalpel` ships the matching module-level functions + frozen `DocxDiffSettings`/`DocxDiffRevision`/`DocxDiffFormatChange` dataclasses + `DocxDiffRevisionType`/`DocxDiffRevisionGranularity`/`DocxDiffFormatComparison` enums (`python/src/docx_scalpel/{session,types,enums}.py`). All stateless (two DOCX blobs in, no session) since `DocxDiff` is a pure two-document compare. Verified: full .NET suite (1954/0/1), `build-wasm.sh`, `npm run build` + `tsc --noEmit` clean, pyhost `dotnet build`, `docx-scalpel` import + `mypy` clean, the new Playwright spec (4/4) green.
- **Diff engine — M2.4 Task 4 native renderer completion + the parity GATE (G2: GO).** *Internal/experimental.* `IrMarkupRenderer` now emits the full native OOXML revision vocabulary, and the WmlComparer parity bar is met at **218/218**. **(1) `w:rPrChange`.** FormatChanged token spans render as the RIGHT-side runs (accepted-state formatting) each carrying a `w:rPrChange` whose inner `w:rPr` is the LEFT run's old formatting (recovered positionally from the left source run at the aligned char, since rawLeft==rawRight on that branch); FormatOnly blocks stamp every run likewise. Accept drops the marker (keeps right format); reject swaps to the inner rPr (restores left). The round-trip invariant is **STRENGTHENED to cover format**: in addition to per-block ContentHash, accept restores the RIGHT and reject the LEFT *boundary-normalized modeled-only block format signature* (`IrModeledFormat.BlockSignature`, run-boundary-independent so rPrChange resegmentation never false-fails) — proven over the corpus both directions, 50 fuzz seeds, and a dedicated 30-seed format-mutation class. **(2) Native moves.** MoveBlock/MoveModifyBlock render `w:moveFromRangeStart/End`+`w:moveFrom` (source) and `w:moveToRangeStart/End`+`w:moveTo` (destination) with a deterministic shared `w:name` (`move1`, `move2`, … keyed by `MoveGroupId`); a MoveModify destination nests `w:ins`/`w:del` inside the moveTo range for the in-move edits. `RenderMoves=false` keeps the plain ins/del demotion; a new `IrDiffSettings.SimplifyMoveMarkup` post-pass rewrites moveFrom/moveTo → del/ins and strips range markers (mirrors `WmlComparer.SimplifyMoveMarkupToDelIns`). **`WmlComparer.GetRevisions` run over OUR output recognizes the moves as `Moved`** (the shipped reader is the oracle). **(3) Tables.** A Modified table pair with an `IrTableDiff` renders row/cell-precise markup — EqualRow passthrough, Insert/DeleteRow via `w:trPr/w:ins|w:del` + run wrapping, ModifyRow via nested per-cell block ops (reusing the body dispatch for cell paragraph token diffs). A companion `RevisionProcessor` fix drops a `w:tbl` whose every row is row-deleted instead of leaving an empty-table shell, so a whole inserted/deleted table round-trips at content-hash grain (closes the two table-structural allowlist pairs). **(4) Note scopes.** `IrEditScript.NoteOps` render INSIDE the footnotes/endnotes parts: each note is located by id and its blocks rebuilt from its ops (same dispatch as the body); a wholly-inserted note (or missing note part) is created by cloning the right note wrapper / seeding boilerplate. The invariant extends to footnote/endnote scopes (body-referenced notes). A zero-width-inline dedup (note ref/drawing/tab at a shared token boundary) and a paragraph-mark `w:rPr` schema-order fix (was `AddFirst`, mis-placing it before `w:pStyle`) round out correctness. **THE GATE (decision G2):** the parity bar is the union of two soft-asserted, ratcheting scoreboards — `IrParityScoreboardTests` (**179** GetRevisions rows, floor 179) + the new `IrMarkupParityScoreboardTests` (**39** produced-markup rows — 16 native-move + 3 move-stress + 8 rPrChange + 5 legal-numbering + 6 body-level + 1 parallel-race, floor 39, **39/39 PASS, 0 deviation**) = **218/218**. Thread-safety is proven by construction (no mutable statics; the ParallelRace case runs 16 concurrent renders that each round-trip independently). The `Task4BlockedPairs` round-trip allowlist burned down **11→8 fixtures / 6 distinct root causes**, every one reader/aligner/rId-remap-level and precisely catalogued (note-ref renumber WC-1710 family; SmartArt/drawing diagram rel-id instability WC-1940 family; hyperlink rId collision needing a true remap; body-level bookmark markers) — none a renderer-markup gap. A scoped `MoveRelatedPartsToDestination(skipDanglingRelationships)` flag restores the old engine's loud failure on a dangling rId while the IR caller tolerates it. **G2 recommendation: GO** (gate report: `docs/superpowers/plans/2026-06-11-diff-m24-native-markup-parity.md` `## M2.4 Outcome`, incl. the Word manual-verification checklist + D4 default-engine considerations for M2.5). No public API surface (M2.5).
- **Diff engine — M2.4 Task 3 native OOXML revision renderer core (`IrMarkupRenderer`).** *Internal/experimental.* The IR diff engine now PRODUCES a tracked-changes DOCX, not just a revisions list: `IrMarkupRenderer.Render(IrEditScript, left, right, IrDiffSettings) → WmlDocument` obeys the `WmlComparer` output contract — **accept-all-revisions yields the RIGHT document's content; reject-all yields the LEFT's** — proven against `RevisionProcessor` as the round-trip oracle. **Package base + provenance.** The output is assembled on a clone of the **LEFT** document's package (styles/numbering/fonts/settings/media carry over by reuse, mirroring `WmlComparer.ProduceDocumentWithTrackedRevisions`); only `w:body` is rebuilt, the LEFT's trailing `w:sectPr` preserved, and the RIGHT document's missing styles + numbering copied in (`CopyMissingStyles/NumberingFromOneDocToAnother`, now `internal`). The renderer takes the original `WmlDocument`s (not just the IRs) and re-reads both internally with `RetainSources=true` so it can **clone runs from the source OOXML and split them at token-char boundaries** — preserving every run property including the unmodeled rPr the `IrRunFormat` model omits. **Markup.** `EqualBlock` → right block verbatim; `InsertBlock` → right block with every run wrapped `w:ins` + an inserted paragraph mark (`w:ins` in `w:pPr/w:rPr`); `DeleteBlock` → left block, runs `w:del` (`w:t`→`w:delText`) + deleted paragraph mark; `ModifyBlock` (paragraph) → per-token-span run wrapping (Equal/FormatChanged spans → right runs as-is; Insert → `w:ins`; Delete → `w:del`/`delText`), with a **raw-text guard** that demotes an "Equal-by-match-key" span whose raw bytes differ (NBSP↔space conflation, case folding) to a del+ins pair so the byte-level round-trip still holds. A `w:hyperlink`/`w:sdt` is wrapped from the INSIDE (the schema forbids it as a `w:ins`/`w:del` child). **Revision ids** are a single ascending per-`Render` counter from 1 (NO static state — the `s_MaxId` lesson); author/date from settings (deterministic epoch default). **Conservative Task-4 fallbacks** keep THE INVARIANT holding now: tables get whole-table row+run+paragraph-mark del/ins markup (`w:trPr/w:ins|w:del` appended in schema order), moves render as plain ins+del pairs, FormatChanged spans render as right-side runs without `w:rPrChange` (documented gap — text round-trips via ContentHash; left FORMATTING on reject is the Task-4 item). **THE INVARIANT** (`IrMarkupRendererTests`): `AcceptRevisions(Render)`→IrReader→per-block `ContentHash` sequence EQUALS right's, `RejectRevisions`→EQUALS left's — **165/184 WC corpus round-trips pass both directions** (the 19 remaining are 11 documented Task-4-blocked pairs: footnote/endnote scope markup, opaque SmartArt diagram content, in-paragraph image/math swap, hyperlink-rId-collision remap, table-into-move structural, body-level bookmark/perm markers — each tied to its Task-4 item in a ratcheting allowlist), **50/50 fuzz seeds green**, and **0 NEW OpenXmlValidator schema errors over 81 checked pairs** (vs each input's own baseline, filtering the same tolerated-attribute whitelist `WmlComparer`'s own tests use). A package-relationship guard (`MoveRelatedPartsToDestination` now skips non-part r:ids instead of throwing) and hyperlink/external-relationship import round out media continuity for inserted content. No public API surface (M2.5); native move/format/table/note markup is Task 4.
- **Diff engine — M2.4 Task 2 render-time WmlComparer-compatible granularity + DetectMoves switch.** *Internal/experimental.* The IR diff engine now matches `WmlComparer.GetRevisions`'s coarser revision atomization **at render time only** — the edit script's grain, the aligner, and the token diff are untouched (binding adjudication). A new `IrDiffSettings.RevisionGranularity { Fine (default), WmlComparerCompatible }` governs the `IrRevisionRenderer` projection: **Fine** is the engine-truth one-revision-per-token-span grain (byte-stable; every Fine-mode test passes unchanged); **WmlComparerCompatible** (set by `IrWmlComparerAdapter`) reproduces WmlComparer's contiguous-`w:ins`/`w:del`-region grain via four render-time transforms derived empirically from the failing scoreboard rows + WmlComparer's atomization: **(1)** coalesce adjacent same-kind token-op revisions into one contiguous region, **bridging across an Equal op that is purely separators** between two changed words (a word-bearing Equal is a true boundary); **(2)** word-boundary common-affix trim — trim the shared char prefix/suffix of a region's del/ins text but **backed off so the cut never splits a word in either side** (so `This is a test.`/`This.` → del ` is a test`, while `Title`/`Title1` stay whole and `12,34`/`12,4` are untouched); **(3)** zero-width prune — suppress empty-text Inserted/Deleted from masked textbox placeholder tokens (both modes) and from section-break blocks (compatible mode), while keeping empty-text math/image block revisions WmlComparer counts; **(4)** Choice/Fallback textbox dedup — collapse the adjacent DrawingML/VML duplicate the reader emits per logical textbox (pair-walked so two distinct same-text textboxes are not greedily merged). The **DetectMoves switch** is now a render-time relabel (`IrDiffSettings.RenderMoves`): when the adapter maps `DetectMoves=false`, every aligned move — exact OR fuzzy — renders as an Inserted+Deleted pair instead of a Moved pair (the earlier threshold-push only gated the fuzzy pass and left exact relocations as Moved); compatible mode also demotes a Moved block below `MoveMinimumTokenCount` words to ins+del. `MoveSimilarityThreshold`/`MoveMinimumWordCount` map 1:1 from `WmlComparerSettings`. **Scoreboard: 133→179/179** — **159 PASS + 20 PASS_WITH_DOCUMENTED_DEVIATION** (a new visible scoreboard state for adjudicated engine-level differences that cannot be reconciled at render time without changing the aligner/grain: finer token-diff splits, the aligner's whole-table-delete on an unpaired endnote table, reader textbox duplication in separate cells, and WmlComparer's own French-apostrophe under-report where the IR is more correct — each with a catalogued reason; a stale catalog entry that now passes is flagged). `ParityFloor` ratcheted **133→179** (on PASS + DEVIATION). Differential harness gains a **compatible-mode variant** (same edit script, second classification pass): the TokenSpanGranularity bucket collapses **62→24** and MATCH rises **90→132** vs the Fine pass, reported side by side (assertions stay totality-only). Prelims: the masked-textbox placeholder Delete no longer surfaces as a spurious empty revision (two-textbox-one-removed test) and `IrEditScriptJson` round-trips a hand-built script carrying `IrTextboxDiff`s (record-equality + deterministic re-write). 50-seed fuzzer green; full suite green (Fine-mode byte-stability confirmed). No public API; no produced OOXML markup (that is M2.4 Task 3+).
- **Diff engine — M2.4 Task 1 scope-complete diffing (footnotes, endnotes, textbox interiors).** *Internal/experimental.* The IR diff engine no longer aligns only the body: footnote/endnote scopes and textbox interiors are now diffed, closing the scoreboard's `ScopeGapNewEmpty` family (footnote/endnote/textbox edits that previously produced **zero** revisions). **(1) Note scopes.** `IrEditScriptBuilder.Build` now appends a per-note block diff for every footnote/endnote id present in either side's store, in a deterministic document order (`IrEditScript.NoteOps`): **body ops, then footnotes by note id ascending, then endnotes by note id ascending** — mirroring `WmlComparer.GetRevisions`'s coverage exactly (it diffs the main part + footnotes + endnotes via `GetFootnoteEndnoteRevisionList`; it does **not** diff header/footer scopes, so neither do we — headers were deliberately left out, no test demands them). A matched note runs the full `IrBlockAligner.AlignBlocks` over its two block lists (so a footnote-text edit surfaces as a `ModifyBlock` token diff inside the note); an only-left note becomes all-`DeleteBlock`, an only-right note all-`InsertBlock`. Note anchors carry their own `fn`/`en` scope and are already in the shared `AnchorIndex`, so anchor→block/token resolution and the renderer work unchanged. **(2) Textbox interiors.** A Modified paragraph pair whose textbox placeholder tokens differ now recurses: textboxes pair **positionally** within the paragraph (i-th left with i-th right; a surplus box → all-insert/all-delete), their inner blocks align via `AlignBlocks`, and the nested ops attach as `IrEditOp.TextboxDiffs` (mirroring the `IrTableDiff` nesting on a Modified table). The paragraph's own token diff **excludes** the placeholder-token change (the differ keys on a masked token list whose textbox placeholders share one constant key, so they pair Equal) — no double-reporting. `IrRevisionRenderer` renders both: note-scope ops after the body, textbox-nested ops after the paragraph's own token ops. The JSON writer/reader round-trips `noteOps`/`textboxDiffs`; the apply-verifier reconstructs every matched note's and textbox's right blocks (so the corpus + fuzz own-oracle prove the new scopes). **Scoreboard: 129→133/179** (the WC-1600/1660/1670/1680/2050/2060 footnote/endnote zero-revision rows and the WC-1770 textbox row now pass; the remaining note/textbox rows over/under-report at a finer granularity than the old engine — Task 2 reconciles that). Differential harness: the `ScopeGapNewEmpty` bucket **collapses to 0** (note/textbox pairs move into Match/Granularity). Fuzzer gains an `EditFootnote` mutation (comparable class — WmlComparer diffs notes) and a ~25%-footnote document generator; 50-seed run green with zero regressions. `ParityFloor` ratcheted 129→133. No public API; no produced OOXML markup (that is M2.4 Task 3+).
- **Diff engine — M2.3 Task 4 WmlComparer parity scoreboard (M2.3 close).** *Internal/experimental, test-only.* The standing-USER-DIRECTIVE deliverable: the definitive measurement of how close the IR diff engine is to passing the same tests as the shipped `WmlComparer`, establishing exactly what M2.4 must drive to 100%. **Inventory** of all 8 `Docxodus.Tests/WmlComparer*.cs` files (Theory InlineData rows counted individually; `#if false` dead code excluded): **308 live cases** — categorized by what each asserts (A produced-doc content, B accept/reject round-trip, C GetRevisions counts/types/texts, D move semantics, E format-change details, F consolidate, G settings, H thread-safety, I other) and by M2.4-readiness: **179 RUNNABLE_NOW** (assertable against the IR revisions surface), **39 MARKUP_BLOCKED** (need native OOXML `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange` markup — M2.4), **84 CONSOLIDATE** (`WmlComparer.Consolidate` — out of v1 scope, flagged for user decision), **6 NOT_APPLICABLE** (assert old-engine settings-objects/internals with no IR behavioral meaning — justified each). A test-side **`IrWmlComparerAdapter`** exposes the IR pipeline (`IrReader → IrEditScriptBuilder → IrRevisionRenderer`) through a `WmlComparer`-shaped `GetRevisions(left, right, WmlComparerSettings)`, mapping `AuthorForRevisions`/`CaseInsensitive`/`CultureInfo`/`ConflateBreakingAndNonbreakingSpaces`/`MoveSimilarityThreshold`/`MoveMinimumWordCount` onto `IrDiffSettings` (with `DetectMoves=false` pushing the fuzzy-move threshold above 1.0 — documented as a PARTIAL off-switch since exact-content relocations are still caught by aligner anchoring; `DetailThreshold`/`SimplifyMoveMarkup`/`DateTimeForRevisions` documented unmappable). **`IrParityScoreboardTests`** (Trait `Category=Parity`) ports each RUNNABLE_NOW case's EXACT assertion data (WC003 105 revision-count rows + WC004 56 compare-to-self-⇒-0 rows + WC005 case-insensitive + 3 FormatChange-E details + 1 FormatChange-both + 14 move-detection-D — original tests UNMODIFIED) as soft-asserted Scoreboard rows tagged with the original test id, emitting a PASS/FAIL table + totals and asserting ONLY totality (no crash, every case scored — expected failures are the measurement, not a gate). **Baseline: 129/179 = 72.1% pass** (C 113/161, C+G 1/1, D 12/14, E 3/3). The 50 failures bucket cleanly to the Task-2 triage: 8 `ScopeGapNewEmpty` (footnote/endnote edits the body-only IR path doesn't reach ⇒ 0 revisions) + 13 partial-scope under-reports + 27 `TokenSpanGranularity`/table-cell over-reports (incl. 3 `OldEmpty` where the IR is arguably more correct) + 2 move-via-anchoring. **M2.4's gate criterion is now THE SCOREBOARD AT 100%**; the M2.4 burn-down (priority order) is: (1) non-body scope reading [unlocks 21 RUNNABLE_NOW + the dominant MARKUP_BLOCKED family], (2) native OOXML markup [unblocks all 39 MARKUP_BLOCKED], (3) granularity reconciliation [27 over-reports — partly a user call on whether finer-but-correct counts must match], (4) a real move off-switch [2 cases]. 130 `Ir.Diff` tests; Release green. No public API.
- **Diff engine — M2.3 Task 3 generative fuzzer (`IrDiffFuzzTests`).** *Internal/experimental, test-only.* A deterministic seeded fuzzer over the IR diff pipeline. A test-side `DiffFuzzer` synthesizes from each integer seed a base document (10–40 paragraphs of seeded word-soup with occasional duplicate/boilerplate paragraphs + a ~20% chance of a 2×2 table) and a 1–5-item mutation list (EditWord / InsertParagraph / DeleteParagraph / RelocateParagraph / BoldWord / EditTableCell / InsertRow / DeleteRow), applies the mutations to produce the right document, and emits both via `IrTestDocuments.FromBodyXml`. Determinism is the binding constraint: every choice draws from a seed-constructed `Random` (no clock, no unseeded RNG, no environment input except the seed-COUNT knob), so a seed fully reproduces its (base, mutations, right) on every machine. Per case: **(a) own-oracle invariants — always**: IrReader both sides (RetainSources=false) → `IrBlockAligner` + `IrAlignmentAsserts` totality/per-kind invariants → `IrEditScriptBuilder` → `IrEditScriptVerifier` apply-verification → `IrEditScriptJson` round-trip record-equality + determinism (any failure is a hard test failure dumping the seed + base-paragraph count + mutation list, reproducible via the public `IrDiffFuzzTests.ReproduceCase(seed)` debugger entry point); **(b) differential spot check vs `WmlComparer`** — only for the cross-engine-comparable class (text edits / paragraph insert-delete / table-cell edit / row insert-delete; cases containing Relocate (move-vs-ins+del framing differs) or BoldWord (rPrChange reporting differs) are excluded by construction), compared under the Task-2 combined char-bag equivalence extracted into a shared `RevisionEquivalence` helper (the `Normalize` + per-kind multiset + whitespace-free Inserted+Deleted char-bag contract, now reused by both this fuzzer and the Task-2 harness). A differential mismatch is **not** an automatic failure (the engines legitimately atomize and under-report differently); the test FAILS only on the one asymmetric regression signal — the new engine surfaced ZERO revisions where the old engine saw content (one-sided new-empty) — and otherwise counts + characterizes the mismatch to a gitignored `FuzzArtifacts/` dir. Seed count defaults to 50 (CI) and is overridable via the `DOCXODUS_FUZZ_SEEDS` env var (e.g. 500 nightly); wall-time reported. Outcomes: 50-seed and 500-seed runs both green — own-oracle passed every seed, **zero new-empty regressions**; the 500-seed run produced only 2 characterized non-regression mismatches, both the documented TokenSpanGranularity family (a delete adjacent to the table boundary makes the IR engine attribute a wider block-pairing span — same content, different surrounding context, two-sided). No engine bug found; no public API.
- **Diff engine — M2.3 Task 1 revisions surface (`IrRevisionRenderer`).** *Internal/experimental.* A `WmlComparerRevision`-shaped read-only projection of an `IrEditScript`: `IrRevisionRenderer.Render(script, left, right, settings) → IrNodeList<IrRevision>`. `IrRevision` mirrors the public `WmlComparer.WmlComparerRevision`'s consumer-relevant shape — `IrRevisionType {Inserted, Deleted, Moved, FormatChanged}`, `Text`, `Author`/`Date`, `MoveGroupId`/`IsMoveSource`, and an `IrFormatChangeDetails` (`OldProperties`/`NewProperties` modeled-field dictionaries keyed by WmlComparer-friendly names + `ChangedPropertyNames`, aligned to `WmlComparer.FormatChangeDetails` for adapter friendliness) — plus `LeftAnchor`/`RightAnchor` block anchors as a documented **extension** over `WmlComparerRevision` (an IR-engine addition; an adapter targeting the exact WmlComparerRevision surface ignores them). Mapping in script order: InsertBlock/DeleteBlock → block-level Inserted/Deleted (concatenated raw block text — tables join descendant-paragraph text); ModifyBlock → one revision per token-op span (Insert→Inserted, Delete→Deleted, FormatChanged→**one per maximal uniform `(oldFormat,newFormat)` sub-run**, so a heterogeneous FormatChanged span splits into multiple revisions); MoveBlock → a Moved source+destination pair sharing a `MoveGroupId`; MoveModifyBlock → that pair **plus** the destination's nested token-op revisions emitted immediately after the destination Moved revision (ordering rule: relocate, then describe edits); FormatOnlyBlock → per-position FormatChanged revisions when both sides tokenize to equal counts, else a documented **single whole-block FormatChanged fallback** (the run-boundary word-split case) — with one whole-block empty-details revision when the FormatOnly delta is unmodeled-only (undescribable as an rPrChange); TableDiff recurses (row insert/delete/move → Inserted/Deleted/Moved with row text, cell ops recurse the block machinery); EqualBlock → nothing. New `IrDiffSettings.AuthorForRevisions` (default `"Open-Xml-PowerTools"`, matching `WmlComparerSettings`), `Deterministic` (default **true** — reproducible output is a program principle; inverts `WmlComparerSettings`'s `DateTime.Now` nondeterminism wart), and `DateTimeForRevisions` (pinned `2000-01-01T00:00:00Z` epoch when deterministic; `DateTime.Now`-in-`"o"`-format via `WithWallClockRevisionDate()` otherwise). Output is deterministic (two renders record-equal). Tested per-mapping (incl. MoveModify ordering, heterogeneous sub-run split, FormatOnly fallback, table recursion), author/date settings, determinism, and a WC-corpus totality smoke (all 92 pairs × both directions — every revision has non-null text and resolvable anchors). No public API; no produced OOXML markup (that is M2.4).
- **Diff engine — M2.2 table row/cell granularity + diff-time format-comparison policy (M2.2 close).** *Internal/experimental.* Two additive items finish M2.2. (1) **Nested table diffs.** A Modified table pair now carries a nested `IrTableDiff(IrNodeList<IrRowOp>)` on its `ModifyBlock` op (new `IrEditOp.TableDiff`): `IrTableDiffer` aligns rows by `ContentHash` (a self-contained unique-hash + LIS spine + positional gap fill mirroring the body aligner at row grain — rows have no `FormatFingerprint`, so kinds reduce to `EqualRow`/`ModifyRow`/`InsertRow`/`DeleteRow` + free-off-spine `MovedRow`), pairs cells positionally within `ModifyRow`s (`IrCellOp`), and recurses each differing cell's paragraph blocks through the SHARED block machinery — `IrBlockAligner.AlignBlocks` (new generalized entry point) + `IrEditScriptBuilder.ProjectAlignment` (extracted) — so a **cell-text edit surfaces as a token diff inside that cell, not a whole-table blob**. JSON writer/reader extended (`tableDiff`/`rowOps`/`cellOps`/`blockOps`); the apply-verifier reconstructs table content row-by-row/cell-by-cell and validates row + cell anchors against the actual tables (row/cell anchors are not in `AnchorIndex`). Corpus (forward): 18 tables produce nested diffs — 53 row ops (Equal=26 Modify=20 Insert=1 Delete=6), 39 cell ops, **25 cells carrying a token diff** that M2.1 buried in whole-table Modified blobs. (2) **`IrDiffSettings.FormatComparison = ModeledOnly (default) | Full`** resolves the M2.1 FormatFingerprint run-boundary-noise finding. Diagnosis over WC-BodyBookmarks (the sole source of the corpus' 1,714 FormatOnly entries): every FormatOnly pair is ContentHash-equal with all MODELED run-format fields byte-identical — the only difference is unmodeled-rPr noise (`w:lang`×4597, `w:iCs`×1328, `w:bCs`×550, `w:rFonts` cs faces×33, `w:szCs`/`w:rtl`), all legitimate IR facts but undescribable by `w:rPrChange`-grade reporting. ModeledOnly compares `IrRunFormat` records EXCLUDING `UnmodeledDigest` at the token level and, at the block level, recomputes a **boundary-normalized modeled-only signature** at diff time (the per-token `(MatchKey, modeled-format)` sequence — invariant to the run-resegmentation churn that flips the stored fingerprint) instead of trusting the reader's block `FormatFingerprint`. Purely diff-time: **the IR's stored hashes do not change** (no snapshot churn); no IR normalization rule was added (those facts stay available under `Full`; `w:noProof` is already dropped by N2). Effect under the default: WC-BodyBookmarks **FormatOnly 1714→50**, corpus-wide **FormatOnly 1714→50 / Unchanged 556→2220**; invariants hold both directions, Release build green. No public API.
- **Diff engine — M2.2 similarity gap pairing + cross-gap fuzzy moves (`MovedModified` becomes reachable).** *Internal/experimental.* `IrBlockAligner` now consults `IrDiffSettings` (it was hash-only): three new diff-time settings — `BlockSimilarityThreshold` (0.5; below half token-overlap, "same block edited" is a worse script than Insert+Delete), `MoveSimilarityThreshold` (0.8, mirroring `WmlComparerSettings.MoveSimilarityThreshold`), `MoveMinimumTokenCount` (3, mirroring `WmlComparerSettings.MoveMinimumWordCount`; Word-kind tokens only). A new `IrBlockSimilarity` scorer computes Jaccard over token `MatchKey` **multisets** (non-paragraph blocks score 0 unless `ContentHash`-equal — tables stay out of fuzzy pairing until Task 4), with a per-`Align`-call tokenization cache. Two new aligner passes replace M2.1's blind positional pairing: (1) **in-gap similarity pairing** — after the exact in-order refinement, greedily pair the best-scoring left×right candidates ≥ `BlockSimilarityThreshold` (deterministic ties: smallest left, then right), with a 1×1 unambiguous-residue fallback that keeps a solitary in-place edit Modified regardless of score; leftovers fall out Deleted/Inserted; (2) **cross-gap fuzzy moves** — over the global leftover Deleted×Inserted sets, pairs with ≥ `MoveMinimumTokenCount` Word tokens on both sides and similarity ≥ `MoveSimilarityThreshold` become `MovedModified` (a score-1.0 + equal-`ContentHash` residue classifies as plain `Moved` instead — no edit to re-diff). `MovedModified` now flows through `IrEditScriptBuilder` as a `MoveModifyBlock` source+destination pair whose **destination carries the in-move token diff** (source vs destination tokens) — the relocated-and-edited capability `WmlComparer` cannot express — and the apply-verifier reconstructs the destination from the SOURCE block's tokens. WC corpus drift vs M2.1/T2 (forward): aligner Modified 1488→1419, Inserted 901→970, Deleted 35→104 (69 below-threshold multi-block-gap residues correctly split into Delete+Insert instead of false positional Modifieds); Moved/MoveBlock unchanged (exact moves still caught by anchoring); 0 `MovedModified`/`MoveModifyBlock` on the corpus (it is in-place edits, not relocations-with-edits) — invariants hold both directions, scale guard green. No public API.
- **Diff engine — M2.2 intra-block token diff + edit script (the diff-as-data product).** *Internal/experimental.* Continuing `Docxodus/Ir/Diff/` (all `internal`, `#nullable enable`, WASM-safe, no new dependencies): (1) `IrTokenDiffer` — a Myers O(ND) token diff over `IrDiffToken.MatchKey` with a per-token `IrRunFormat` format-change post-pass, producing `IrTokenDiff` (Equal/Insert/Delete/FormatChanged token-index spans); (2) `IrEditScript` + `IrEditScriptBuilder` — the anchor-addressed, ordered block-level edit script (`EqualBlock`/`FormatOnlyBlock`/`ModifyBlock`/`InsertBlock`/`DeleteBlock`/`MoveBlock`/`MoveModifyBlock`), projecting each `IrBlockAligner` entry to one op (or, for a move, a **source+destination** op pair sharing a deterministic `MoveGroupId` — the source interleaved at its left-anchored unified-diff position, the destination in right order), token-diffing every Modified **paragraph** pair (non-paragraph Modified pairs carry a null `TokenDiff` — table granularity is M2.2 Task 4); (3) `IrEditScriptJson` — a hand-written, deterministic `Write`/`Read` round-trip (token ops as compact 5-element arrays) mirroring `IrDiagnosticJson`. The exit invariant is proven by a test-side `IrEditScriptVerifier` that reconstructs the right body's per-block text from the left IR + script (apply-verification) — green over all synthetic cases and the full 92-pair WC corpus (each script built, apply-verified, and JSON-round-tripped, both directions). `MoveModifyBlock` is wired here but only becomes reachable with the similarity-based fuzzy moves in the Task 3 entry above. No public API.
- **Diff engine — M2.1 tokenizer + block alignment (Phase 2 groundwork).** *Internal/experimental.* New `Docxodus/Ir/Diff/` layer (all `internal`, `#nullable enable`, WASM-safe, no new dependencies): a diff-time `IrDiffTokenizer` over IR paragraphs (word/separator/atomic tokens honoring `WordSeparators`/case/culture as diff settings, hyperlink-target-in-key, transparent field results) and an `IrBlockAligner` that aligns two documents' body block lists into a typed `IrBlockAlignment` via unique-hash `(ContentHash, FormatFingerprint)` anchoring + an LIS spine, with **moves falling out of the alignment by construction** (an exact-content block off the in-order spine is `Moved`, no Jaccard pass) and boilerplate resolving within gaps by order (no global O(n²)). M2.1 closes with a 92-pair WC-corpus alignment smoke (invariants hold forward and reversed), adversarial fixtures (500 near-identical/identical, full-rewrite, contiguous block move), and an anti-O(n²) scale guard. No public API; no renderer yet (the edit script and revision surface are M2.2+).
- **Document IR — M1.5 pre-Phase-2 hardening (textboxes, memory, perf).** *Internal/experimental.* Three additive items prepared the IR for the Phase-2 diff engine; combined with the M1.5 sweep (see Fixed) they lifted corpus markdown byte-equivalence **608/668 → 648/668** and put the IR within perf/memory budget with a sound, hash-complete model:
  - **Textbox bodies (`IrTextbox`).** Textbox bodies (`w:txbxContent` reachable from a
    DrawingML `w:drawing`/`wps:txbx` or a VML `w:pict`/`v:textbox`, including the
    `mc:AlternateContent` Choice/Fallback pair Word emits) are no longer opaque: their
    inner blocks are fully modeled — anchored, hashed
    (`ContentHash`/`FormatFingerprint`), and registered in the document `AnchorIndex` —
    by the normal block walker (depth-capped). A new
    `IrContentHashBuilder.SentinelTextbox` (`0x0B`) folds each inner block's
    `ContentHash` into the *containing* paragraph's hash, closing the diff-engine blind
    spot where textbox text was invisible to `ContentHash`. The emitter mirrors the
    oracle exactly (textbox content stays out of the rendered markdown but its `w:t`
    flows into `TextPreview`/`ScopeHasContent`/cell text, and its inner paragraphs are
    indexed). This was the dominant equivalence lift of the milestone (**608 → 642**),
    closing every textbox fixture and all five header/footer content-detection
    fixtures. Image promotion and textbox modeling are independent.
  - **Optional provenance retention (`RetainSources`).** New
    `IrReaderOptions.RetainSources` (default `true`). When `false`,
    `IrDocument.Sources` is empty and every node's `IrProvenance.Element` is null (a
    shared empty provenance instance, zero per-node allocation), so the parsed
    `XDocument`s become collectible once `IrReader.Read` returns — dropping the largest
    fixture's retained snapshot from **≈11.1× to ≈2.7× the main-part XML size**
    (live-heap delta; reported, not gated). Part-URI facts survive in both modes via
    the additive scope-level `IrScope.PartUri` / `IrCommentStore.PartUri` (the emitter
    prefers these over per-node provenance). Content is provably identical across modes
    — anchors, `ContentHash`, `FormatFingerprint` unchanged (verified corpus-wide + in
    `IrRetentionTests`); diagnostic JSON byte-stable. The Phase-2 diff engine and bulk
    pipelines should read with `RetainSources=false`.
  - **Read perf pass.** `IrReader.Read` no longer runs
    `RevisionProcessor.AcceptRevisions` unconditionally: the default
    `RevisionView.Accept`/`Reject` path skips the full open/clone/walk/re-serialize
    package round-trip on the revision-free majority of documents via the cheap
    in-memory revision-markup scan made sound under Fixed (so "no markup" means a
    byte-identical no-op). This cut the corpus IR-vs-oracle wall-time ratio **~1.94× →
    ~1.16×** (best-of-3, 668 fixtures); the opt-in perf gate (`DOCXODUS_RUN_PERF=1`)
    was tightened 2.0× → 1.5×.

  No public API / WASM / npm / python surface (M1.5 is `internal` by design).
- **Document IR — Phase 1 (M1.1–M1.4) complete** — *internal/experimental.* The
  read-only, immutable, typed, anchor-identified, normalized in-memory DOCX model
  (`Docxodus/Ir/`, all `internal`) is feature-complete through its Phase-1 gate:
  core types + reader (M1.1), normalization rules N1–N15 + `ContentHash`/
  `FormatFingerprint`/`UnmodeledDigest` hashing (M1.2), style/numbering/theme
  registries + lazy effective formats + all scopes (M1.3), and the markdown
  projection ported onto the IR as its validating consumer (M1.4). The IR-path
  markdown emitter reaches **608/668 corpus fixtures byte-equal** with the shipped
  `WmlToMarkdownConverter` (which stays the untouched production oracle); the 60
  remaining divergences are fully triaged — accepted oracle bugs (special-char
  drops, multi-run hyperlink splits, where the IR is *more* correct) plus deferred
  IR work (textbox/shape body content modeled as opaque, and its downstream
  header/footer `ScopeHasContent` detection difference). Phase-1 gate met: perf
  IR-path `Read`+`Emit` 1.90× the oracle's corpus wall time (≤ 2.0× budget,
  `IrMarkdownPerfBudgetTests`, Trait `Perf`); memory ≈11× the largest-body
  fixture's main-part XML retained (measured + reported, not gated); architecture
  doc `docs/architecture/document_ir.md` written. Cutover of the shipped converter
  to the IR path (decision D3) is **deferred to Phase 2** — the IR path ships as a
  CI-validated alternative, not the default. No public API / WASM / npm / python
  surface (Phase 1 is `internal` by design).
- **Document IR markdown emitter — tables, images, section breaks, settings modes (M1.4 Task 2)** —
  *internal/experimental.* `IrMarkdownEmitter` now ports the projection's table
  rendering (simple tables → GFM pipe tables; merges / nesting / over-long cells →
  the opaque ` ```table rows/cols ` block, via the oracle's exact `CanRenderAsGfm`
  simplicity predicate and `CellTextForGfm` escaping), in-paragraph section breaks
  (`{#sec:scope:unid}` + `---` thematic break), the `tbl`/`tr`/`tc`/`sec` anchor-index
  entries (with `TextPreview` parity), and the `AnchorIdRendering`
  (FullUnid/Abbreviated/Sequential, same per-(kind,scope) `AnchorIdMap` construction
  order) and `EmptyParagraphs` settings modes. Images and unmodeled block elements
  project to nothing, matching the oracle (which emits no `w:drawing`/opaque-block
  markup). Two additive IR extensions back the port: `IrInlineImage.Unid` (the source
  `w:drawing`'s `pt:Unid`, equality-neutral) and `IrParagraph.InlineSectionBreakAnchor`
  (the in-pPr `w:sectPr`'s anchor, captured by the reader so the emitter/index can
  reproduce the section transition the body walk's pPr skip otherwise hides). Corpus
  equivalence rises from 205 to 344/668 byte-equal; emitter still never throws.
- **Document IR markdown emitter scaffold + equivalence harness (M1.4 Task 1)** —
  *internal/experimental.* New `IrMarkdownEmitter.Emit(IrDocument, settings)`
  reimplements the markdown projection as an IR consumer (the shipped
  `WmlToMarkdownConverter` stays the byte-untouched oracle), returning a
  `MarkdownProjection`-shaped result (markdown + public `AnchorTarget` index).
  Task-1 scope is BODY paragraphs under DEFAULT settings: headings (`#`-level from
  the pStyle), plain/empty paragraphs (AnchorOnly trim), bulleted list items
  (symbol-glyph → `-`, 2-space-per-ilvl indent), block `{#kind:scope:unid}`
  anchors, inline bold/italic/code/strike with the oracle's exact delimiters and
  escaping, hyperlinks, tabs, and line breaks. Tables/images/opaque blocks,
  multipart scopes, section breaks, numbered-counter markers, heading
  auto-number prefixes, and non-default settings modes are stubbed
  (TODO(M1.4-T2/T3)). A corpus equivalence harness
  (`IrMarkdownEquivalenceTests`, Trait `Corpus`) drives both paths over every
  `TestFiles/*.docx`, compares markdown + body anchor index, writes per-fixture
  diffs to the gitignored `Docxodus.Tests/Ir/EquivalenceArtifacts/`, and asserts
  byte-equality on a curated must-pass list plus per-rule unit tests. Baseline:
  205/668 fixtures byte-equal; emitter never throws (totality).
- **Document IR remaining scopes + comment targets (M1.3)** —
  *internal/experimental.* `IrReader` now honors all `IrScopes` flags and reads
  the header/footer, footnote/endnote, and comment scopes in addition to the body.
  Header/footer parts are enumerated in the same order as the markdown projection
  (`hdr1`/`ftr1`… scope names), each walked by the shared block walker into
  `IrDocument.Headers`/`Footers` with an occurrence kind resolved from the body
  section `w:headerReference`/`w:footerReference`. Footnotes/endnotes populate
  `IrNoteStore` keyed by note id (Word-reserved separator/continuation notes
  skipped via the projection's `IsBoilerplateNote`). Comments populate
  `IrCommentStore` with author/initials/date and blocks, plus N15 comment-range
  *targets*: `w:commentRangeStart`/`End`/`w:commentReference` positions tracked
  during the body walk into per-block `IrCommentTarget(blockAnchor, startChar,
  endChar)` records (visible-`IrTextRun`-char offsets; one target per block for
  cross-block ranges; zero-length target for a reference with no range; orphan
  starts discarded). Comment plumbing stays dropped from the inline stream, so
  body `ContentHash`/`FormatFingerprint` are byte-stable. The diagnostic JSON
  document level is now `{"scopes":[…]}` (body first, then hdr\*/ftr\*/fn/en/cmt);
  all snapshots regenerated (body content/anchors/hashes unchanged modulo the
  wrapper). Corpus totality holds 668/668 reading every scope.
- **Document IR effective-format resolution (M1.3)** — *internal/experimental.*
  New `IrEffectiveFormats(IrDocument)` resolves the *effective* paragraph/run
  format non-destructively by cascading docDefaults → the paragraph/character
  style chain (`basedOn`, applied root-first, cycle-guarded, depth ≤ 16) → direct
  properties, merging per field with later-non-null-wins. `w:rFonts/@w:asciiTheme`
  indirection on a mapped layer resolves through the theme fonts
  (major\*→MajorAscii, minor\*→MinorAscii). Toggle properties are last-writer-wins
  at this fidelity tier (a documented divergence from OOXML toggle-XOR, deferred to
  M1.4+). Style-layer `w:pPr`/`w:rPr` map through the same `IrReader.MapParaFormat`/
  `MapRunFormat` mappers as direct props (refactored to internal statics); the
  effective record's `UnmodeledDigest` is the direct record's. Hash-neutral (no
  reader output change; snapshots byte-stable). Per-style-chain memo cache,
  lock-guarded.
- **Internal Document IR groundwork (M1.1)** — *internal/experimental, no public
  surface.* A typed, normalized, anchor-identified, immutable in-memory model of a
  Word document under `Docxodus/Ir/`: the IR type model (blocks, inlines, formats,
  document/scopes), content-derived SHA-256 hashing (`IrHasher` — `ContentHash` +
  `FormatFingerprint`), and a total body-scope reader (`IrReader.Read`) that
  preserves anything unmodeled as `Opaque` nodes so it never throws on
  weird-but-valid OOXML. Adds a stable, hand-written diagnostic JSON projection
  (`IrDiagnosticJson.Write`, spec §9 — a debugging/test format, **not** a versioned
  contract) plus conformance tests: reader totality over the entire `TestFiles/`
  corpus and golden snapshots over curated fixtures. Groundwork for the planned
  IR/diff-engine program; not referenced by any shipped converter or wrapper yet.
- **Document IR normalization rules (M1.2, partial)** — *internal/experimental.*
  `IrReader` now applies five more §5.2 normalization rules so equality-irrelevant
  OOXML noise stops affecting hashes: **N3** drops `w:bookmarkStart`/`w:bookmarkEnd`
  (at paragraph level and inside runs); **N4** drops `w:lastRenderedPageBreak`
  (layout cache); **N7** maps `w:noBreakHyphen`→U+2011 and `w:softHyphen`→U+00AD as
  text that participates in N5 coalescing; **N8** maps `w:sym` with a parseable hex
  `@w:char` to that BMP code point as text, folding the whole `w:sym` element
  (including `@w:font`) into the run's `UnmodeledDigest` so the glyph font still
  flips the `FormatFingerprint` (unparseable `w:sym` stays `Opaque`); and the
  strip half of **N15** drops comment plumbing (`w:commentRangeStart`/`End`,
  `w:commentReference`) from the inline stream (target-span recording into the
  comment store lands in M1.3). All five previously surfaced as `IrOpaqueInline`
  and perturbed hashes. Golden snapshots regenerated accordingly; block anchors are
  unchanged.
- **Document IR fields & hyperlinks (M1.2, N9 + N14)** — *internal/experimental.*
  `IrReader` now promotes two more constructs from `Opaque` to typed inlines.
  **N14**: `w:hyperlink` → `IrHyperlink` — child runs are walked through the same
  inline pipeline as direct paragraph runs (empty-drop + N5 coalescing within the
  link), an `@r:id` resolves against the main part's hyperlink relationships to the
  external URI (a missing relationship tolerates to `Target=null`), and `@w:anchor`
  internal links use the convention `Target = "#" + anchor` (`InternalTarget`
  bookmark resolution is deferred). The target is bracketed into `ContentHash`
  (sentinels `0x08`/`0x09`), so a target change is a content change and linked text
  is never content-equal to identical plain text; the link's run formats participate
  in the block `FormatFingerprint` in order. **N9**: `w:fldSimple` and complex
  `w:fldChar begin/separate/end` run sequences → `IrFieldRun(Instruction,
  CachedResult)` via a depth-counting field state machine (nested fields flatten
  into the outermost; an unterminated `begin` falls back to opaque losslessly). A
  field contributes only its cached-result bytes to `ContentHash` (no instruction,
  no sentinels), so a `PAGE` field showing "5" is content-equal to a literal "5".
  HC031 golden snapshot regenerated; block anchors unchanged. (The diagnostic JSON
  still renders the new inline kinds as `"unsupported"` until the M1.2 writer task.)
- **Document IR note refs, images & SDT unwrap (M1.2, N12)** —
  *internal/experimental.* `IrReader` promotes two more inline constructs and
  unwraps content controls. Note references (`w:footnoteReference`/
  `w:endnoteReference`) → `IrNoteRef(Kind, NoteId)`; only the kind sentinel
  (`0x05`/`0x06`) feeds `ContentHash` — the note id is positional bookkeeping, so
  renumbering notes never flips a body hash, while footnote and endnote refs stay
  distinguishable. Inline images (a `w:drawing` whose descendant `a:blip` has an
  `@r:embed` resolving to an image part) → `IrInlineImage(PartUri, ImageBytesHash,
  WidthEmu, HeightEmu, AltText)`; the part bytes are SHA-256'd (cached per embed rel
  id so a reused logo hashes once) and `ContentHash` mixes the sentinel `0x07` plus
  that bytes hash, so "same image re-added under a different rel id" is content-equal
  while different bytes diverge. Extent (`wp:extent`) and alt text (`wp:docPr/@descr`
  ?? `@name`) are surfaced but do **not** yet affect `ContentHash` or
  `FormatFingerprint` (a `TODO(M2)` flags surfacing resize as a change). A
  `w:pict` (VML), a drawing without `a:blip@embed`, or a missing/wrong-typed image
  rel falls back to `Opaque` — never throws. **N12**: block-level `w:sdt` (body or
  cell) unwraps to its `w:sdtContent` blocks (each inner `w:p`/`w:tbl` keeps its own
  anchor); inline `w:sdt` and `w:smartTag` (nesting allowed) splice their child runs
  into the paragraph's inline stream and coalesce normally. Formerly-opaque body SDT
  blocks now expose their inner paragraphs/tables with their own anchors; all
  pre-existing non-SDT anchors are unchanged. DB007/HC031/HC042 golden snapshots
  regenerated.
- **Document IR M1.2 complete — diagnostic writer in lockstep + completeness
  guard** — *internal/experimental.* The diagnostic JSON writer now renders the
  four promoted inline kinds with real branches instead of an `"unsupported"`
  fallback: `IrFieldRun` → `{"kind":"field","instruction","cachedResult":[…recursive
  inlines…]}`, `IrHyperlink` → `{"kind":"hyperlink","target"(omitted when
  null),"inlines":[…]}`, `IrNoteRef` → `{"kind":"noteRef","noteKind","noteId"}`, and
  `IrInlineImage` → `{"kind":"image","partUri","imageBytesHash","widthEmu","heightEmu",
  "altText"(omitted when null)}` (`partUri` is the relative part URI — no filesystem
  path leaks). A reflection-driven completeness guard now asserts every concrete
  `IrInline`/`IrBlock` subtype serializes to a known kind (never `"unsupported"`), so
  the writer can no longer drift behind a new reader kind. Every `"kind":"unsupported"`
  disappears from the golden snapshots in favor of typed field/hyperlink/note-ref/image
  objects; all block anchors and content/format hashes are byte-unchanged (the writer
  does not affect hashes). With this, **M1.2 is complete**: normalization rules N3–N15
  (strip-half), typed fields/hyperlinks/note-refs/images, and SDT/smartTag unwrap.
  Also hardened: `ResolveImagePart` narrows its catch to package/IO-shaped exceptions
  (OOM/systemic escape), and SDT/smartTag unwrap recursion (block and inline) is now
  depth-capped at 64 with an opaque fallback beyond the cap (totality without stack
  risk on adversarially-deep nesting).

## [6.4.0] - 2026-05-30

### Added
- **npm `DocxSession` find-by surface** (issue #171). The TypeScript wrapper at
  `npm/src/session.ts` now exposes the six `DocxSession` methods whose bridge
  shells landed in #168 but were unreachable from the typed API: `exists`,
  `findByText`, `findAllByText`, `findByRegex`, and `findByKind` (`replaceMatch`
  was already present). New `FindOptions` type in `npm/src/types.ts`
  (`ignoreCase` / `ignoreWhitespace` / `kindFilter` / `scopes` / `scopeFilter`,
  matching the .NET `FindOptions` record), and the corresponding `Exists` /
  `FindByText` / `FindAllByText` / `FindByRegex` / `FindByKind` signatures added
  to the `DocxSessionBridge` exports interface. Wire shapes are byte-identical to
  what `tools/python-host` consumes, preserving the cross-transport parity
  invariant. Tests: `npm/tests/find-by.spec.ts` exercises the *typed* wrapper
  (via a new `window.Docxodus.openTypedSession` harness helper backed by an
  IIFE-bundled `session.ts`), covering case-sensitive/insensitive text search,
  broad-pattern regex, kind+scope filtering, existence probes, and the
  `grep → replaceMatch` round-trip.

### Fixed
- **Python `FindOptions` scope filtering was a silent no-op.** In `docx_scalpel`, `FindOptions` exposed a single `scope_filter: ProjectionScopes` field and serialized it as an **int** under the `scopeFilter` wire key. But the stdio host (and WASM bridge) parse `scopeFilter` as a **string** (a single named part like `"hdr1"`) and read the coarse `ProjectionScopes` flag set from a separate `scopes` (number) key — which the wrapper never emitted. The net effect: `find_by_text` / `find_all_by_text` / `find_by_regex` ignored any scope restriction passed from Python and always searched all scopes. `FindOptions` now mirrors the .NET record's two distinct controls: `scopes: ProjectionScopes | None` (coarse category flag set → wire `scopes`, int) and `scope_filter: str | None` (fine named-part post-filter → wire `scopeFilter`, string). **API change:** `scope_filter` is now a `str` (was `ProjectionScopes`); callers that want category filtering should use the new `scopes` field (e.g. `FindOptions(scopes=ProjectionScopes.HEADERS | ProjectionScopes.FOOTERS)`). Tests: `python/tests/test_find_options_scopes.py` pins the wire mapping and verifies `scopes=ProjectionScopes.BODY` actually drops a footnote-scope hit on HC031.

## [6.3.0] - 2026-05-30

### Fixed
- **Percentage table/cell widths no longer crash `WmlToHtmlConverter`** (issue #210). `convertDocxToHtml` threw `FormatException` ("Format_InvalidStringWithValue, 100%") whenever a table-level (`w:tblW`) or cell-level (`w:tcW`) width used `w:type="pct"` with a percent-suffixed value such as `w:w="100%"` / `w:w="50%"`. This is the form the `docx` npm library emits for `WidthType.PERCENTAGE`, and it is valid per the OOXML `ST_TblWidth` / `ST_MeasurementOrPercent` schema (which permits either a plain integer in fiftieths-of-a-percent **or** a `"<number>%"` string). The converter cast the attribute straight to `int`, which throws on `"100%"`; DXA (twips) widths were unaffected because they are always plain integers. Width parsing for `w:tblW`/`w:tcW` now goes through a single `ParseTblWidthValue` helper that tolerates the percent-suffixed form: an explicit `"100%"` is treated as a literal percentage, while a bare integer under `pct` is still interpreted as fiftieths of a percent (`5000` → `100%`). Non-numeric/garbage widths are ignored gracefully instead of throwing. Tests: `HcTablePercentageWidthTests` in `Docxodus.Tests/HtmlConverterTablePercentageWidthTests.cs`.

### Added
- **Python DOCX→HTML conversion** — `convert_docx_to_html(data, options)` and
  `DocxSession.to_html(options)` in `docx_scalpel`, backed by a new shared
  `HtmlConversionOps` core renderer that the WASM bridge now also delegates to.
  New `HtmlOptions` dataclass mirrors the existing WASM/npm conversion options.
  New stdio-host ops: `convert_to_html`, `session_to_html`.

## [6.2.0] - 2026-05-28

### Added
- **`WorkerDocxodus.prepare()` — optional comparison-path warmup** (consumer issue JSv4/crowdsourced-redlines-js#2). `createWorkerDocxodus()` warms the .NET WASM runtime but does **not** exercise the comparison engine, so the first `compareDocuments()` pays a one-time warmup cost — comparison-assembly initialization plus JIT of the diff/XML stack — on top of the diff work (~2x the steady-state latency). Consumers worked around this by shipping seed `.docx` fixtures and running a throwaway compare for the side effect. The new optional `prepare(): Promise<void>` pays the cost up front with no caller IO: it runs a complete comparison inside the worker against two tiny seed documents constructed in-memory on the .NET side (no seed fixtures to ship), forcing the full compare path to resolve and JIT. After `await prepare()`, the first real `compareDocuments()` / `compareDocumentsToHtml()` runs at steady-state speed and triggers no further `.wasm` fetches. `prepare()` is never called automatically — skip it and the first compare absorbs the warmup as before. It is idempotent (repeated/concurrent calls share one in-flight warmup and resolve immediately once complete) and concurrent-safe (a `compareDocuments()` issued while a `prepare()` is in flight does not double-load assemblies). Implemented as a new `Warmup()` `[JSExport]` on `DocumentComparer`, a `"prepare"` worker message, and the `WorkerDocxodus.prepare()` proxy method. Tests: `npm/tests/worker-prepare.spec.ts` verifies (via page-level `.wasm` request monitoring + in-worker timing) that after `prepare()` a real compare fetches no additional `.wasm`, a warmed first compare runs ~2x faster than a cold one (~758ms vs ~1504ms), a second `prepare()` resolves in <50ms, and concurrent prepare+compare never double-loads.

## [6.1.0] - 2026-05-28

### Changed
- **`DocxSession.DeleteRange` / `DeleteSection` honor `TrackedChangeMode.RenderInline`** (issue #177). The bulk-delete primitives now produce native Word tracked-deletion markup instead of silently performing a structural delete in tracked mode. Each removed paragraph has every direct-child run wrapped in `w:del` (reusing `WrapRunsInDel`) and the paragraph-mark marked deleted via `w:pPr/w:rPr/w:del` — the combination Word interprets as "this entire paragraph is a tracked deletion", so accepting the change actually removes the block (the old `DeleteBlock`-tracked path left empty paragraphs behind). Tables get `w:trPr/w:del` on every row (Word's row-deletion convention — there is no table-level "delete" markup) plus the same run/paragraph-mark wrapping inside every cell; nested tables recurse. Anchors stay live in the document tree, so the top-level block anchors are reported via `EditResult.Modified` instead of `Removed` — matching `DeleteBlock`'s existing tracked-mode contract. Block kinds outside `w:p`/`w:tbl` (e.g. `w:sdt` content controls appearing mid-range) still fall back to structural removal in tracked mode. No wire-shape changes — the WASM bridge, npm wrapper, and Python stdio host pick up the new behavior automatically through `DocxSessionOps`. Tests: `DS271`–`DS273`.

### Added
- **`DocxSession` block-metadata read surface.** New methods
  `GetBlockMetadata` / `GetBlockMetadatas` / `GetListMembership` /
  `GetSectionInfo` expose paragraph style id+name, outline level, list
  membership (`numId`/`abstractNumId`/`ilvl`/format/start-override/
  inherited-from-style flag), and the enclosing `w:sectPr` (page
  size/orientation/margins/columns/header/footer parts). New types
  `BlockMetadata`, `ListMembership`, `SectionInfo`, and the
  `NumberFormat` enum (`Decimal`/`UpperLetter`/`LowerLetter`/
  `UpperRoman`/`LowerRoman`/`Bullet`). Surfaced in WASM
  (`DocxSessionBridge`), npm (`DocxSession.getBlockMetadata` etc.), and
  Python (`docx_scalpel.session.DocxSession.get_block_metadata` etc.).
- **Markdown projection: list-item classification now follows `pStyle`
  chain.** `WmlToMarkdownConverter` previously labeled a paragraph as
  a list item only when it carried inline `w:numPr`. Now it also walks
  the `pStyle → basedOn` chain (16-level cycle guard) and labels the
  paragraph as a list item if any ancestor style contributes `w:numPr`.
  Brings the projector into agreement with `GetListMembership` for
  style-inherited list items.
- Annotation write surface on `DocxSession` (`AddAnnotation`,
  `RemoveAnnotation`, `UpdateAnnotation`, `MoveAnnotation`) exposed across
  .NET, WASM (`@docxodus/wasm`), and Python (`docx-scalpel`). New
  `EditErrorCode` values: `DuplicateAnnotationId`, `AnnotationNotFound`,
  `EmptyAnnotationSpan`. `EditResult` gained an `AnnotationId` field.
  `AnnotationUpdate` is the new partial-update payload for
  `UpdateAnnotation`. `listAnnotations` now surfaces the `metadata` bag in
  its JSON output (previously omitted).
- **`FillOptions.CoalesceWhitespaceAroundEmptyFill`** (issue #188). New opt-in flag on `DocxSession.FillPlaceholders` that smooths over the canonical template-filling cosmetic foot-gun: returning `""` from the picker (the standard "drop this optional clause entirely" signal) deletes the brackets exactly, leaving surrounding whitespace verbatim. The NVCA Model COI repro `"… on March 14, 2024 [under the name [_______________]]."` with the outer wrapper dropped becomes `"… on March 14, 2024 ."` (note the stray space before the period). When `CoalesceWhitespaceAroundEmptyFill = true`, an empty fill (after `$`-prefix preservation has been applied) absorbs adjacent chars based on the immediate neighbors of the span: whitespace on both sides → collapse to one space; whitespace before + clause-terminating punctuation (`. , ; : ! ?`) after → drop the leading space; matched open/close brackets (`() [] {}`) on either side → drop both. NBSP / narrow NBSP / thin space are folded to ASCII space. Default `false` (preserve current literal-delete behavior). The .NET implementation reads the live flat text of the enclosing block so the rules work regardless of `Boundary` setting; the npm TS implementation uses the match's already-populated `contextBefore` / `contextAfter` (so callers combining `boundary: ContextBoundary.Bracket` with this option won't see the bracket-coalesce rule fire on the JS side — leave `boundary` at the default `Char`). Tests: `DS247a`–`DS247f`.
- **`DocxSession.GetDiff(DiffFormat.Unified | SideBySide)` — line-based diff formats** (issue #178). The two enum values previously reserved as v2-deferred (`NotSupportedException`) now produce real output. `Unified` returns a `patch(1)`-compatible unified diff over the initial-vs-current markdown projections (`--- initial` / `+++ current` headers, 3 lines of context per hunk, hand-rolled `O(n*m)` LCS over `\n`-split lines; empty string when nothing has changed). `SideBySide` returns a `diff -y`-style two-column rendering — the initial projection padded to 72 chars on the left, a one-character marker (`' '` unchanged, `'|'` modified, `'<'` only-initial, `'>'` only-current), then the current projection on the right. Adjacent `Delete + Insert` pairs collapse to a single `|` "modified" row. Decision: hand-rolled LCS over a pulled-in dependency to keep the WASM build AOT-friendly (no reflection-based serializers) and avoid adding a NuGet edge case. The npm wrapper's `getDiff` is now overloaded — `getDiff()` / `getDiff(DiffFormat.Json)` returns `DiffEntry[]`, `getDiff(DiffFormat.Unified | SideBySide)` returns `string`. Tests: `DS289`–`DS289e` (replacing the prior `NotSupportedException` assertion with positive-case coverage of unified hunk shape, side-by-side marker column alignment, insert/delete/modify marker detection, and an out-of-range enum guard); Playwright `edit-summary-and-diff.spec.ts` extended with two end-to-end cases over the WASM bridge.
- **`docx_scalpel.DocxSession.fill_placeholders(picker, options?)` — Python wrapper for the C# template-fill loop** (issue #192, item 1). Mirrors the C# `DocxSession.FillPlaceholders` and the TypeScript `session.fillPlaceholders`: bundles reverse-offset ordering within a paragraph, `$`-prefix preservation (`$[___]` → `$0.20` instead of `0.20`), and multi-pass iteration for nested AlternativeClause brackets — the three foot-guns every template-fill agent was previously re-implementing in ~25 lines of subtle correctness. New `FillOptions` (kinds / scope / max_passes / preserve_dollar_prefix / context_chars / boundary) and `BulkEditResult` (filled / skipped / passes / still_present / unfilled / errors) dataclasses on the public surface. Runs entirely in Python over the existing `find_placeholders` + `replace_match` primitives (no new wire op), matching the TS wrapper's design. The `still_present` field is a post-loop `find_placeholders` count — mirrors the C# `BulkEditResult.StillPresent` field added in #191 for the trustworthy single-call "is the template done?" check. Tests: `python/tests/test_fill_placeholders.py` covering BlankFill replacement, picker-returning-None skip dedup, dollar-prefix on/off, nested-clause multi-pass convergence, max_passes validation, the AlternativeClause-visiting default kinds, and the `still_present == 0` convergence assertion.
- **`BulkEditResult.StillPresent` — trustworthy "is the template done?" metric on `FillPlaceholders`** (issue #189). `BulkEditResult.Skipped` counts placeholders the picker returned `null` for in the first pass that saw them, deduplicated across passes — but this stays `> 0` even when later passes finish the job (e.g. a nested-outer wrapper becomes fillable once its inner is stripped, or a structural delete removes the placeholder entirely). Agents reading `Skipped > 0` after a clean fill ended up cross-checking against `GetEditSummary().RemainingPlaceholders.Count` to know the truth. New additive field `BulkEditResult.StillPresent` is a post-loop `FindPlaceholders(opts.Kinds, opts.Scope).Count` — the metric to assert on for the single-call check. `Skipped > 0 && StillPresent == 0` now correctly reads as "picker skipped on first sight but later passes resolved it." `Skipped` retained for back-compat with its docstring sharpened to direct callers to `StillPresent`. Mirrored on the npm wrapper (`BulkEditResult.stillPresent`) — the TS-side multi-pass loop in `npm/src/session.ts` recomputes via `findPlaceholders`. Tests: `DS247` (multi-pass convergence: `Skipped > 0`, `StillPresent == 0`), `DS248` (picker returns null everywhere: `StillPresent` equals remaining count).
- **`FindOptions.Scopes` (`ProjectionScopes` flag set) + `session.AnchorsByScope`.** The `FindBy*` helpers previously had to default to body and use a string `ScopeFilter` to widen — surveying headers/footers/footnotes meant either passing a magic string like `"hdr1"` (which only matches one part) or walking `Project().AnchorIndex` and filtering by scope name manually. The new `FindOptions.Scopes` field is typed and composable: `Scopes = ProjectionScopes.Headers | ProjectionScopes.Footers` searches every header and every footer in one call. Defaults to `All` so existing callers see no behavior change. The string `ScopeFilter` remains for the rare case of pinning one specific named part (e.g. `"hdr1"` only); it now applies as a finer post-filter on top of `Scopes`. `session.AnchorsByScope(scopes)` is the search-free convenience for the common "enumerate every anchor in scope X" pattern. A new `ProjectionScopesExtensions.IncludesScope(scopeName)` helper exposes the scope-name → flag mapping (`hdr*` → `Headers`, `ftr*` → `Footers`, etc.) for callers that want it directly. Wire shape: `FindOptions` JSON now reads optional `scopes` (number); WASM/Python bridges pick it up automatically. Tests: `DS290`–`DS294`.
- **`DocxSession.CompactRuns(scopes?)` — remove formatting-only run residue.** Public, transactional, scope-aware primitive that removes every `w:r` whose only content is a `w:rPr` (no text, no tabs, no breaks, no field/footnote/comment references). Useful after any workflow that deletes inline content and leaves behind styled-but-empty runs — accepting tracked changes, removing footnotes/comments, run-text refactors. One pre-op snapshot is taken so a single `Undo()` rolls every removal back together; block-level anchors are unaffected because run-level Unids aren't part of the `AnchorIndex`. Defaults to `ProjectionScopes.All` so a call after a body edit also tidies header/footer/footnote/endnote/comment parts; callers that only want body cleanup can pass `ProjectionScopes.Body`. Returns a `CompactResult { RunsRemoved }` so callers can detect "did anything change" without a separate projection round-trip. Tests: `DS295`–`DS298`.
- **`AnchorTarget.AutoNumberPrefix` + `FullText`, mirrored on `AnchorInfo`.** Paragraphs / headings / list items in the body that carry numbering (inline `w:numPr` or numbering inherited from a style) now expose Word's resolved numbering label — `"1."`, `"1.1"`, `"First"`, etc. — as `AutoNumberPrefix` on the projection's `AnchorTarget` and on the `AnchorInfo` returned by `GetAnchorInfo` / `GetAnchorInfos`. `FullText` is a derived convenience that joins prefix + `TextPreview` with a space when a prefix is present. Closes the foot-gun where a caller could see `"# First The total…"` in the markdown projection but a `Grep`/`FindByText` for `"First"` would silently miss it (run text contains only `"The total…"`). The prefix is *not* added to `TextPreview` and is *not* searchable via `Grep` — `Grep` continues to walk run text only — but callers iterating `AnchorIndex` for previews or building search facets now have the rendered label available without re-resolving numbering. Mirrored on the WASM bridge (`MarkdownAnchorTargetDto`, `AnchorInfo` serializers) and the npm wrapper types. Body-only in v1 — header/footer numbering paths aren't routed through `ListItemRetriever` yet. Tests: `DS222`, `DS222a`, `DS222b`.

### Changed
- **Deterministic content-addressable Unids in the markdown projector.** `WmlToMarkdownConverter` now assigns `PtOpenXml.Unid` values via a content-addressable hash (`UnidHelper.AssignToAllElementsDeterministic`) rather than `Guid.NewGuid()`. The Unid is SHA-256(`parent_unid : tag : content_sig : dup_index`) truncated to 32 hex. Properties: same bytes → same Unids across sessions; editing a paragraph's text changes only that paragraph's Unid; inserting a unique-content paragraph anywhere doesn't shift any sibling's Unid; inserting/editing a duplicate-content paragraph shifts `dup_index` of later duplicates only. Closes the cross-session non-determinism foot-gun where a CLI script capturing anchor ids in one run would find them unresolvable in a follow-up run over the same bytes (without `PersistAnchorIds = true`). `WmlComparer` intentionally keeps the random-Guid path (`UnidHelper.AssignToAllElements`) — its matching heuristics expect content-independent Unids, and making them content-addressable inflates the detected revision count on fixtures with same-tag-but-distinct-content elements (verified against `WC003_Compare` on `WC022-Image-Math-Para`). Container elements (those that have block-level descendants) collapse to a tag-name-only signature so editing one block doesn't ripple through the parent's Unid into sibling blocks. Tests: `DS300`–`DS304`.
- **`FillOptions.Kinds` default → `PlaceholderKinds.All`.** The prior default (`BlankFill | Instruction`) silently excluded `AlternativeClause` placeholders, so a picker with `[two]` → `"two"` style bracket-stripping rules would appear to do nothing on those matches — confusing for any caller that wrote a single picker covering every kind it might see. The new default invokes the picker for every kind in the doc; pickers should return `null` for placeholders they don't recognize (the long-standing skip contract). Callers that relied on the prior filter behavior can set `Kinds = PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction` explicitly.

### Changed
- **`docx_scalpel` — `from_wire` decoders renamed to `_from_wire`** (issue #192, item 4). `AnchorTarget`, `TextMatch`, `EditResult`, `TemplatePlaceholder`, and every other value type's JSON-deserializer classmethod is now leading-underscore. The decoders are the stdio transport's JSON-to-dataclass adapter and were never intended for caller use — but their public name made them show up in `dir(ds.AnchorTarget)`, IDE autocomplete, and `help(...)` output, drawing new users into asking what they're for. The public surface is now exactly the dataclass fields plus the user-facing helpers. No back-compat shim — the package is `0.1.0a*` and nothing outside the wrapper itself should have been calling `from_wire`.

### Fixed
- **`DocxSession.GetDiff` — `ArgumentException` on duplicate Unid keys** (issue #187). `ComputeDiff` built its initial/current lookup dictionaries via `AnchorIndex.Values.ToDictionary(t => t.Unid, …)`, which threw the moment two `AnchorTarget`s shared a raw Unid. Two ways this happened in practice: (1) under non-`FullUnid` rendering the `AnchorIndex` is dual-keyed (each target is reachable via its full Unid and its rendered alias), so `Values` enumerated every target twice; (2) the deterministic Unid scheme seeds each scope's root with the root element's local name (`"hdr"` for every header part, `"ftr"` for every footer part), so structurally-identical first paragraphs across multiple header/footer parts hashed to the same raw Unid in different scopes — reproduced on the public NVCA Model COI on the default settings. `ComputeDiff` now keys by `(Anchor.Scope, Unid)` and dedupes the `Values` enumeration with `DistinctBy`, which fixes both paths without changing diff semantics: scope is stable across mutations and the kind-flip case (`p`→`h` via `SetParagraphStyle`) still resolves to the same composite key. Tests: `DS289a` (cross-scope collision smoke), `DS289b` (mutation isolates to the edited scope), `DS289c` (Abbreviated rendering doesn't crash).
- **`tools/python-host/pyhost.csproj` — suppress StyleCop SA1633/SA1636 file-header rules** (issue #173). `dotnet build -c Release tools/python-host/pyhost.csproj` was failing because `Directory.Build.props` sets `TreatWarningsAsErrors=true` for Release and the python-host project inherited the StyleCop ruleset without suppressing the file-header warnings on `Dispatcher.cs` and `Program.cs`. Added `<NoWarn>$(NoWarn);SA1633;SA1636</NoWarn>` to the csproj, matching the existing convention in `wasm/DocxodusWasm/DocxodusWasm.csproj` for tooling/wasm subprojects.
- **`DocxSession.GetDiff` JSON serialization in WASM** (issue #166). `SerializeDiff` originally called `System.Text.Json.JsonSerializer.Serialize(string)` for `anchorId` / `before` / `after` escaping, which uses the reflection-based serializer that the WASM build explicitly disables. Browser callers got `JsonSerializerIsReflectionDisabled` thrown for any non-empty diff (empty `"[]"` short-circuited). Replaced with a hand-rolled `AppendJsonString` helper that mirrors `DocxSessionJson.JsonString`'s escape table. The .NET-side `DS285`/`DS286`/`DS287` tests passed because the standard runtime allows reflection; the Playwright spec from Unit E uncovered the WASM-side breakage.

### Added
- **`DocxSession.ProjectAnchor(anchorId, depth?)`** — project a slice of the document keyed by anchor (one paragraph, a subtree, or a whole heading section) instead of paying the cost of projecting the entire document each time. `ProjectionDepth.SelfOnly` returns just the addressed block, `Subtree` adds descendants, and the default `SubtreeAndFollowingSiblings` extends headings forward through the section bounded by the next same-or-higher heading. The returned `MarkdownProjection.AnchorIndex` is filtered to the scoped Unids only. Useful for showing an LLM one section at a time. Shared core (`DocxSessionOps.ProjectAnchor`) wires the WASM bridge and the Python stdio host; npm wrapper: `session.projectAnchor(anchorId, depth?)` with the `ProjectionDepth` const re-exported. (#167)
- **`WmlToMarkdownConverterSettings.AnchorIdRendering`** — new projection setting controlling how anchor ids appear in `{#…}` tokens. `FullUnid` (default, legacy) keeps the 32-hex-char Unid; `Abbreviated` trims each Unid to the shortest unique prefix per `(kind, scope)` bucket (4-char floor) saving ~5-10% of token budget; `Sequential` replaces Unids with 1-based per-bucket counters in document order — maximally token-efficient for one-shot LLM contexts. The returned `MarkdownProjection.AnchorIndex` is **dual-keyed** in non-`FullUnid` modes: lookups by either full Unid or rendered id resolve to the same `AnchorTarget`, so callers can roundtrip rendered ids straight back to anchor-addressed methods (`DocxSession.ProjectAnchor`, `ReplaceText`, …) without an explicit translation step. `Anchor.Token` continues to return the canonical full-Unid form regardless of rendering mode. Plumbed through the WASM bridge (`MarkdownProjectionSettingsDto.AnchorIdRendering`) and the npm wrapper (`MarkdownProjectionSettings.anchorIdRendering` + exported `AnchorIdRendering` enum). (#167)
- **`GetEditSummary` + `GetDiff` for edit-state introspection** (issue #166). `DocxSession.GetEditSummary()` returns a single `EditSummary` record composing existing primitives — `RemainingPlaceholders` (from `FindPlaceholders`), `BareUnderscoreRuns` (from `Grep`), `TotalAnchors`, `FootnoteCount`, `InlineFootnoteRefCount`, `CommentCount`. Lets verification logic at the end of an edit pipeline be declarative (`Assert.Empty(summary.RemainingPlaceholders)`) instead of a regex zoo. `DocxSession.GetDiff(DiffFormat = Json)` compares the projection captured at session construction time against the current projection and returns an anchor-keyed JSON array of `DiffEntry` records (`op: delete | insert | modify`, `anchorId`, optional `before` / `after`). Gated by new `DocxSessionSettings.CaptureInitialProjection` (default `true`; set `false` to skip the ~200ms upfront cost when you don't plan to diff). `DiffFormat.Unified` and `DiffFormat.SideBySide` are reserved enum values that throw `NotSupportedException` in v1 — see issue #178 for the line-based diff follow-up. `DocxSession.RemainingPlaceholders(kinds)` is a thin discoverability alias for `FindPlaceholders`. Shared core (`DocxSessionOps.GetEditSummary` / `RemainingPlaceholders` / `GetDiff`) propagates to both the WASM bridge and the Python stdio NDJSON host (`get_edit_summary`, `remaining_placeholders`, `get_diff` ops). npm wrapper: `session.getEditSummary()`, `session.remainingPlaceholders(kinds?)`, `session.getDiff(format?)` with `DiffFormat`/`EditSummary`/`DiffEntry` types re-exported. Tests: `DS280`–`DS289`, Playwright `edit-summary-and-diff.spec.ts`.
- **`DeleteRange` and `DeleteSection` for bulk block removal** (issue #165). `DocxSession.DeleteRange(fromAnchorId, toAnchorIdExclusive)` deletes every top-level block-level sibling between two anchors in one call, with one transactional `Undo()` snapshot. Both anchors must share a direct parent and live in the same package part — anchors in different parts return `AnchorsNotAdjacent`, anchors with different parents (e.g. one inside a table cell) also return `AnchorsNotAdjacent`, and `from` not preceding `to` in document order returns `InvalidPosition`. `DocxSession.DeleteSection(headingAnchorId)` is a thin convenience: resolves the heading's level via `WmlToMarkdownConverter.HeadingLevel`, scans forward siblings for the next heading at the same or higher level, and delegates to a shared internal helper. If the target is the last heading in its parent, the section extends to the end. Tracked-change mode is documented as "v1 does structural delete regardless" — wrapping every run across many blocks in `w:del` is deferred until a consumer needs it. Shared core (`DocxSessionOps.DeleteRange` / `DeleteSection`) propagates to both the WASM bridge and the Python stdio NDJSON host (`delete_range`, `delete_section` ops). npm wrapper: `session.deleteRange(fromId, toIdExclusive)`, `session.deleteSection(headingAnchorId)`. Refactor: `WmlToMarkdownConverter.IsHeading` and `HeadingLevel` promoted `private static → internal static` so `DocxSession` can reuse them without duplication. Tests: `DS260`–`DS270`, Playwright `delete-range-section.spec.ts`.
- **`ContextBoundary` enum + widened default `contextChars`** (issue #164). `DocxSession.Grep`, `GrepCrossBlock`, and `FindPlaceholders` now accept a `ContextBoundary` parameter that controls where the context-computation walker stops: `Char` (default, legacy truncate-at-N behavior), `Bracket` (stop at `[`/`]` — the dominant template-fill case for unambiguous per-placeholder context), `Sentence` (stop at `.!?:;`), `Comma` (stop at `,`). Default `contextChars` widened from 40 → 80 across all three methods so plain `.Contains` checks have enough text to disambiguate without the agent dropping into boundary mode. `FillOptions` gains `ContextChars` + `Boundary` fields threaded into the internal `FindPlaceholders` calls. Shared core (`DocxSessionOps.Grep` / `GrepCrossBlock` / `FindPlaceholders`) propagates the new param so both the WASM bridge and the Python stdio NDJSON host pick it up (npm `GrepOptions.boundary`, exported `ContextBoundary` const). Tests: `DS250`–`DS255`, Playwright `context-boundary.spec.ts`.
- **Template-fill convenience — `FillPlaceholders`, `ReplaceInner`, `AlternativeKinds`** (issue #163). `DocxSession.FillPlaceholders(picker, options?)` bundles the three foot-guns every template-fill agent re-implements: reverse-offset ordering across matches within a paragraph, `$`-prefix preservation (`$[___]` → `$0.20` instead of `0.20`), and multi-pass iteration for nested AlternativeClause brackets. Returns a `BulkEditResult` with `Filled` / `Skipped` / `Passes` counts plus per-failure error and unfilled-placeholder lists. New `DocxSession.ReplaceInner(match, newInner)` overload replaces only the bracketed portion of a match, preserving any prefix or suffix outside it — the canonical use case for `$[___]` matches where the regex `\$?\[…\]` captured a leading `$`. `TemplatePlaceholder.AlternativeKinds` is a new additive field listing secondary classifications when the primary `Kind` is borderline (e.g. a long bracketed clause containing a `_______` blank: primary `Kind` stays `BlankFill` for back-compat, with `AlternativeClause` in `AlternativeKinds`). Shared core: `DocxSessionOps.ReplaceInner` (used by both the WASM bridge and the Python stdio host, so `replace_inner` is also exposed via the NDJSON dispatcher); `DocxSessionJson.SerializePlaceholders` emits `alternativeKinds`. npm wrapper: `session.replaceInner(match, newInner)`, `session.fillPlaceholders(picker, options?)` (TS-side mirror of the .NET control loop), new `FillOptions` and `BulkEditResult` types. Tests: `DS230`–`DS233` (ReplaceInner + AlternativeKinds), `DS240`–`DS246` (FillPlaceholders incl. MaxPasses validation + Passes-counter semantics), Playwright `fill-placeholders.spec.ts`.
- **`Docxodus.Internal.{SessionRegistry, DocxSessionOps, DocxSessionJson}` — shared bridge core for `DocxSession` transports.** Lifts the integer-handle pool, the per-op session-lookup + serialization facade, and the StringBuilder JSON helpers that previously lived inside `wasm/DocxodusWasm/DocxSessionBridge.cs` into the core library under `Docxodus/Internal/`. The WASM bridge is now a thin `[JSExport]`-attributed shell over `DocxSessionOps`; a new stdio NDJSON host at `tools/python-host/` (assembly `docxodus-pyhost`) consumes the same facade, so the WASM/TypeScript and stdio/Python clients see byte-for-byte identical JSON wire shapes. `InternalsVisibleTo` for `DocxodusWasm`, `docxodus-pyhost`, and `Docxodus.Tests`. Pure refactor — all 1411 existing tests pass unchanged.
- **`tools/python-host/` — .NET 8 console host for the upcoming python-docxodus wrapper.** Reads NDJSON requests on stdin, dispatches to `DocxSessionOps`, writes NDJSON responses on stdout (diagnostics on stderr). One host process serves many concurrent sessions via the shared handle pool, so an agentic Python pipeline pays the .NET startup cost once and gets µs-to-low-ms per-op latency thereafter. Distinguishes transport-level failures (`ok: false` envelope) from business `EditResult.Success = false` outcomes (`ok: true` envelope carrying the `EditError`). Built for self-contained single-file `dotnet publish` so the eventual pip wheel ships with zero system dependencies.
- **`Exists` / `GetAnchorInfo` / `GetAnchorInfos` / `FindByText` / `FindAllByText` / `FindByRegex` / `FindByKind` exposed on both bridges.** Closes the remaining gap where these public `DocxSession` methods existed in the .NET API but had no wire serializer (so they were unreachable from any non-.NET client). Lands them once in `DocxSessionOps`; the WASM `[JSExport]` shell and the stdio NDJSON dispatcher pick them up automatically. `FindOptions { IgnoreCase, IgnoreWhitespace, KindFilter, ScopeFilter }` on the wire as `{ignoreCase?, ignoreWhitespace?, kindFilter?, scopeFilter?}`. `GetAnchorInfos` bulk lookup follows issue #162's design: `{anchorIds: string[]} → {id: AnchorInfo | null}`. `ReplaceMatch(TextMatch)` is intentionally **not** a wire op — `ReplaceTextAtSpan(anchor, span.start, span.length, replace)` already exposes the underlying primitive; client wrappers implement `replaceMatch(match, replace)` as a 1-line helper rather than ship an 80-line `TextMatch` parser on every transport.
- **Anchor introspection ergonomics — `TextPreview` on `AnchorTarget`, boilerplate footnote filter, `GetAnchorInfos` bulk lookup** (issue #162). `WmlToMarkdownConverter` now computes the first ~80 chars of each block element's flat text during projection and exposes it as `AnchorTarget.TextPreview` — agents no longer need an N-anchor walk via `session.GetAnchorInfo` to surface previews when iterating the `AnchorIndex`. Word-reserved `w:footnote`/`w:endnote` separators (`type="separator"` / `type="continuationSeparator"`) no longer appear in the projection's `AnchorIndex` (they were internal Word plumbing surfaced as un-deletable `fn:fn:*` anchors). New `DocxSession.GetAnchorInfos(IEnumerable<string>)` returns a dictionary mapping each requested id to its `AnchorInfo?` in a single pass; unknown ids map to `null`. WASM bridge surfaces `textPreview` on `MarkdownAnchorTargetDto`, on session `Project()` responses, and on `FindBy*` results; adds `GetAnchorInfo` / `GetAnchorInfos` JSExports. npm wrapper: new `textPreview` field on `MarkdownAnchorTarget`, `AnchorTargetRef`, and `DocxSessionProjection.anchorIndex`; `session.getAnchorInfo(id)` and `session.getAnchorInfos(ids[])` methods. Tests: `MD005` (anchor TextPreview), `MD006` (boilerplate filter), `DS220`–`DS221` (bulk lookup), Playwright `anchor-introspection.spec.ts`.

## [6.0.0] - 2026-05-25

### Fixed
- **`WmlComparer` — defensive null/empty guards on three sibling consumer sites flagged by issue #128.** Follow-up to PR #124, which guarded `FindIndexOfNextParaMark`. The same `cul`-can-contain-`ComparisonUnitGroup` (and empty-descendant) hazard existed in three more places that would have crashed with `NullReferenceException` or `InvalidOperationException` (`.Last()` on empty) had the inputs reached them: `FindCommonAtBeginningAndEnd` (boundary atom dereference), `SplitAtParagraphMark` (paragraph-mark search), and `DoLcsAlgorithm` (last-atom lookup). The producer (`CreateComparisonUnitAtomListRecurse` + `ElementsToThrowAway`) already correctly filters body-level `w:bookmarkStart`/`w:bookmarkEnd`/`w:permStart`/`w:permEnd`/`w:proofErr`, so these guards are belt-and-braces. Adds `WmlComparerBodyLevelElementsTests` with five small programmatic fixtures (bookmarks, perm markers, proof-error markers at body level) that assert `Compare` succeeds — replacing the original 4 MB binary pair's weaker "no NRE" assertion for the body-level case.

### Added
- **`DocxSession.GrepCrossBlock` — text search that may span adjacent paragraphs** (issue #146). Extends `Grep` (#143) so a single match can cross block boundaries among adjacent block-level siblings (paragraphs/headings/list items) under the same direct parent. Returns `CrossBlockMatch` records, each carrying `EnclosingAnchors` (every block the match touches, in doc order) and a per-block `Slices[]` breakdown — every slice has its own `SpanInBlock`, `Fragments`, and `Anchor`, so callers can preserve per-fragment formatting when rewriting. Block boundaries appear in the concatenated text as a single `\n`, so `^`/`$` with `RegexOptions.Multiline` anchor at boundaries and `.` doesn't cross unless `Singleline` is set. Matches are scoped strictly: they never cross OOXML package parts (body → footnote), container boundaries (body → table cell), or non-paragraph siblings (a `w:tbl` between two paragraphs breaks the run). Superset of `Grep`: single-block matches still appear with one `Slice` — callers wanting only cross-block hits can filter `Slices.Count > 1`. Edit semantics deferred (per-slice vs merge vs boundary-preserve has no obviously-right default; file follow-up when a consumer needs it). WASM bridge (`GrepCrossBlock` JSExport) + npm wrapper (`session.grepCrossBlock(pattern, options?)`) + new TS types (`CrossBlockMatch`, `BlockSlice`). Tests: `DS200`–`DS209` + Playwright spec.
- **`DocxSession.FindByAnnotation` / `FindByLabel` / `FindByBookmark` / `ListAnnotations` — annotation-based anchor discovery** (issue #132). Bridges the read-side annotation API (`AnnotationManager`, which persists user labels as `_Docxodus_Ann_{id}` bookmarks + a custom XML metadata part) to the write-side session, so an agent told to "edit the indemnification clause" looks up the annotation by id and immediately gets the `AnchorTarget`s to hand to `ReplaceText` / `Raw.GetXml`. v1 returns every block-level anchor (paragraph/heading/list-item/cell/row/table) whose subtree overlaps the bookmark range, sorted in document order; callers filter by `kind` if they only want text-bearing blocks. `FindByLabel` keys by annotation id so multiple regions sharing one label stay disambiguated. `FindByBookmark` accepts any bookmark name (managed or user-authored) as an escape hatch. Long-lived sessions read annotations directly off the open `WordprocessingDocument` (new `AnnotationManager.GetAnnotations(WordprocessingDocument)` overload) — no byte-level save/reopen per query. WASM bridge (`FindByAnnotation`/`FindByLabel`/`FindByBookmark`/`ListAnnotations` JSExports) + npm wrapper (`session.findByAnnotation/findByLabel/findByBookmark/listAnnotations`) + new TS types (`AnchorTargetRef`, `DocumentAnnotation`). Tests: `DS180`–`DS187`.
- **`DocxSession.FindByText` / `FindAllByText` / `FindByRegex` / `FindByKind` helpers** (issue #137). Thin wrappers over `Grep` and the `AnchorIndex` for the workflows every consumer was reimplementing. \`FindOptions { IgnoreCase, IgnoreWhitespace, KindFilter, ScopeFilter }\` lets one call cover the common variants. \`IgnoreWhitespace\` flows down to \`Grep\`'s \`WhitespaceMode.Normalize\` so a needle with regular spaces hits NBSP-using text. \`FindByKind\` reads the projection's \`AnchorIndex\` directly (no text scan) for "enumerate every heading in the body." Tests: \`DS160\`–\`DS166\`.
- **\`DocxSession.Grep\` accepts \`WhitespaceMode\` for NBSP-tolerant matching** (issue #136). New \`WhitespaceMode { Preserve (default), Normalize }\` enum + \`whitespace\` parameter on \`Grep\`. In \`Normalize\` mode the match runs against a flat text where U+00A0 (NBSP), U+202F (narrow NBSP), and U+2009 (thin space) are folded to ASCII space; substitutions are 1:1 character-for-character so fragment \`Span\` offsets returned in the \`TextMatch\` still address the original positions. A follow-up \`ReplaceMatch\` lands in the right place even though the match was discovered via normalized text. Plumbed through the WASM bridge (\`GrepOptions.whitespace\` numeric flag) + npm wrapper. Tests: \`DS150\`–\`DS152\`.
- **\`DocxSessionSettings.SmartQuotes\`** (issue #140). When true, \`ReplaceText\` / \`ReplaceTextRange\` / \`ReplaceTextAtSpan\` (and \`ReplaceMatch\` by extension) payloads have ASCII \`"\` and \`'\` converted to typographic curly quotes (U+201C/U+201D and U+2018/U+2019) based on context: open quote at the start of the string, after whitespace, or after an open-bracket-like character; close quote elsewhere. Avoids the cosmetic regression where a Bluth-Co fill landed as \`"foo"\` adjacent to surrounding already-curly \`"foo"\` text. Plumbed through the WASM bridge (\`DocxSessionSettings.smartQuotes\` JSON flag) + npm wrapper. Tests: \`DS170\`–\`DS174\`.
- **`DocxSession.ApplyFormatToSubstring(anchor, substring, op)` + `ApplyFormat(TextMatch, op)`** (issue #138). Substring overload finds the first occurrence of the visible text in the anchor's flat text and converts to a `CharSpan` internally — eliminates the offset-arithmetic trap where an auto-number prefix shifts visible-text indices vs run-text indices. The `TextMatch` overload pairs naturally with `Grep`/`ReplaceMatch` for "format the exact match I just found." Distinct from the existing `ApplyFormat(string, CharSpan?, FormatOp)` to keep `(anchor, null, op)` whole-paragraph calls unambiguous to the C# overload resolver. WASM bridge: `ApplyFormatBySubstring` JSExport. npm wrapper: `session.applyFormatBySubstring(anchor, substring, op)` and `session.applyFormatToMatch(match, op)`. Tests: `DS130`–`DS132`.
- **`WmlToMarkdownConverterSettings.EmptyParagraphs` setting — render-mode toggle for empty paragraphs** (issue #135). New `EmptyParagraphMode` enum: `AnchorOnly` (default — bare `{#p:body:UNID}` line, current behavior), `MarkedEmpty` (appends `∅` sentinel so agents can pattern-match), `Suppress` (drops the paragraph entirely + removes it from `AnchorIndex`). Plumbed through the WASM bridge (`MarkdownProjectionSettingsDto.EmptyParagraphs`) and the npm wrapper (`MarkdownProjectionSettings.emptyParagraphs` + exported `EmptyParagraphMode` enum). Tests: `MD030`–`MD032`.
- **`DocxSession.FindPlaceholders` — typed enumeration of template slots** (issue #142). Built on `Grep` (#143); classifies bracketed regions into three kinds an agent treats differently:
  - `BlankFill` — `[___]` or `$[___]` value slots
  - `AlternativeClause` — `[entire clause text in brackets]` optional clauses to keep/strip
  - `Instruction` — `[insert X]`, `[specify Y]`, `[*italicized hint*]` — drafter hints; the inner text is exposed as `Hint` with surrounding asterisks stripped
  Returns `TemplatePlaceholder` records wrapping the underlying `TextMatch` so the caller has anchor, span, fragment list, and surrounding context for each match without a second pass. `PlaceholderKinds` flag enum lets callers narrow (e.g. just `BlankFill`). The complete template-fill workflow now collapses to: `foreach (var p in session.FindPlaceholders(PlaceholderKinds.BlankFill).OrderByDescending(p => p.Match.Span.Start)) session.ReplaceMatch(p.Match, value);` — the 200-line Bluth-Co fill script replaced by five lines. WASM bridge (`FindPlaceholders` JSExport) + npm wrapper (`session.findPlaceholders()`, `PlaceholderKinds` flag exports). Tests: `DS120`–`DS126`. Architecture: see `docs/architecture/docx_mutation_api.md` (FindPlaceholders section).
- **`DocxSession.ReplaceTextRange` — surgical text replacement that preserves run formatting** (issue #139). Built on `Grep` (#143). Three public surfaces:
  - `ReplaceTextRange(anchorId, find, replace, options?)` — finds every literal occurrence of `find` in the anchor's flat text and replaces each with `replace`. Returns one `EditResult` per attempted match.
  - `ReplaceMatch(TextMatch, replace)` — convenience for `Grep` results.
  - `ReplaceTextAtSpan(anchorId, spanStart, spanLength, replace)` — exact-span variant for the template-fill case where five identical `[___]` placeholders in the same paragraph each need a different value (the spans disambiguate; the literal text would not).
  `ReplaceOptions { IgnoreCase, MaxReplacements }`. Replacement text inherits the formatting of the FIRST run the match spanned; middle/trailing runs keep their `w:rPr` but lose the slice of text the match consumed (so the bold formatting on a phrase that got partially overwritten survives for any surviving text). Matches are applied in reverse document order so multi-match-per-paragraph cases don't invalidate each other's offsets, and the whole call records a single undo snapshot. WASM bridge (`ReplaceTextRange` + `ReplaceTextAtSpan` JSExports) + npm wrapper (`session.replaceTextRange()`, `session.replaceMatch(match, replace)`). Tests: `DS110`–`DS119`. Architecture: see `docs/architecture/docx_mutation_api.md` (ReplaceTextRange section).
- **`DocxSession.Grep` — cross-run text search with run-fragment breakdown** (issue #143). The foundational primitive `FindByText`/`ReplaceTextRange`/`FindRegexSpans` (#137/#139/#142) will build on. Searches the flat text of every paragraph/heading/list-item in scope and returns matches in document order, each with the `<w:r>` runs the match spans plus per-fragment formatting (bold/italic/strike/underline/code/color/hyperlink/runStyle). Lets callers rewrite a match in place while preserving each fragment's formatting — the format-preservation problem that the Bluth-Co smoke-test fill hit when collapsing runs. `Grep` accepts standard `RegexOptions` and a `ProjectionScopes` filter (defaults to body), with configurable surrounding-context length. Shared text-map+offset-map helper at `Docxodus/Internal/RunTextMap.cs` so future search/replace work doesn't reinvent the run walker. Public surface: `TextMatch`, `RunFragment`, `RunFormatting`. WASM bridge (`Grep` JSExport) + npm wrapper (`session.grep(pattern, options?)`). Tests: `DS100`–`DS108`. Architecture: see `docs/architecture/docx_mutation_api.md` (Grep section).
- **`DocxSession` — stateful in-memory DOCX mutation API** — The write-side counterpart to `WmlToMarkdownConverter` for agentic editing pipelines. Spec at `docs/architecture/docx_mutation_api.md`. Mutations are keyed by markdown-projection anchor ids; every method returns a typed `EditResult` envelope (no exceptions across the API boundary). Surface:
  - Lifecycle: `new DocxSession(bytes, settings?)`, `Project()`, `Save()`, `Exists()`, `GetAnchorInfo()`, `Undo()`/`Redo()`, `Dispose()`
  - Tier A (text CRUD): `ReplaceText`, `DeleteBlock`
  - Tier B (structural): `InsertParagraph`, `SplitParagraph`, `MergeParagraphs`
  - Tier C (formatting): `ApplyFormat` (whole-paragraph or `CharSpan`), `SetParagraphStyle`, `SetListLevel`, `RemoveListMembership`
  - Tier D (advanced): `ReplaceCellContent`; `Settings.TrackedChanges = RenderInline` makes mutations land as `w:ins`/`w:del`
  - Raw OOXML escape hatch: `session.Raw.GetXml/InsertXml/ReplaceXml` for content the markdown subset can't express (complex tables, math, content controls); optional `Settings.ValidateRawOps` runs `OpenXmlValidator` post-apply with rollback on failure
  - Bounded snapshot undo/redo (default depth 50) over per-part XML clones
  - Markdown payload parser (`Internal/MarkdownPayloadParser`) accepts the projector-symmetric subset (paragraphs, headings, lists, blockquotes, fenced code; bold/italic/code/strike/hyperlinks, escapes) and rejects out-of-subset syntax with typed `EditErrorCode`s (e.g. `TableInsertNotSupported`, `FootnoteRefNotSupported`)
  - WASM `[JSExport]` bridge at `wasm/DocxodusWasm/DocxSessionBridge.cs` with explicit session handles (no JS-side GC observability)
  - npm wrapper at `npm/src/session.ts` exposing `openDocxSession()` and the `DocxSession` class with `Symbol.dispose` support; full type surface in `npm/src/types.ts` (snake_case `EditErrorCode` union, `EditResult`, `AnchorRef`, `CharSpan`, `FormatOp`, `DocxSessionSettings`)
- **Full `WmlToMarkdownConverter` implementation** — Replaces the v5.5.4 scaffold with the complete anchor-addressed Markdown projection described in `docs/architecture/markdown_projection.md`. Covers:
  - Paragraphs and headings (Heading 1–6 + Title/Subtitle, with `HeadingLevelOffset`)
  - Inline runs: bold, italic, code (rStyle/monospace heuristic), strikethrough, hyperlinks (internal + external), Markdown metacharacter escaping
  - Lists with `ListItemRetriever`-resolved numbering ("1.", "1.2.", "a.", bullet); 2-space indent per level; `ResolveNumbering=false` falls back to "-" markers
  - Tables: GFM pipe tables when the shape is simple (no `gridSpan>1` / `vMerge` / nested tables / oversized cells); opaque fenced ` ```table` blocks otherwise; addressable per-cell via `{#tc:body:UNID}` anchors
  - Multipart scopes: `# Headers`/`## hdrN`, `# Footers`/`## ftrN`, GFM-style `[^fn-XXXX]`/`[^en-XXXX]` footnote and endnote references and definitions, `# Comments` list with author/date
  - Tracked-change modes: `Accept` (default), `RenderInline` (`{+ins+}`/`{-del-}`), `StripDeletions`
  - Per-element anchor index reachable via `MarkdownProjection.AnchorIndex` and `AnchorTarget.Resolve(WordprocessingDocument)`
  - WASM `[JSExport] ConvertWmlToMarkdown` and npm `convertWmlToMarkdown` wrapper with TypeScript enums for `ProjectionScopes`, `AnchorRenderMode`, `TableRenderMode`, `TrackedChangeMode`

### Changed
- **`UnidHelper`** — Extracted the `PtOpenXml.Unid` assignment logic out of `WmlComparer` into an internal shared helper so the same code paths are used by both `WmlComparer` and `WmlToMarkdownConverter`. Added `AssignToSelfAndDescendants(XElement)` overload that assigns a Unid to the root unconditionally — used by `DocxSession` when inserting freshly-built block elements that need to be addressable on the next projection.
- **`DocxSession.MergeParagraphs` now inserts a single-space separator** at the seam when both sides end/start with non-whitespace, so merged sentences no longer jam together (`"First." + "Second."` → `"First. Second."` instead of `"First.Second."`). Behavior change for callers that relied on raw concatenation. Regression test: `DS085_MergeParagraphs_InsertsSeparator_WhenBothEndsAreNonWhitespace`.

### Fixed
- **`DocxSession.DeleteBlock` now accepts footnote/endnote/comment anchors and cleans up their in-body references** (issue #133). \`DeleteBlock(footnoteAnchor)\` previously failed with \`AnchorWrongKind\`; the workaround was \`Raw.ReplaceXml\` on each footnote, which left orphan \`<w:footnoteReference w:id=\"X\"/>\` markers in the body and rendered as broken superscript in Word. The op now removes the definition AND every cross-reference pointing at its id, across every projected part of the package — \`w:footnoteReference\` / \`w:endnoteReference\` for fn/en, plus the \`w:commentReference\` + \`w:commentRangeStart\` + \`w:commentRangeEnd\` triple for comments. Empty wrapper runs (a \`<w:r>\` whose only meaningful child was the removed reference) are also stripped to avoid leaving styled-empty spans. Word-reserved fn/en kinds (\`type=\"separator\"\` and \`type=\"continuationSeparator\"\`) are refused with a typed error. Smoke on the NVCA Model COI: 95 of 97 footnotes stripped (2 separators correctly refused), zero orphan references in body, output 25% smaller than input. Tests: \`DS140\`–\`DS143\`.
- **\`DocumentSnapshot\` now captures every projected part, not just MainDocumentPart.** Required for the cross-part \`DeleteBlock\` (fn/en/cmt) above so undo restores both the definition AND the in-body references in one shot. Also fixes a previously-latent bug where \`Save()\` stripping Unids from non-main parts failed to restore them via the snapshot, leaving subsequent ops in the session unable to resolve anchors in headers/footers/footnotes/etc. No public-API change.
- **Markdown projector now resolves style-inherited numbering on headings, matching `WmlToHtmlConverter`** (issue #141). The projector's `ListNumberResolver` was guarding on inline `w:numPr` and short-circuiting for paragraphs whose numbering came from their style (e.g. an NVCA Heading1 with `<w:pStyle val="Heading1"/>` where the Heading1 style declares numPr). `ListItemRetriever` handles both inline and style-level numPr; removing the guard lets it. The NVCA Model COI's "First Article" / "Second Article" headings now project as `# 1. That the name…` / `# 2. That the Board…` instead of `# That the name…` / `# That the Board…`, lining up with the HTML converter's `1.` / `2.` rendering of the same paragraphs. Regression test: `MD033_HeadingNumberPrefix_ResolvesFromStyleLevelNumPr`.
- **`DocxSession.ReplaceText` no longer doubles auto-numbered heading prefixes.** The markdown projector emits resolved numbering inline (`## Fourth The total number…`) so a numbered heading reads as a human would see it. An agent that echoed the visible heading back as its `ReplaceText` payload caused Word to render the prefix twice (`"Fourth Fourth: …"`) because the auto-number from `w:numPr` was still being applied to the new run text. `ReplaceText` now resolves the paragraph's auto-number via the shared `Internal.ListNumberResolver` and strips a matching leading prefix (plus one optional separator: ASCII space, tab, or NBSP) from the payload before parsing. Idempotent when the prefix isn't present. Regression test: `DS091`/`DS091b`.
- **`DocxSession.Raw.ReplaceXml` no longer reports the same anchor in both `Created` and `Removed`.** The documented `Raw.GetXml → mutate → Raw.ReplaceXml` round-trip preserves Unids, but the prior impl unconditionally put `target.Anchor` in `Removed` and re-added the (same-Unid) element to `Created` — so callers pattern-matching on the lifecycle lists saw a phantom delete-then-recreate. Classification is now by Unid set intersection: overlap → `Modified`, old-only → `Removed`, new-only → `Created`. Regression tests: `DS092` (round-trip preserves Unid → `Modified`) and `DS092b` (fresh XML with new Unids → `Removed`/`Created`).
- **`DocxSession.Save` strips the internal `PtOpenXml:Unid` attribute from every part by default.** The projector assigns a Unid to every descendant of every projected scope (\~14 k attributes on the NVCA Model COI), and the prior impl serialized them all — turning a 148 KB input into a 588 KB output (4× bloat). The attribute is internal to the projector and not in the OOXML schema; stripping it on save is the correct default. The escape hatch for callers that need anchor-id stability across save/reopen is `DocxSessionSettings.PersistAnchorIds = true`. Regression tests: `DS093` (default strips), `DS094` (opt-in preserves).

- **`DocxSession` Tier B/C ops now walk `<w:hyperlink>` / `<w:sdt>` / `<w:fldSimple>` / `<w:smartTag>` containers when computing offsets and iterating runs.** The prior implementation iterated only `Elements(W.r)` (direct paragraph children), which caused four interlocking bugs uncovered by smoke-testing the NVCA Model COI:
  - `SplitParagraph` left hyperlinks stuck to the first half regardless of the split offset (so a split at offset 5 of `"Mix of bold ... [link]."` produced `"Mix olink"` + `"f bold ... ."`). Containers crossing the boundary are now split into two siblings sharing the same `r:id`/attributes.
  - `MergeParagraphs` silently discarded hyperlinks / bookmarks / sdts in the second paragraph (only direct `<w:r>` children were moved before `secondEl.Remove()`). All non-`pPr` children are now moved.
  - `ApplyFormat` skipped runs inside hyperlinks and used `ParagraphText` (direct-runs-only) for span validation, while `GetAnchorInfo.TextPreview` summed descendant text — so an agent computing offsets from the markdown projection got `OffsetOutOfRange` on valid spans. Both now share the descendant-walking `InlineRuns` helper, and hyperlink-internal runs are formatted.
  - `ReplaceText` discarded bookmarkStart/End, comment range markers, perm markers, and proofErr because `RemoveNodes()` cleared everything but `pPr`. These markers are now preserved across the replace (pre-content markers wrap before the new runs, post-content markers after). Regression tests: `DS080`-`DS088`.
- **`DocxSession.PromoteHyperlinkRelationships` dedupes by URL.** Each `ReplaceText`/`InsertParagraph` previously called `AddHyperlinkRelationship` unconditionally, so repeated edits with the same link accumulated orphan rIds in `document.xml.rels`. Same-URL ops now reuse the existing relationship. Regression test: `DS089`.
- **`DocxSession.InsertParagraph` reports a `Created[i].Kind` consistent with the next projection.** A bullet payload (`- item`) previously returned `Kind = "li"` even when no `<w:numPr>` was injected, so the returned anchor id (`li:body:…`) never appeared in the projection (`p:body:…` did). Bullet/ordered-item payloads now inherit `<w:numPr>` from a nearest-sibling list item when one exists; the reported kind is computed via the same predicate the projector uses. Regression test: `DS090`.
- **`WmlToMarkdownConverter` projection fidelity** — Surfaced and fixed during smoketesting against the NVCA Model Certificate of Incorporation (a heading-heavy legal document):
  - **Numbered headings keep their auto-number.** A `Heading{1..9}` paragraph that also carries `w:numPr` (the standard legal-doc convention for `FIRST: …` / `1.1 …` clause numbering) now prepends the resolved number to the heading text. Previously the auto-number was silently dropped, leaving headings like `## : The name of this corporation is …`.
  - **`w:sectPr` emits `---` thematic break with anchor.** Section breaks inside a paragraph's `pPr` now produce a `{#sec:scope:UNID}\n---` pair so callers can navigate sections; the trailing top-level `sectPr` (metadata only) is still suppressed in output but registered in `AnchorIndex` for editing.
  - **Inter-scope `---` separators.** A horizontal rule is emitted between adjacent non-empty scope sections (`# Document` / `# Headers` / `# Footers` / `# Footnotes` / `# Endnotes` / `# Comments`) so downstream parsers can split per scope without inspecting heading text.
  - **Heading7-9 preserve depth.** Word styles `Heading7`/`8`/`9` now emit 7/8/9 hashes instead of being silently clamped to `######`. Strict CommonMark renderers will treat 7+ hashes as literal text; LLM consumers and structured parsers can recover the original outline depth.
  - **Empty header/footer scopes are suppressed.** DOCX files commonly declare 6+ header/footer parts for first-page/even-page/default variants and leave the unused ones blank; the projection no longer emits `## hdrN` titles for scopes whose only content is whitespace.
  - **Anchor-only paragraph lines no longer carry a trailing space.** Empty paragraphs (visual spacers in Word) now render as `{#p:body:UNID}\n` instead of `{#p:body:UNID} \n`.

## [5.5.4] - 2026-05-24

### Fixed
- **NullReferenceException in `FindIndexOfNextParaMark` with body-level bookmarks (#124, thanks @papyria)** — `FindIndexOfNextParaMark` assumed all elements in the comparison-unit array were `ComparisonUnitWord`, but documents with `bookmarkStart`/`bookmarkEnd` as direct children of `w:body` produce other `ComparisonUnit` types. Now handles any `ComparisonUnit` with `Contents` (including `ComparisonUnitGroup`) and adds a null guard for the `LastOrDefault()` call.

### Added
- **`WmlToMarkdownConverter` scaffold (#127)** — Public surface for an anchor-addressed markdown projection of Word documents. `Convert(WmlDocument, WmlToMarkdownConverterSettings)` / `Convert(WordprocessingDocument, ...)` return a `MarkdownProjection` (markdown text + anchor index) with anchors of the form `{#kind:scope:unid}` derived from Docxodus' existing Unid system. **Scaffold only** — projection logic ships in subsequent phases. See `docs/architecture/markdown_projection.md` for the spec.

### Maintenance
- **Bump `Microsoft.NET.Test.Sdk` from 18.4.0 to 18.5.1 (#125)**

## [Unreleased] - .NET 8 / Open XML SDK 3.x Migration

### Fixed (npm)
- **TypeScript subpath exports not resolving under `moduleResolution: "node"` (Issue #113)** - Added `typesVersions` fallback to npm package.json so `docxodus/react` and `docxodus/worker` subpath imports resolve types correctly under all TypeScript module resolution modes. Also reordered export conditions to put `types` before `import` per TypeScript requirements.

### Added
- **Incremental annotation overlay API (Issue #106)** - Decouple HTML conversion from annotation projection to avoid full WASM re-conversion
  - `ProjectAnnotationsOntoHtml()` - Project a full annotation set onto already-converted HTML
  - `AddAnnotationToHtml()` - Add a single annotation to existing HTML without re-converting the document
  - `RemoveAnnotationFromHtml()` - Remove a single annotation by ID, unwrapping spans back to plain text
  - `GenerateVisibilityCss()` - Generate CSS to hide/show annotations by label ID for instant toggling
  - `GenerateAnnotationCssString()` - Generate annotation CSS separately for independent management
  - All methods available in .NET, WASM (JSExport), and npm TypeScript wrapper
  - CSS-based label filtering enables responsive toggle without any re-rendering

### Fixed
- **NullReferenceException in FindIndexOfNextParaMark when comparing documents with body-level bookmarks** - `FindIndexOfNextParaMark` assumed all elements in the comparison unit array were `ComparisonUnitWord`, but documents with `bookmarkStart`/`bookmarkEnd` as direct children of `w:body` produce other `ComparisonUnit` types. Now handles any `ComparisonUnit` with `Contents` (including `ComparisonUnitGroup`) and adds a null guard for the `LastOrDefault()` call.
- **Paginated rendering: text clipped at page bottom + inconsistent paragraph spacing (Issue #114)**
  - Fixed `lineRule` default handling: when `w:lineRule` is absent but `w:line` is present, treat as "auto" per OOXML spec (ISO/IEC 29500). Previously the line value was ignored, causing accumulated line-height mismatches that clipped the last line on pages.
  - Fixed `contextualSpacing` handling: now suppresses both `spacingAfter` (margin-bottom) AND `spacingBefore` (margin-top) for consecutive same-style paragraphs. Previously only `spacingAfter` was suppressed, leaving inconsistent inter-paragraph gaps.
  - Fixed pagination engine bottom margin over-reservation: the last block's bottom margin is no longer counted against page space since it's invisible (clipped by `overflow: hidden`). This prevents premature page breaks where content would have been visible.
- **Annotation projection fails on sanitized HTML (Issue #110)** - `ProjectAnnotationsOntoHtml`, `AddAnnotationToHtml`, and `RemoveAnnotationFromHtml` now handle HTML fragments with multiple root elements (e.g., DOMPurify-sanitized output) and HTML named entities (`&nbsp;`, `&ndash;`, etc.)
  - Root cause: `XElement.Parse()` requires valid XML with a single root element; sanitized HTML strips `<html>`/`<body>` wrappers leaving multiple roots
  - Fix: Auto-wraps multi-root HTML in a synthetic container for parsing, unwraps on serialization; replaces common HTML entities with numeric XML equivalents
- **Table container missing top margin (Issue #108)** - Tables preceded by paragraphs with no after-spacing now get a default `margin-top: 7.5pt` for visual separation
  - Also handles floating table spacing from `w:tblpPr` (`topFromText`/`bottomFromText` attributes)
  - Tables preceded by paragraphs with explicit after-spacing correctly skip the default margin
- **Move markup Word compatibility (Issue #96)** - Documents with move operations no longer cause Word "unreadable content" warnings
  - Root cause: `FixUpRevMarkIds()` was overwriting IDs of `w:del`/`w:ins` after `FixUpRevisionIds()` had already assigned unique IDs, causing collisions with move element IDs
  - Fix: Removed redundant `FixUpRevMarkIds()` call - `FixUpRevisionIds()` already handles all revision element IDs correctly
  - Added `SimplifyMoveMarkup` setting to optionally convert move markup to simple `w:del`/`w:ins` if desired
  - Added comprehensive ID uniqueness tests to prevent regression
  - `DetectMoves` now defaults to `true` (move detection is safe to use)
- **Footnote/endnote numbering** - Fixed footnotes and endnotes displaying raw XML IDs instead of sequential display numbers
  - Per ECMA-376, `w:id` is a reference identifier, not the display number
  - Added `FootnoteNumberingTracker` class to scan document and build XML ID → display number mapping
  - Footnotes/endnotes now render with sequential numbers (1, 2, 3...) based on document order
  - Also fixed footnote ordering in the footnotes section to match document order
  - Updated both regular and paginated rendering modes
  - See `docs/ooxml_corner_cases.md` for detailed documentation
- **Legal numbering continuation pattern** - Fixed incorrect multi-level list numbering when items continue a flat sequence at different indentation levels
  - Documents with items like 1., 2., 3. at level 0 followed by item at level 1 (with start=4) now render as "4." instead of "3.4"
  - Added "continuation pattern" detection in `ListItemRetriever.cs` that recognizes when a deeper-level item continues a flat list
  - When detected, uses level 0's format string, run properties, and paragraph properties with the current counter value
  - Fixes underline appearing on continuation items when level 1's rPr has underline but level 0's doesn't
  - Fixes tab/indentation spacing to use level 0's tab stops and indentation for consistency
  - Updated `FormattingAssembler.cs` to use `GetEffectiveLevel()` in paragraph property stack and annotation functions
  - See `docs/ooxml_corner_cases.md` for detailed documentation of this edge case
- **Tab width calculation** re-enabled in `WmlToHtmlConverter` for proper tab stop positioning
  - Previously disabled due to Azure font measurement failures; now uses estimation fallback
  - `MetricsGetter._getTextWidth()` returns character-based estimation when SkiaSharp measurement fails
  - Estimation formula: `charWidth = fontSize * 0.6 / 2` per character (same as WASM builds)
  - Tab positioning now properly accounts for preceding text width
  - Works in Azure, WASM, and environments without fonts installed
  - Added Playwright visual tests for tab rendering verification
- **Thread-safety issues** in `WmlToHtmlConverter` and `FontFamilyHelper` that could cause corruption during concurrent document conversions
  - `ShadeCache` in `WmlToHtmlConverter` now uses `ConcurrentDictionary` for thread-safe shade color caching
  - `FontFamilyHelper._unknownFonts` now uses `ConcurrentDictionary` for thread-safe font tracking
  - `FontFamilyHelper.KnownFamilies` now uses `Lazy<T>` for thread-safe lazy initialization
  - Added `WmlToHtmlConverter.ClearShadeCache()` and `FontFamilyHelper.ClearUnknownFontsCache()` methods for memory management in long-running processes

### Breaking Changes
- **Target Framework**: Changed from net45/net46/netstandard2.0 to .NET 8.0
- **Open XML SDK**: Upgraded from 2.8.1 to 3.2.0
- **Graphics Library**: Replaced System.Drawing with SkiaSharp 2.88.9

### Added
- **Table Width DXA Support** - Tables with DXA (twips) widths now render correctly
  - Previously, only percentage widths were handled; DXA widths were ignored
  - Tables with `w:tblW[@w:type="dxa"]` now render with proper `width: XXpt` CSS
  - Conversion uses standard formula: `dxa / 20 = points`
  - Addresses converter gaps #1 (Table Width Calculation)
- **Borderless Table Detection** - Tables without borders now get semantic markup
  - Tables with `w:tblBorders` set to `nil`/`none` or missing get `data-borderless="true"` attribute
  - Useful for identifying layout tables vs data tables
  - Enables CSS-based styling for signature blocks and multi-column layouts
  - Addresses converter gaps #3 (Borderless Table Detection)
- **Document Language Attribute** - HTML output now includes `lang` attribute for improved accessibility
  - New `DocumentLanguage` setting to manually override the language (default: auto-detect)
  - `<html>` element now includes `lang` attribute (e.g., `<html lang="en-US">`)
  - Language is auto-detected from:
    1. `w:themeFontLang` in document settings
    2. Default paragraph style's `w:rPr/w:lang`
    3. Falls back to "en-US"
  - Foreign text spans get `lang` attribute when different from document default
  - Improves screen reader pronunciation and browser font selection
  - Addresses converter gaps #10 (Document Language Attribute) and #11 (Foreign Text Spans)
- **Improved Font Fallback** - Unknown fonts now get appropriate generic fallback, and CJK text gets language-specific font chains
  - Unknown fonts are classified by name patterns and get proper fallback:
    - Fonts with "sans" pattern → `font-family: 'FontName', sans-serif`
    - Fonts with "mono", "code", "courier" patterns → `font-family: 'FontName', monospace`
    - Other fonts default to serif fallback
  - Fixed Courier New and Lucida Console to include `monospace` fallback (was missing)
  - CJK (Chinese, Japanese, Korean) text gets language-specific font fallback chains:
    - Japanese (ja-JP): `'Noto Serif CJK JP', 'Yu Mincho', 'MS Mincho', ...`
    - Simplified Chinese (zh-hans): `'Noto Serif CJK SC', 'Microsoft YaHei', 'SimSun', ...`
    - Traditional Chinese (zh-hant): `'Noto Serif CJK TC', 'Microsoft JhengHei', 'PMingLiU', ...`
    - Korean (ko): `'Noto Serif CJK KR', 'Malgun Gothic', 'Batang', ...`
  - Addresses converter gaps #13 (Limited Font Fallback) and #14 (No CJK Font-Family Fallback Chain)
- **Theme Color Resolution** - Document theme colors are now resolved to actual RGB values
  - New `ResolveThemeColors` setting (default: true) enables theme color resolution
  - Reads color scheme from `theme1.xml` (`a:clrScheme` element)
  - Supports all 12 theme colors: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
  - Applies `w:themeTint` (lighten toward white) and `w:themeShade` (darken toward black) modifiers
  - Resolves `w:themeColor` in run colors, paragraph shading, cell shading, and fills
  - Falls back to explicit color value if theme color not found
  - Addresses converter gap #6 (Theme Colors Not Resolved)
- **@page CSS Rule** - Optional CSS `@page` rule generation for print stylesheets
  - New `GeneratePageCss` setting (default: false) enables `@page` rule generation
  - Reads page dimensions from `w:sectPr/w:pgSz` and margins from `w:sectPr/w:pgMar`
  - Generates CSS `@page { size: Xin Yin; margin: ... }` rules
  - Supports US Letter, A4, and custom page sizes with proper inch conversions
  - Useful for print stylesheets and PDF generation
  - Addresses converter gap #1 (No Page/Document Setup CSS)
- **Unsupported Content Placeholders** - Visual indicators for content that cannot be fully converted to HTML
  - New `RenderUnsupportedContentPlaceholders` setting (default: false for backward compatibility)
  - Supports these unsupported content types:
    - **WMF/EMF images**: Legacy Windows Metafile formats display `[WMF IMAGE]` / `[EMF IMAGE]`
    - **SVG images**: Scalable Vector Graphics display `[SVG IMAGE]`
    - **Math equations (OMML)**: Office Math Markup displays `[MATH]`
    - **Form fields**: Checkboxes, text inputs, dropdowns display `[CHECKBOX]`, `[TEXT INPUT]`, `[DROPDOWN]`
    - **Ruby annotations**: East Asian text annotations display base text with `[RUBY]` marker
  - Placeholders are styled with CSS (color-coded by type) and include:
    - `data-content-type` attribute for the content type
    - `data-element-name` attribute for the XML element name
    - `title` attribute with descriptive tooltip
  - New TypeScript enum `UnsupportedContentType` for type-safe placeholder identification
  - See `docs/architecture/unsupported_content_placeholders.md` for full documentation
- **External Annotation System** (Issue #57) - Store annotations externally without modifying the DOCX file
  - New `ExternalAnnotationSet` type extends `OpenContractDocExport` with document binding:
    - `documentId`: Unique identifier for the source document
    - `documentHash`: SHA256 hash for integrity validation
    - `createdAt`, `updatedAt`: ISO 8601 timestamps
    - `textLabels`, `docLabelDefinitions`: Label definitions keyed by ID
  - `ExternalAnnotationManager` static class provides core functionality:
    - `ComputeDocumentHash()`: SHA256 hash of document bytes
    - `CreateAnnotationSet()`: Create annotation set from document (wraps OpenContractExporter)
    - `CreateAnnotation()`: Create annotation from character offsets
    - `CreateAnnotationFromSearch()`: Create annotation by text search with occurrence index
    - `FindTextOccurrences()`: Find all occurrences of text in document
    - `Validate()`: Validate annotations against document (hash check + text verification)
    - `SerializeToJson()` / `DeserializeFromJson()`: JSON serialization
  - `ExternalAnnotationProjector` for HTML projection:
    - `ProjectAnnotations()`: Post-process HTML to wrap annotated text with styled spans
    - `ConvertWithAnnotations()`: Combined conversion + projection
    - Supports annotation labels (Above, Inline, Tooltip, None modes)
    - CSS generation with customizable class prefix
  - TypeScript/npm wrapper functions:
    - `computeDocumentHash()`: Get document hash for validation
    - `createExternalAnnotationSet()`: Create annotation set from DOCX
    - `validateExternalAnnotations()`: Validate annotations against document
    - `convertDocxToHtmlWithExternalAnnotations()`: Convert with annotations projected
    - `searchTextOffsets()`: Search for text occurrences in document
    - `createAnnotation()`, `createAnnotationFromSearch()`, `findTextOccurrences()`: Client-side helpers
  - Full type definitions: `AnnotationLabel`, `ExternalAnnotationSet`, `ExternalAnnotationValidationResult`, etc.
  - 21 unit tests covering hash computation, annotation creation, validation, serialization, and projection
- **OpenContracts Export Format** (Issue #56) - Export documents to OpenContracts format for interoperability
  - New `OpenContractExporter.Export()` method for complete document export:
    - `title`: Document title from core properties
    - `content`: Complete document text (paragraphs, tables, headers, footers, footnotes, endnotes)
    - `description`: Optional document description
    - `pageCount`: Estimated page count
    - `pawlsFileContent`: PAWLS-format page layout with token positions
    - `docLabels`: Document-level labels
    - `labelledText`: Annotations including structural elements (sections, paragraphs, tables)
    - `relationships`: Parent-child relationships between annotations
  - Full text extraction ensures 100% text coverage:
    - Main body paragraphs and tables
    - Nested tables
    - Headers and footers
    - Footnotes and endnotes
    - Content controls (structured document tags)
  - PAWLS (Page-Aware Layout Segmentation) format for layout data:
    - Page boundary information (width, height, index)
    - Token positions (x, y, width, height, text)
    - Supports annotation targeting by character offset
  - Structural annotations automatically generated:
    - Section annotations with page dimensions
    - Paragraph annotations with text spans
    - Table annotations with content ranges
    - Parent-child relationships (section contains paragraphs)
  - TypeScript API: `exportToOpenContract()` function with full type definitions
  - WASM export: `DocumentConverter.ExportToOpenContract()`
  - Compatible with OpenContracts ecosystem for document analysis
  - **New CLI tool: `docx2oc`** - Command-line tool for OpenContracts export
    - Usage: `docx2oc <input.docx> [output.json]`
    - Default output: same filename with `.oc` extension
    - Installable as .NET tool: `dotnet tool install --global Docx2OC`
- **ReadyToRun and AOT Compilation** - Performance optimizations to reduce cold-start times
  - .NET library: Added `PublishReadyToRun` for pre-compiled native code during publish
  - WASM: Added `RunAOTCompilation` for Release builds to pre-compile IL to WebAssembly
  - Eliminates JIT warmup overhead (~180ms savings on first conversion in .NET)
  - Provides consistent performance with no JIT variance in WASM
- **Lightweight WASM Image Handling** - Images are now embedded as base64 data URIs without SkiaSharp native library
  - Removed SkiaSharp native WASM dependency (~15MB+ savings in bundle size when native lib excluded)
  - Images are passed through directly from DOCX using `ImageBytes` property
  - Dimensions come from document markup (EMUs), not image decoding
  - Browser natively decodes image formats (PNG, JPEG, GIF, etc.)
  - Fallback handling: If SkiaSharp decode fails, images still work via raw bytes
  - Added image handling tests for documents with embedded and hyperlinked images
- **Frame Yielding for UI Responsiveness** (Issue #44 Phase 1) - WASM operations now yield to the browser before heavy work begins
  - All async functions in the npm wrapper (`convertDocxToHtml`, `compareDocuments`, `compareDocumentsToHtml`, `getRevisions`, `addAnnotation`, `addAnnotationWithTarget`, `getDocumentStructure`) automatically yield using double-`requestAnimationFrame` pattern
  - This allows React state updates (loading spinners, progress indicators) to paint before blocking WASM execution
  - Transparent to consumers - no API changes required
  - Gracefully skipped in non-browser environments (Node.js, SSR)
- **Web Worker Support for Non-blocking Operations** (Issue #44 Phase 2) - Fully non-blocking WASM execution via Web Workers
  - New `docxodus/worker` export provides worker-based API: `import { createWorkerDocxodus } from 'docxodus/worker'`
  - Worker API mirrors main API: `convertDocxToHtml`, `compareDocuments`, `compareDocumentsToHtml`, `getRevisions`, `getVersion`
  - Main thread remains fully responsive during WASM execution - animations continue, user interactions work
  - Zero-copy transfer of document bytes via Transferable for optimal performance
  - Worker can be terminated when no longer needed
- **Document Metadata API for Lazy Loading** (Issue #44 Phase 3) - Fast metadata extraction without full HTML rendering
  - New `getDocumentMetadata()` function returns document structure information:
    - `sections`: Array of section metadata with page dimensions and content ranges
    - `totalParagraphs`, `totalTables`: Document-wide content counts
    - `hasFootnotes`, `hasEndnotes`, `hasComments`, `hasTrackedChanges`: Feature detection
    - `estimatedPageCount`: Heuristic-based page count estimation
  - Section metadata includes:
    - Page dimensions: `pageWidthPt`, `pageHeightPt`, `marginTopPt`, etc. (all values in points, 1pt = 1/72 inch)
    - Content area: `contentWidthPt`, `contentHeightPt`
    - Header/footer heights: `headerPt`, `footerPt`
    - Content tracking: `paragraphCount`, `tableCount`, `startParagraphIndex`, `endParagraphIndex`
    - Header/footer presence: `hasHeader`, `hasFooter`, `hasFirstPageHeader`, `hasEvenPageHeader`, etc.
  - Available in main API, worker API, and raw WASM: `DocumentConverter.GetDocumentMetadata()`
  - Enables efficient lazy loading for paginated document viewing
  - Security: Maximum document size limit of 100MB to prevent memory exhaustion
  - Graceful handling of malformed documents and invalid header/footer references
  - Known limitation: Section breaks inside tables or text boxes are not detected (see #51)
- **Page Range Rendering for Virtual Scrolling** (Issue #31 Phase 4) - Render specific page ranges for lazy loading
  - New `RenderPageRange()` method in `WmlToHtmlConverter` renders only specified pages
  - Page-to-block mapping uses heuristic-based estimation (paragraphs and tables per page)
  - HTML output includes pagination metadata via data attributes:
    - `data-start-page`, `data-end-page`: Requested page range
    - `data-total-pages`: Total estimated pages in document
    - `data-start-block`, `data-end-block`: Block index range for rendered content
    - `data-block-index`: Per-element block indices for tracking
  - WASM exports: `DocumentConverter.RenderPageRange()`, `DocumentConverter.RenderPageRangeFull()`
  - TypeScript wrapper: `renderPageRange()` with full options support
  - Worker proxy support: `WorkerDocxodus.renderPageRange()` for non-blocking execution
  - React components for virtual scrolling:
    - `useVirtualPagination` hook: Manages viewport-aware page loading with IntersectionObserver
    - `VirtualPaginatedDocument` component: Auto-renders visible pages plus configurable buffer
  - All existing converter options supported (tracked changes, comments, headers/footers, etc.)
  - Graceful handling of out-of-bounds page requests (internally clamped to valid range)
- **Custom Annotations** - Full support for adding, removing, and rendering custom annotations on DOCX documents
  - `AnnotationManager` class for programmatic annotation CRUD operations:
    - `AddAnnotation()`: Add annotation by text search or paragraph range
    - `RemoveAnnotation()`: Remove annotation by ID
    - `GetAnnotations()`: Retrieve all annotations from a document
    - `GetAnnotation()`: Get a specific annotation by ID
    - `HasAnnotations()`: Check if document has any annotations
  - `DocumentAnnotation` class with properties:
    - `Id`: Unique annotation identifier
    - `LabelId`: Category/type identifier for grouping
    - `Label`: Human-readable label text
    - `Color`: Highlight color in hex format (e.g., "#FFEB3B")
    - `Author`: Optional author name
    - `Created`: Optional creation timestamp
    - `Metadata`: Custom key-value pairs
  - `AnnotationRange` class for specifying annotation targets:
    - `FromSearch(text, occurrence)`: Find text by search
    - `FromParagraphs(start, end)`: Span paragraph indices
  - **Document Structure API** for element-based annotation targeting:
    - `DocumentStructureAnalyzer.Analyze()`: Returns navigable tree of document elements
    - `DocumentElement` class with path-based IDs (e.g., `doc/p-0`, `doc/tbl-0/tr-1/tc-2`)
    - Supported element types: `Document`, `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `TableColumn`, `Hyperlink`, `Image`
    - `TableColumnInfo` for virtual column elements (columns aren't real OOXML elements)
  - `AnnotationTarget` class with flexible targeting modes:
    - `Element(elementId)`: Target by element ID from structure analysis
    - `Paragraph(index)`, `ParagraphRange(start, end)`: Target by paragraph index
    - `Run(paragraphIndex, runIndex)`: Target specific run
    - `Table(index)`, `TableRow(tableIndex, rowIndex)`: Target tables/rows
    - `TableCell(tableIndex, rowIndex, cellIndex)`: Target specific cell
    - `TableColumn(tableIndex, columnIndex)`: Metadata-only column annotation
    - `TextSearch(text, occurrence)`: Search text globally
    - `SearchInElement(elementId, text, occurrence)`: Search within specific element
  - WASM methods: `GetDocumentStructure()`, `AddAnnotationWithTarget()`
  - TypeScript helper functions: `findElementById()`, `findElementsByType()`, `getParagraphs()`, `getTables()`, `getTableColumns()`
  - TypeScript targeting factories: `targetElement()`, `targetParagraph()`, `targetTableCell()`, etc.
  - React `useDocumentStructure` hook with structure navigation helpers
  - Annotations stored as Custom XML Part in DOCX (non-destructive)
  - Bookmark-based text range marking for precise positioning
  - HTML rendering with configurable label modes:
    - `AnnotationLabelMode.Above`: Floating label above highlight
    - `AnnotationLabelMode.Inline`: Label at start of highlight
    - `AnnotationLabelMode.Tooltip`: Label shown on hover
    - `AnnotationLabelMode.None`: Highlight only, no label
  - New settings in `WmlToHtmlConverterSettings`:
    - `RenderAnnotations`: Enable/disable annotation rendering
    - `AnnotationLabelMode`: Select label display mode
    - `AnnotationCssClassPrefix`: Customize CSS class names (default: "annot-")
    - `IncludeAnnotationMetadata`: Include metadata in HTML data attributes
  - WASM/npm support:
    - `getAnnotations()`, `addAnnotation()`, `removeAnnotation()`, `hasAnnotations()` functions
    - `Annotation`, `AddAnnotationRequest`, `AddAnnotationResponse`, `RemoveAnnotationResponse` types
    - `AnnotationLabelMode` enum
    - `ConversionOptions` extended with annotation rendering options
  - React support:
    - `useAnnotations` hook for annotation state management
    - `AnnotatedDocument` component with click/hover event handling
    - `useDocxodus` hook extended with annotation methods
  - 20 .NET unit tests and 21 Playwright browser tests for full coverage (including 11 for element-based targeting)
- **Comment Rendering in HTML Converter** - Full support for rendering Word document comments in HTML output
  - `CommentRenderMode` enum with three rendering modes:
    - `EndnoteStyle` (default): Comments rendered at end of document with bidirectional anchor links
    - `Inline`: Comments rendered as tooltips with `title` and `data-comment` attributes
    - `Margin`: Comments positioned in a flexbox-based margin column alongside content, with author/date headers and back-reference links
  - New settings in `WmlToHtmlConverterSettings`:
    - `RenderComments`: Enable/disable comment rendering
    - `CommentRenderMode`: Select rendering mode
    - `CommentCssClassPrefix`: Customize CSS class names (default: "comment-")
    - `IncludeCommentMetadata`: Include author/date in HTML output
  - Comment highlighting with configurable CSS classes
  - Full comment metadata support (author, date, initials)
  - Margin mode includes print-friendly CSS media queries
  - WASM/npm support via `commentRenderMode` parameter and TypeScript `CommentRenderMode` enum
- **WebAssembly NPM Package** (`docxodus`) - Browser-based document comparison and HTML conversion
  - `wasm/DocxodusWasm/` - .NET 8 WASM project with JSExport methods
  - `npm/` - TypeScript wrapper with React hooks
  - Full document comparison (redlining) support in the browser
  - DOCX to HTML conversion
  - React hooks: `useDocxodus`, `useConversion`, `useComparison`
  - Build script: `scripts/build-wasm.sh`
- **Native Move Markup in WmlComparer** - Produces Word-native move tracking markup (`w:moveFrom`/`w:moveTo`)
  - Compared documents now contain proper OpenXML move elements, not just `w:del`/`w:ins`
  - Move pairs linked via `w:name` attribute for Word compatibility
  - Range markers (`w:moveFromRangeStart`/`w:moveFromRangeEnd`, `w:moveToRangeStart`/`w:moveToRangeEnd`) properly paired
  - Microsoft Word shows moves in "Track Changes" panel as relocated content
  - New `Moved` value in `WmlComparerRevisionType` enum
  - New properties on `WmlComparerRevision`: `MoveGroupId` (links source/destination), `IsMoveSource` (true=from, false=to)
  - New settings in `WmlComparerSettings`:
    - `DetectMoves`: Enable/disable move detection (default: true)
    - `MoveSimilarityThreshold`: Jaccard similarity threshold 0.0-1.0 (default: 0.8)
    - `MoveMinimumWordCount`: Minimum words to consider for move (default: 3)
  - Uses word-level Jaccard similarity for accurate matching
  - Respects `CaseInsensitive` setting for similarity comparison
  - Full WASM/npm support with new TypeScript helpers:
    - `RevisionType.Moved` enum value
    - `isMove()`, `isMoveSource()`, `isMoveDestination()` type guards
    - `findMovePair()` function to find linked move revisions
    - `moveGroupId` and `isMoveSource` properties on `Revision` interface
- **Format Change Detection in WmlComparer** - Detects and tracks formatting-only changes (`w:rPrChange`)
  - When text content is identical but formatting changes (bold, italic, font size, etc.), produces native Word format change markup
  - Compared documents now contain `w:rPrChange` elements that Microsoft Word recognizes in Track Changes
  - New `FormatChanged` value in `WmlComparerRevisionType` enum
  - New `FormatChange` property on `WmlComparerRevision` with:
    - `OldProperties`: Dictionary of original formatting properties
    - `NewProperties`: Dictionary of new formatting properties
    - `ChangedPropertyNames`: List of what changed (e.g., "bold", "italic", "fontSize")
  - New setting in `WmlComparerSettings`:
    - `DetectFormatChanges`: Enable/disable format change detection (default: true)
  - Full WASM/npm support with new TypeScript helpers:
    - `RevisionType.FormatChanged` enum value
    - `isFormatChange()` type guard
    - `FormatChangeDetails` interface with `oldProperties`, `newProperties`, `changedPropertyNames`
    - `formatChange` property on `Revision` interface
- **Improved Revision API** - Better TypeScript support for the `getRevisions()` API
  - `RevisionType` enum with `Inserted`, `Deleted`, and `Moved` values for type-safe comparisons
  - `isInsertion()`, `isDeletion()`, `isMove()`, `isMoveSource()`, `isMoveDestination()` helper functions
  - `findMovePair()` function to find the matching revision for a move
  - Comprehensive JSDoc documentation on the `Revision` interface
  - All types are properly exported from the package
- **Paginated Headers and Footers** - Headers/footers now render correctly with pagination enabled
  - When both `RenderHeadersAndFooters` and `RenderPagination=Paginated` are enabled, headers and footers appear on each page
  - Per-section header/footer support with section index tracking
  - First page headers/footers supported (when `w:titlePg` is set in document)
  - Even page headers/footers supported for different odd/even page layouts
  - Headers/footers rendered into hidden registry for client-side cloning per-page
  - New data attributes: `data-header-height`, `data-footer-height` on section elements
  - TypeScript `PageDimensions` interface extended with `headerHeight` and `footerHeight`
  - CSS classes `.page-header` and `.page-footer` for positioning within page boxes
  - Automatic hiding of system page number when document has footer content
  - See `docs/architecture/paginated_headers_footers.md` for full architecture details
- **Per-page Footnote Rendering** - Footnotes now appear at the bottom of each page where they are referenced
  - When `RenderFootnotesAndEndnotes=true` with `RenderPagination=Paginated`, footnotes are distributed per-page
  - Footnote registry stores footnotes in a hidden container for client-side distribution
  - `data-footnote-id` attributes added to footnote references for tracking
  - Single-pass, forward-only pagination algorithm (lazy-loading compatible)
  - Pagination engine measures footnote space and includes it in page layout calculations
  - Footnotes render with separator line (`<hr>`) above them
  - **Footnote continuation**: Long footnotes that don't fit on a page are split at paragraph boundaries and continue on subsequent pages (matching Word/Office behavior)
  - **Dynamic footnote area expansion**: Footnote area can expand upward into body content space (up to 60% of page height) to fit more footnote content before splitting, reducing wasted space
  - Endnotes remain at document end (not per-page) - traditional behavior preserved
  - New TypeScript methods: `parseFootnoteRegistry()`, `extractFootnoteRefs()`, `measureFootnotesHeight()`, `addPageFootnotes()`, `splitFootnoteToFit()`, `measureContinuationHeight()`
  - New TypeScript interfaces: `FootnoteContinuation`, `PartialFootnote`
  - New TypeScript constants: `MAX_FOOTNOTE_AREA_RATIO` (0.6), `MIN_BODY_CONTENT_HEIGHT` (72pt)
  - New CSS classes: `.page-footnotes`, `.footnote-item`, `.footnote-number`, `.footnote-content`, `.footnote-continuation`
- `SkiaSharpHelpers.cs` - Color utilities for SkiaSharp compatibility
- `GetPackage()` extension method in `PtOpenXmlUtil.cs` for SDK 3.x Package access
- `SkiaSharp.NativeAssets.Linux.NoDependencies` package for Linux runtime support

### Fixed
- **React hooks loading state not rendering before WASM blocks** (Issue #45) - Fixed `isConverting`/`isComparing`/`isLoading` states in React hooks not painting before WASM execution blocks the main thread. Added `requestAnimationFrame` yielding after state updates in:
  - `useConversion`: `convert()` function
  - `useComparison`: `compare()` and `compareToHtml()` functions
  - `useAnnotations`: `reload()`, `add()`, and `remove()` functions
  - `useDocumentStructure`: `reload()` function

- **Header/footer positioning in paginated mode** - Fixed headers and footers overlapping with body content. Headers now properly constrain to the top margin area (`height: marginTop`) and footers constrain to the bottom margin area (`height: marginBottom`). Uses flexbox layout for proper content alignment within constrained areas.

- **DocumentBuilder relationship copying** - Fixed bug where relationship IDs from source documents could incorrectly match existing IDs in target header/footer parts when using InsertId functionality. This caused validation errors like "The relationship 'rIdX' referenced by attribute 'r:embed' does not exist."
  - Removed flawed early-return optimization in `CopyRelatedImage()` that skipped processing when target part had matching relationship ID
  - Fixed diagram relationship handling (`R.dm`, `R.lo`, `R.qs`, `R.cs` attributes) to properly copy parts from source documents
  - Fixed chart and user shape relationship handling
  - Fixed OLE object relationship handling
  - Fixed external relationship attribute update to use correct attribute name parameter

- **SpreadsheetWriter date handling** - Fixed date cells being written with invalid ISO 8601 string format. Dates are now properly converted to Excel serial date numbers (days since December 30, 1899) which is required for transitional OOXML format.

- **WmlComparer null Unid handling** - Fixed null reference exceptions when comparing documents with elements lacking Unid attributes.

- **WmlComparer footnote/endnote comparison** (6 tests: WC-1660, WC-1670, WC-1710, WC-1720, WC-1750, WC-1760) - Fixed `AssignUnidToAllElements` to assign Unid to footnote/endnote elements themselves, enabling proper reconstruction of multi-paragraph footnotes/endnotes by `CoalesceRecurse`.

- **WmlComparer table row comparison** (1 test: WC-1500) - Added LCS-based row matching (`ApplyLcsToTableRows`) for large tables (7+ rows) when content differs significantly, preventing cascading false differences from insertions/deletions in the middle of tables.

- **WASM CDN loading CORS issue** - Fixed cross-origin loading failures when WASM files are served from CDNs (jsDelivr, unpkg). The .NET WASM runtime uses `credentials:"same-origin"` for fetch requests, which conflicts with CDN's `Access-Control-Allow-Origin: *` wildcard header. Build script now patches `dotnet.js` to use `credentials:"omit"` for CDN compatibility.

- **Vite bundler compatibility** - Added `@vite-ignore` comment to dynamic import in `npm/src/index.ts` to prevent Vite from trying to analyze/resolve the WASM loader path during development builds.

- **Pagination content overflow** - Fixed content overflowing page boundaries in the paginated view. The issue was caused by applying CSS transform scale to the content area while using inconsistent coordinate systems for positioning. The fix applies the scale transform to the entire page box instead, ensuring proper clipping and consistent scaling of all page elements.

- **WmlComparer legal numbering preservation** ([Issue #1634](https://github.com/dotnet/Open-XML-SDK/issues/1634)) - Fixed comparison losing legal numbering (`w:isLgl`) when comparing documents with different numbering styles. The comparer now properly merges numbering definitions from the revised document into the result:
  - Copies `abstractNum` and `num` elements from revised document when missing in original
  - Reuses existing definitions when content matches (regardless of ID)
  - Remaps IDs when conflicts occur to avoid duplicates

- **WmlToHtmlConverter null rPr crash** - Fixed `InvalidOperationException` crash in `DefineRunStyle` and `GetLangAttribute` when converting runs without `w:rPr` elements. Changed `.First()` to `.FirstOrDefault()` with null checks to handle runs that have no explicit run properties gracefully.

### Changed
- Replaced `FontPartType`/`ImagePartType` with `PartTypeInfo` pattern for SDK 3.x compatibility
- Replaced `.Close()` calls with `Dispose()` pattern
- Migrated all color handling from `System.Drawing.Color` to `SKColor`
- Migrated font handling from `FontFamily`/`FontStyle` to `SKFontManager`/`SKTypeface`
- Migrated image handling from `Bitmap`/`ImageFormat` to `SKBitmap`/`SKEncodedImageFormat`

### Documentation
- Updated `docs/architecture/wml_to_html_converter_gaps.md` with comprehensive gap analysis including pagination mode limitations, DrawingML text handling, and prioritized fix recommendations

### Test Status
- 1051 passed, 0 failed, 1 skipped out of 1052 tests (~99.9% pass rate)
- Header/footer and footnote pagination changes tested via manual integration testing
