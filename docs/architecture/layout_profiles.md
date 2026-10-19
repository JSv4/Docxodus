# Layout compatibility profiles

Status: proposed design; the selector and policy module below are not implemented.
Implementation follow-up: [#1245](https://github.com/JSv4/Docxodus/issues/1245).
Audited against upstream `main` at
[`31655e6c`](https://github.com/JSv4/Docxodus/tree/31655e6caa0b1d2f4a5a2a6e21d8d5a65f88df98)
on 2026-10-18. Historical issue titles are hypotheses; the evidence ledger below includes
their later measurements and resolutions.

## Decision

Own the profile selection and defaults in the C# HTML rendering core. Carry that one selection
through the rendering facades and generated HTML to the TypeScript paginator. There are two
places that implement behavior: C# for HTML/CSS decisions, and TypeScript for browser layout and
pagination. Both consume the same selected policy. Start with the two measured footnote rules;
keep the existing converter and paginator.

```csharp
var settings = new WmlToHtmlConverterSettings
{
    LayoutProfile = LayoutProfile.LibreOffice25_8V1,
    RenderPagination = PaginationMode.Paginated,
};
var html = WmlToHtmlConverter.ConvertToHtml(document, settings);
```

```ts
type LayoutProfile = "word-generic-v1" | "libreoffice-25.8-v1";

const html = await convertDocxToHtml(docx, {
  paginationMode: PaginationMode.Paginated,
  layoutProfile: "libreoffice-25.8-v1",
});
paginateHtml(html, container); // consumes the policy carried by this HTML

await convertDocxToPdf(docx, {
  reviewProfile: "final",
  commentProfile: "hidden",
  layoutProfile: "word-generic-v1",
});
```

Omission resolves to `word-generic-v1`. Its initial policy preserves the audited main's
behavior, including its known gaps. `word-generic` means the measured Word behavior documented
here, taking the document's compatibility settings into account; it does not claim a particular
Word build or complete Word equivalence. The LibreOffice profile initially covers the footnote
differences measured on **25.8.7.3**. Its name is not evidence that every 25.8 patch behaves alike.
Coverage outside those rules remains the shared Docxodus implementation.

The profile selects rendering behavior. It does not modify the DOCX, choose accept/reject state,
change comment presentation, select fonts, or choose the reference used to score an artifact.
`reviewProfile`, `commentProfile`, `markupPresentation` and the font contract remain independent.
In particular, `markupPresentation: "word"` is not a layout-profile switch.

## Smallest implementation

### Core setting, two implementation sites

Add a public `LayoutProfile` enum and a small core resolver (for example
`Docxodus/LayoutProfiles.cs`). It owns the canonical wire IDs, default and resolved policy values.
Add the setting to `WmlToHtmlConverterSettings` and the facade's `HtmlConversionOptions`; include
the legacy `HtmlConverterSettings` compatibility constructor. Resolve once per render, rejecting
unknown IDs. Every transport forwards an omitted value or the selected ID; it does not invent a
separate default.

Feature owners may use a switch or a named policy value to select the measured behavior. Put a
decision in the helper that owns the feature, and share it when several call sites need the same
rule. A table-spacing difference belongs in the converter's table code; note-page allocation
belongs in the paginator. There is no requirement to move either algorithm to the other language,
and no separate renderer or branch per transport.

For browser-owned rules, the core emits a small versioned policy descriptor with the HTML. For
example, use one escaped `data-docx-layout-policy` JSON attribute on `#pagination-staging`, alongside
the existing `data-footnote-layout-like-word8` metadata. The descriptor contains only the selected
profile and the values needed by browser layout. Its initial shape is:

```ts
interface PaginationLayoutPolicy {
  readonly schemaVersion: 1;
  readonly id: LayoutProfile;
  readonly maximumNoteBandFraction: number;
  readonly continuingNoteOwnsPage: boolean;
}

// word-generic-v1:      fraction 1.00, ownsPage true
// libreoffice-25.8-v1:  fraction 0.95, ownsPage false
```

`npm/src/layout-profiles.ts` defines the public ID type and validates/reads this descriptor; it
does not maintain another profile-to-policy table. `PaginationEngine` consumes the resolved values.
Use the existing XML/JSON serializers to escape metadata, validate its schema and value types,
and reject unknown versions. This needs no code generator, plugin registry or general policy
language. Callers choose a profile, not custom fractions or independent feature toggles.

The initial behavioral branches are still in TypeScript, but the core setting is immediately used
to select and emit their policy. Later converter-owned rules can use the same setting without a
second public-API introduction. Emit profile identity on full flow HTML as well, and retain it on
the finalized paginated artifact after staging is removed. Incremental fragment roots carry the
profile ID so the editor can check that replacements match their host document.

New output must always carry the resolved metadata. Keep existing unmarked HTML working through
an explicit legacy-default viewer path, tested separately from profiled output. It cannot satisfy
a request for a non-default profile or a profile-bound export report; reconvert with a supporting
core in those cases. Malformed metadata is an error, not a reason to use the legacy fallback.
Separate converter and paginator instances can render different profiles without global state.

### Two rules, applied consistently

| Rule | `word-generic-v1` | `libreoffice-25.8-v1` |
|---|---|---|
| Maximum note band, separator included | Full body band minus body already placed | At most 95% of the body band, also limited by body already placed |
| A carried note continues beyond this page | Finish the page before placing body content | Permit body content in the remaining band, subject to the existing fit/keep rules |

For body-band height `H` and already committed body space `U`, the upper bound is
`max(0, min(H * fraction, H - U))`. Use the paginator's existing effective page budgets and margin
accounting to derive those values. This is a cap, not a fixed-size note box. The LibreOffice
measurements include the 95% cap on note-only pages; do not expand those to 100% just because the
body has ended. Neither profile restores the old 60% cap.

Centralize that calculation in the existing note-budget helper and use it for first-note admission,
fresh-page probes, continuation partitioning and final band painting. Audit every use of
`MAX_FOOTNOTE_AREA_RATIO`; changing the constant at one call site is insufficient. The
`continuationRunsOn` early page finish becomes conditional on `continuingNoteOwnsPage`.

Keep these invariants shared: no note covers body already placed; a continued note stays last;
notes retain citation order; text and row fragments are neither lost nor duplicated; keep rules
and the liveness fallback still apply. Preserve `footnoteLayoutLikeWord8` from the source
compatibility metadata. A profile is not an instruction to erase explicit formatting or
compatibility flags. Any future rule interpreting them differently needs its own evidence.

There is one important interaction to recover, not a reason to fork the paginator:
[PR #1215](https://github.com/JSv4/Docxodus/pull/1215) removed the deferred-note rejoin path from
[PR #1209](https://github.com/JSv4/Docxodus/pull/1209), because it was unreachable under the full-band
Word rule. A capped profile can make it reachable again. If the last body page lets a carried note
finish after repartitioning, release deferred notes onto that same page in order. Reuse a common
helper and cover the last-body-page transition explicitly; toggling the two values alone does not
establish LibreOffice compatibility.

### Public surfaces and switching

Expose the same selection on native conversion, browser conversion/viewer/editor, Python and MCP
rendering, standalone HTML/PDF export and delivery rendering. This follows the existing rendering
option plumbing; it does not add a setting to unrelated document mutations or comparison options.
Core ownership alone does not make the handwritten mappings forward a new property automatically:

| Surface | Required forwarding |
|---|---|
| Core and facade | `WmlToHtmlConverterSettings`, legacy settings copy, `HtmlConversionOptions`, full/chrome conversion and `BuildBlockConverterSettings` |
| WASM and npm | `DocumentConverter` and session render bridges; `ConversionOptions`, `needsCompleteConversion`, worker calls, and matching export declarations |
| Incremental editor | `editorRenderProfile()`, its core parser, single/batch/range block renders and windowed mounts; never silently fall back to an older bridge that drops the selected profile |
| Python and MCP | Rendering option types/serialization, host parsing and rendering tool schemas/dispatch; preserve the same selection in full and block previews |
| CLI | Both native `docx2html` and Node export accept `--layout-profile` and use the core-defined IDs |
| Export and delivery | `PaginatedHtmlOptions`/Node options, `conversionOptions()`, .NET host-renderer configuration, framed requests, reports and batch identities |

For a low-level `paginateHtml(html, ...)` call, HTML is already converted: consume its descriptor.
An optional `expectedLayoutProfile` can assert a match, but must not override it. An input setting
that disagrees with the HTML requires reconversion, even while the only differing rules happen to
be in TypeScript. This prevents a future hybrid of Word-converted tables and LibreOffice pagination.
For a viewer given DOCX bytes, changing `layoutProfile` reconverts and then paginates.

React must include the selector in conversion/render dependencies. The editor retains it across
edits, remounts and pagination toggles, and forwards it on every incremental render. Switching
reconverts the current source/session snapshot and repaginates; it does not mutate existing page
clones or revive stale staging. Fragment profile mismatches trigger a fresh matching render.
Snapshot the option before asynchronous work. An editor UI control is optional follow-up; the
API switch is sufficient for this scope.

## Identity and compatibility

The selected policy must be observable and must invalidate layout-dependent results even when two
profiles happen to produce identical pixels for a particular document.

- Include the **resolved** ID in the canonical layout-options digest for every render. Omitted
  and explicit Word defaults must produce identical identities. A profile change must change
  the layout digest and renderer fingerprint. Existing runtime/source/font digests retain their
  own meanings; exact reproduction still requires the renderer build and font environment.
  Bind the canonical core-resolved descriptor as well as its ID, and verify it matches the
  requested profile before layout/report publication. HTML metadata is provenance, not proof
  that arbitrary caller-supplied HTML was actually produced by the claimed renderer.
  Version the changed layout material as `docxodus:layout-options:v2`; keep the old material
  definition for validating historical reports.
- Record the resolved ID as `renderReport.options.layoutProfile` in both complete and failed
  reports once input validation succeeds. Do not invent a valid profile for rejected input.
- Keep **PageMap v1 unchanged**. Its existing `rendererFingerprint` binds the choice, so pages
  and fragments need no new fields. Cache and batch-group keys must include the canonical
  layout digest. A profile switch invalidates prior page citations and cached maps.
- Direct viewer callers that supply `layoutToken` must bind it to the profile as well. Document
  and test that obligation; an arbitrary caller-supplied fingerprint is not independently
  verifiable by the paginator. Wrappers must invalidate an old map while rerendering, and must
  not reuse its token when the selector changes.

The current report is v2 with a closed options schema. Introduce **render report v3** for the new
required field, retaining the frozen v1/v2 schemas. Use one current writer rather than choosing
report versions according to the profile. This changes report identity even for the default,
but does not change the default page layout. Release notes must distinguish those two facts.

Migrate the browser/Node writers and validators, schema and schema tests, the .NET delivery reader,
delivery identity checks, and benchmark consumers together. The .NET reader must require the
requested resolved profile, not accept a report claiming another policy for its request. Preserve
old-report validation through the old schema; do not silently interpret v2 as an explicitly
selected v1 profile. Check the strict allowlists in `npm-export/src/index.ts` and `host.ts` rather
than assuming a TypeScript property propagates over a wire. Add the selector to a versioned host
request (v2); keep the frozen v1 definition and give unsupported peers a clear compatibility error.
Include the choice in the .NET host renderer's `DescribeBatch` identity as well as the browser's
layout digest. A default-only host is not sufficient for the initial cross-surface implementation.

Profile versions describe intended behavior; renderer fingerprints identify implementations.
A bug fix that makes an implementation conform to its recorded rule can keep the ID. A deliberate
change to the rule or an incompatible reference target requires a new profile version. Extending
coverage requires updating the evidence and documenting whether existing policy semantics change.
Do not silently retarget `libreoffice-25.8-v1` when CI upgrades LibreOffice.

## Historical evidence ledger

This is a focused rendering audit, not an inventory of every issue mentioning an office suite.
The detailed measurements remain in [OOXML corner cases](../ooxml_corner_cases.md), the linked
issues/PRs and the existing Word evidence records. Issue status below was checked on 2026-10-18.
Use this ledger to decide which rules merit profiles, not to resurrect every historical behavior.

| Evidence | Finding after investigation | Profile disposition |
|---|---|---|
| [#1119](https://github.com/JSv4/Docxodus/issues/1119), [#1178](https://github.com/JSv4/Docxodus/issues/1178), [#1215](https://github.com/JSv4/Docxodus/pull/1215), closed/merged | LibreOffice 25.8.7.3 reserves 5% of the body band. Later Word captures show a note can take all space below its citation, and a running continuation can own a page. | **Initial policies.** The older #1119 corner-case section describes an intermediate implementation and says Word was unmeasured; the later #1178 measurements supersede that claim. |
| [#1179](https://github.com/JSv4/Docxodus/issues/1179), [#1209](https://github.com/JSv4/Docxodus/pull/1209), closed/merged | Both suites keep a continued note last and preserve citation order. Their body placement differs while a continuation runs on (cases B/C). | **Shared ordering; profiled body placement.** Never disable ordering to imitate a page count. |
| [#1179 corner-case table, case E](https://github.com/JSv4/Docxodus/blob/31655e6caa0b1d2f4a5a2a6e21d8d5a65f88df98/docs/ooxml_corner_cases.md#footnotes-a-continued-note-stays-last-and-notes-keep-their-order-issue-1179); benchmark `notes-keep-lines-note-in-table` | Word moves the citing table and following body to the kept note's page; LibreOffice leaves the table on the preceding page. Audited Docxodus follows LibreOffice. | **Confirmed difference, deferred implementation.** Add a narrowly conditioned table-citation policy when fixing Word placement. The initial Word profile still has this known gap. The 2026-10-18 audit added this exact case, code path and acceptance checks to [#1217](https://github.com/JSv4/Docxodus/issues/1217); closed #1179 alone did not track the remaining gap. |
| [#1219](https://github.com/JSv4/Docxodus/issues/1219), open | Measured table text x is about 73.0pt in Word, 77.9pt in LibreOffice and 78.1pt in Docxodus. Compatibility-mode/table-indent behavior is the suspected mechanism. | **Confirmed positional difference; rule pending.** Reduce by compatibility mode, explicit/implicit `tblInd` and cell margins before adding a converter policy. Do not apply a universal −5pt correction. |
| [#950](https://github.com/JSv4/Docxodus/issues/950), resolved by [#957](https://github.com/JSv4/Docxodus/pull/957) | SSIM fell after a Word baseline fix, but measured body lines moved closer to both suites. The endnote block itself started at row 688 in Word, 676 in LibreOffice and 697 in Docxodus. | **Body-line regression hypothesis rejected; endnote placement candidate pending reduction.** Separate block placement from fonts and rasterization before adding a rule. |
| [#398](https://github.com/JSv4/Docxodus/issues/398), closed | The supposed 28px list top-margin deviation disappeared after fixing accumulated line spacing. Word and LibreOffice both start ink at row 117; Docxodus at 118. | **No profile rule.** Shared line-spacing correction; the remaining pixel is rasterization. |
| [#426](https://github.com/JSv4/Docxodus/issues/426), [#427](https://github.com/JSv4/Docxodus/issues/427), [#428](https://github.com/JSv4/Docxodus/issues/428), closed | Word and LibreOffice agree on nested-table paragraph-after suppression, the captured TOC's black field-result text, and the captured page-top heading spacing. | **Shared fixes.** Retain the fixtures' conditions; these are not competing profiles. |
| [#882](https://github.com/JSv4/Docxodus/issues/882), closed | Captures disprove the original assumption about clipping exact-line glyphs. Both suites place the baseline approximately at 80% of exact line height. | **Shared fix.** Small quantized origin differences do not justify a new baseline algorithm. |
| [#850](https://github.com/JSv4/Docxodus/issues/850), [#908](https://github.com/JSv4/Docxodus/issues/908), [#940](https://github.com/JSv4/Docxodus/issues/940), closed | Whole-pixel line pitch, CSS leading placement and mixed paragraph strut metrics were Docxodus/Chromium issues. Later recorded Arial measurements corrected the initial unsupported pitch assumption. | **Keep shared precision fixes.** These issues alone do not establish an alternative LibreOffice baseline policy. |
| [#955](https://github.com/JSv4/Docxodus/issues/955), closed | Word's footnote text baseline was 717.48pt, LibreOffice's 717.35pt; Docxodus was 716.25pt. | **Shared placement fix.** The meaningful error was Docxodus's, not an office-suite split. |
| [#379](https://github.com/JSv4/Docxodus/issues/379), [#403](https://github.com/JSv4/Docxodus/issues/403), closed | Font substitution affected wraps; LibreOffice 24.2 and 25.8 used different default footnote separator widths (25% of column versus two inches). | **Environment/version evidence.** Keep fonts independent and pin reference builds. Do not add a legacy separator rule to the 25.8 profile. |
| [#108](https://github.com/JSv4/Docxodus/issues/108), closed; [#1244](https://github.com/JSv4/Docxodus/issues/1244), open; current zero-height table fixtures | The converter's 7.5pt top-margin heuristic shifts even the first visible row. The issue asserted Word/LibreOffice spacing without a measured two-renderer rule. | **Bug candidate pending Word adjudication.** Do not make the arbitrary gap a LibreOffice/Word preference. |
| [#1182](https://github.com/JSv4/Docxodus/issues/1182), [#1217](https://github.com/JSv4/Docxodus/issues/1217), [#1124](https://github.com/JSv4/Docxodus/issues/1124), open | Residual charter pagination, note/export gaps, and justified-word positions remain under investigation. | **Investigate individually.** An aggregate page-count or raster-score difference cannot define a profile rule. |

Some historical differences concern document transformations rather than rendering. The corner-case
entries on Word-compare table-margin normalization and LibreOffice's positional reassociation of
orphaned footnotes are examples. A layout profile must not change comparison output, reference
identity, revision acceptance or package repair to reproduce those behaviors. The redline bookmark
failure [#1218](https://github.com/JSv4/Docxodus/issues/1218) and benchmark adapter defect
[#54](https://github.com/JSv4/docx-fidelity-bench/issues/54) are likewise independent fixes.

## Evidence and maintenance rule

For a rendering change, decide whether it is a shared defect, a measured renderer difference,
an environment effect, or an unresolved hypothesis. Read the issue's resolution and linked
measurements as well as its opening claim. If only one renderer has been measured, say so and
leave the profile classification pending; a lower LibreOffice score is not Word evidence.

For every new or changed profile rule, record in the corner cases and link from this ledger:

- A minimal authored fixture and control, input digest, exact OOXML/compatibility conditions,
  and the issue or PR that established them.
- Word capture path/build/date (or explicitly unknown), LibreOffice version, fonts/substitutions,
  and comparable page/line/marker measurements with tolerances. Follow the existing
  [Word-reference procedure](../../npm/tests/visual-parity/WORD_REFERENCE.md); keep its numbers-only
  evidence convention. Renderer output and a reporter's inference are different evidence levels.
- The affected policy, default behavior, implementation owner and tests. Update both profiles
  deliberately, including when their behavior should stay identical. Retire disproved hypotheses
  explicitly instead of silently deleting their history.

Prefer adding a row to the existing evidence records over creating a separate fixture database.
Keep detailed measurements out of `CLAUDE.md`; it only directs agents here.

## Verification and benchmark integration

Implement in reviewable increments: first add the core setting, facade forwarding and HTML
descriptor with unchanged default geometry; then make the paginator consume it and verify the
LibreOffice footnote policy; complete client/CLI/editor propagation and export identity migration
before claiming cross-surface support. Table citation and table indentation policies remain
separate follow-ups; their implementation can use the setting already present in both layers.

The implementation acceptance checks are:

1. **Default equivalence:** omitted and explicit Word selectors agree. Extracting the policy
   preserves audited-main geometry, line/page assignments, text and existing known gaps.
2. **Measured differences:** run first-line and mid-page tall notes, the four
   `notes-explicit-breaks-*` cases, and the #1178/#1179 continued-note fixtures through both profiles.
   Assert independent line/page measurements for each target, not values computed by the policy.
3. **Shared invariants:** cover later-note order, keep-with-next chains, note tables, indivisible
   oversized parts, the last-body-page rejoin, note-only tails and source compatibility flags.
   Exercise fragmentation on/off where the branch depends on it. No lost/duplicated text, blank
   pages, overlap, or nontermination is an acceptable compatibility behavior.
4. **Plumbing:** native, WASM, Python/MCP, CLI and export requests produce the same selected
   metadata. Full and incremental renders agree; a selector change reconverts React/editor from
   current source; simultaneous engines are independent. Test `needsCompleteConversion`, legacy
   bridge behavior, unmarked HTML and conversion/pagination profile mismatches. Unknown IDs or
   descriptor versions fail consistently. Plain controls retain their geometry.
5. **Identity:** a profile switch changes the digest/fingerprint even on an unaffected document;
   stale citations/maps are rejected by the consuming layer. Verify report v3, frozen-schema
   compatibility, .NET requested-profile binding, and HTML/PDF/PageMap agreement from the same tree.

Add a cross-language contract fixture that serializes the core's actual descriptors and feeds
them to the TypeScript consumer. It must catch missing fields, changed wire IDs and a browser
reinterpreting the core's values; a test that constructs its own parallel policy table is insufficient.

Reuse `pagination-note-area`, `pagination-continued-note-last`, `pagination-note-table` and the
export specs for these assertions. Parameterize the affected cases and a small set of shared
controls; there is no need to duplicate the entire suite per profile. Relevant defaults and
wire contracts still require their normal regression checks.

The benchmark owns reference selection, assertions and score presentation; it passes the public
selector and records it. It must not inject private CSS, patch the paginator, or choose a profile
by fixture name. Record the candidate SHA, resolved profile, reference provider/build, suite,
assertion-contract version and fonts with every score. Compare like profile/reference pairs over
time; retain mixed comparisons with an explicit label rather than silently changing their meaning.

The current suite contains **Word-adjudicated mandatory assertions inside a LibreOffice reference
track**, notably `notes-carried-note-defers-citations` and `notes-keep-lines-note-in-table`. Selecting
the LibreOffice profile does not make that track a pure LibreOffice correctness score. Keep the
existing contract immutable; add an explicitly versioned target-aware assertion contract for new
profile scores, with unconditional content/validity checks and documented target-specific layout
checks. Do not weaken a Word assertion or omit a failure just to make the LibreOffice run pass.
Word reference font eligibility is also independent of the selector; a Word score needs eligible
captures, not merely a profile named Word.

The 2026-10-18 benchmark run is the baseline investigation, not a measurement of either proposed
profile: [run 39140904566](https://github.com/JSv4/docx-fidelity-bench/actions/runs/39140904566).
The adapter repair in benchmark #54 is needed for complete HTML coverage regardless of profile.
