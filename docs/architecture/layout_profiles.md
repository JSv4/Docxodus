# Layout compatibility profiles

Status: proposed design; the selector and policy module below are not implemented.
Audited against upstream `main` at
[`31655e6c`](https://github.com/JSv4/Docxodus/tree/31655e6caa0b1d2f4a5a2a6e21d8d5a65f88df98)
on 2026-10-18. Historical issue titles are hypotheses; the evidence ledger below includes
their later measurements and resolutions.

## Decision

Own the profiles in Docxodus. Add one optional `layoutProfile` to the existing paginated
rendering APIs, backed by one small policy resolver and the existing paginator. Start with
the two measured footnote rules. Keep conversion, font resolution, line measurement, note
splitting and artifact production shared.

```ts
type LayoutProfile = "word-generic-v1" | "libreoffice-25.8-v1";

paginateHtml(html, container, { layoutProfile: "libreoffice-25.8-v1" });

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

### One owner, no second renderer

Create `npm/src/layout-profiles.ts` with the public ID type, a single default, runtime validation
and an internal immutable policy table. Resolve once when a render begins; pass the resolved
policy into the existing `PaginationEngine`. Reject unknown IDs rather than falling back.
The following is the initial internal shape, not another public options API:

```ts
interface LayoutPolicy {
  readonly id: LayoutProfile;
  readonly maximumNoteBandFraction: number;
  readonly continuingNoteOwnsPage: boolean;
}

// word-generic-v1:      fraction 1.00, ownsPage true
// libreoffice-25.8-v1:  fraction 0.95, ownsPage false
```

Only the resolver compares profile IDs. Layout code reads named policy values; callers cannot
provide custom fractions or override individual rules. There is no plugin registry, new render
backend, subclass tree or global setting. Separate engine instances can render different profiles.

Both initial rules belong to `npm/src/pagination.ts`, so **do not add a converter setting or
thread an unused option through `WmlToHtmlConverter`, WASM, session operations, Python or MCP**.
The existing source HTML can feed either profile. When a measured converter-owned difference is
implemented later, introduce its setting through `HtmlConversionOps` and the normal ripple
checklist. At that point, use one shared data contract or generated projection for the policies
needed by both runtimes; do not copy a policy table into C# and TypeScript now.

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

Add the same optional selector to `PaginationOptions` and `PaginatedHtmlOptions`; the Node export
types reuse that contract. Forward it through `paginateHtml`, React's viewer/hook, the paginated
editor's existing options, and standalone HTML/PDF/batch export. Add `--layout-profile` to the
export CLI. Re-export the type and default through the existing package entry points without
pulling editor or Node dependencies into the browser export entry.

React must include the selector in its dependency arrays. The editor must retain it across
edits, remounts and pagination toggles. Switching starts a full render from pristine staging
content, or the editor's current session snapshot; it must not repaginate already fragmented
page clones or revive stale pre-edit staging. Reuse existing render/remount entry points rather
than adding an in-place policy mutation API. Snapshot the option before asynchronous export work.

The first release exposes selection on these JavaScript rendering surfaces. Native conversion
remains profile-neutral; .NET delivery continues to request the default. Its render-report reader
still needs the schema migration below. A later .NET delivery selector belongs on the host
renderer configuration, with an explicitly versioned host request, rather than on every editing
operation. An editor UI control is optional follow-up; the API switch is sufficient for this scope.

## Identity and compatibility

The selected policy must be observable and must invalidate layout-dependent results even when two
profiles happen to produce identical pixels for a particular document.

- Include the **resolved** ID in the canonical layout-options digest for every render. Omitted
  and explicit Word defaults must produce identical identities. A profile change must change
  the layout digest and renderer fingerprint. Existing runtime/source/font digests retain their
  own meanings; exact reproduction still requires the renderer build and font environment.
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
delivery identity checks, and benchmark consumers together. The .NET default-only reader must
require `word-generic-v1`, not accept a report claiming another policy for its request. Preserve
old-report validation through the old schema; do not silently interpret v2 as an explicitly
selected v1 profile. Check the strict allowlists in `npm-export/src/index.ts` and `host.ts` rather
than assuming a TypeScript property propagates over a wire. The initial framed host remains
default-only and does not accept the new selector under its existing request version.

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
| [#1179 corner-case table, case E](https://github.com/JSv4/Docxodus/blob/31655e6caa0b1d2f4a5a2a6e21d8d5a65f88df98/docs/ooxml_corner_cases.md#footnotes-a-continued-note-stays-last-and-notes-keep-their-order-issue-1179); benchmark `notes-keep-lines-note-in-table` | Word moves the citing table and following body to the kept note's page; LibreOffice leaves the table on the preceding page. Audited Docxodus follows LibreOffice. | **Confirmed difference, deferred implementation.** Add a narrowly conditioned table-citation policy when fixing Word placement. The initial Word profile still has this known gap. Closed #1179 and open [#1217](https://github.com/JSv4/Docxodus/issues/1217) do not constitute an exact open issue for it. |
| [#1219](https://github.com/JSv4/Docxodus/issues/1219), open | Measured table text x is about 73.0pt in Word, 77.9pt in LibreOffice and 78.1pt in Docxodus. Compatibility-mode/table-indent behavior is the suspected mechanism. | **Confirmed positional difference; rule pending.** Reduce by compatibility mode, explicit/implicit `tblInd` and cell margins before adding a converter policy. Do not apply a universal −5pt correction. |
| [#950](https://github.com/JSv4/Docxodus/issues/950), resolved by [#957](https://github.com/JSv4/Docxodus/pull/957) | SSIM fell after a Word baseline fix, but measured body lines moved closer to both suites. The endnote block itself started at row 688 in Word, 676 in LibreOffice and 697 in Docxodus. | **Body-line regression hypothesis rejected; endnote placement candidate pending reduction.** Separate block placement from fonts and rasterization before adding a rule. |
| [#398](https://github.com/JSv4/Docxodus/issues/398), closed | The supposed 28px list top-margin deviation disappeared after fixing accumulated line spacing. Word and LibreOffice both start ink at row 117; Docxodus at 118. | **No profile rule.** Shared line-spacing correction; the remaining pixel is rasterization. |
| [#426](https://github.com/JSv4/Docxodus/issues/426), [#427](https://github.com/JSv4/Docxodus/issues/427), [#428](https://github.com/JSv4/Docxodus/issues/428), closed | Word and LibreOffice agree on nested-table paragraph-after suppression, the captured TOC's black field-result text, and the captured page-top heading spacing. | **Shared fixes.** Retain the fixtures' conditions; these are not competing profiles. |
| [#882](https://github.com/JSv4/Docxodus/issues/882), closed | Captures disprove the original assumption about clipping exact-line glyphs. Both suites place the baseline approximately at 80% of exact line height. | **Shared fix.** Small quantized origin differences do not justify a new baseline algorithm. |
| [#850](https://github.com/JSv4/Docxodus/issues/850), [#908](https://github.com/JSv4/Docxodus/issues/908), [#940](https://github.com/JSv4/Docxodus/issues/940), closed | Whole-pixel line pitch, CSS leading placement and mixed paragraph strut metrics were Docxodus/Chromium issues. Later recorded Arial measurements corrected the initial unsupported pitch assumption. | **Keep shared precision fixes.** These issues alone do not establish an alternative LibreOffice baseline policy. |
| [#955](https://github.com/JSv4/Docxodus/issues/955), closed | Word's footnote text baseline was 717.48pt, LibreOffice's 717.35pt; Docxodus was 716.25pt. | **Shared placement fix.** The meaningful error was Docxodus's, not an office-suite split. |
| [#379](https://github.com/JSv4/Docxodus/issues/379), [#403](https://github.com/JSv4/Docxodus/issues/403), closed | Font substitution affected wraps; LibreOffice 24.2 and 25.8 used different default footnote separator widths (25% of column versus two inches). | **Environment/version evidence.** Keep fonts independent and pin reference builds. Do not add a legacy separator rule to the 25.8 profile. |
| [#108](https://github.com/JSv4/Docxodus/issues/108), closed; current zero-height table fixtures | The converter's 7.5pt top-margin heuristic shifts even the first visible row. The issue asserted Word/LibreOffice spacing without a measured two-renderer rule. | **Bug candidate pending Word adjudication.** Do not make the arbitrary gap a LibreOffice/Word preference. |
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

Implement in reviewable increments: first extract the default policy with unchanged layout;
then add and verify the LibreOffice footnote policy; then expose it through the rendering surfaces
and migrate export identity. Do not advertise export support until its report/schema consumers
are updated. Table citation and table indentation policies are separate follow-ups.

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
4. **Plumbing:** a selector change rerenders React/editor from current source; simultaneous
   engines do not affect each other; unknown IDs fail consistently; Node/CLI selection reaches
   the actual paginator. Plain documents without affected features keep the same geometry.
5. **Identity:** a profile switch changes the digest/fingerprint even on an unaffected document;
   stale citations/maps are rejected by the consuming layer. Verify report v3, frozen-schema
   compatibility, .NET default binding, and HTML/PDF/PageMap agreement from the same finalized tree.

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
