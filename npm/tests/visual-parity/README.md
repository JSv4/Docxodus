# PDF visual-fidelity benchmark

This opt-in benchmark renders a small, stratified set of repository-tracked DOCX fixtures through
Docxodus and LibreOffice. The release-gating path tests the artifact users receive: it invokes the
supported `@docxodus/export` API, writes a Docxodus PDF, writes the LibreOffice reference PDF, and
rasterizes both PDFs through the same Poppler command at 96 DPI before computing page metrics.
It also checks physical PDF geometry, selectable text, and supported link annotations independently
of the raster comparison.

The earlier browser-page-to-LibreOffice benchmark remains available as a separate diagnostic and
keeps its existing ratchet history. The generated-PDF run has a distinct pipeline identity and
artifact root so a PDF regression cannot be confused with movement in browser screenshot capture.
Neither run is part of the normal browser suite because LibreOffice and Poppler are host tools.

The initial measured findings and prioritized renderer gaps are recorded in [BASELINE.md](BASELINE.md).

## Licensing and corpus boundary

The corpus is declared in `corpus.ts`. At runtime, every path must pass
`git ls-files --error-unmatch`, resolve to a regular (non-symlink) file, and have the same Git blob
hash as `HEAD`. Absolute paths, paths outside the checkout, missing files, modified tracked files,
and untracked files are rejected before either renderer starts. The benchmark references existing
fixtures in place. It does not copy a third-party corpus or any ignored/untracked harness into the
repository.

Generated PDFs, PNGs, overlays, and JSON must be written outside the checkout. The runners reject
an output directory inside the repository, even if that directory is ignored. A local run starts
with a new empty directory so stale pages cannot contaminate it. CI initializes only the generated-
PDF root with a recognized `ci-context.json` and pending `index.html`; the generated-PDF runner owns
the remaining contents and replaces the pending viewer as evidence becomes available.

## PDF-to-PDF measurement flow

For every selected corpus entry, the generated-PDF runner performs one ordered, bounded flow:

1. Verify that the fixture is a regular Git-tracked file whose worktree bytes equal `HEAD`.
2. Convert those bytes with the public `@docxodus/export` API and its package-owned pinned Chromium,
   requesting the explicit revision and comment profiles declared by the case.
3. Export the same comparison fixture to a reference PDF with the contracted LibreOffice build and
   a fresh user profile.
4. Preserve both original PDFs and record their SHA-256 digests. PDF byte identity is diagnostic,
   not a cross-run promise, because Chromium records volatile PDF metadata.
5. Parse both PDFs before rasterization. Page count and every MediaBox/CropBox origin and physical
   dimension are recorded in points and remain hard signals separate from image similarity.
6. Rasterize both PDFs with the same `pdftoppm` executable, DPI, page-box choice, antialiasing, and
   color-mode argument vector. The complete argument contract and Poppler version are recorded in
   the report; no engine receives a renderer-specific raster option.
7. Compare paired pages for SSIM, perceptual color difference, tolerant ink precision/recall/F1,
   and a reviewable overlay. Each PDF, page raster, and overlay carries a SHA-256 digest.
8. Extract selectable text and link annotations independently of the raster. Expected visible and
   hidden text follows the case's revision/comment profile; supported external targets must match,
   and every internal destination must resolve to a page in the generated PDF.

A page-count mismatch, out-of-contract physical page, semantic extraction failure, unresolved link,
or conversion error is severe regardless of a visually white or similar page. Pixel severity and
the reviewed disposition contract below decide whether a remaining visual difference gates strict
mode.

## Deterministic rendering contract

- The generated-PDF run uses the supported Node API and the exact Chromium revision owned by
  `@docxodus/export`; its render report records the production renderer/font fingerprint. The
  separate browser-page benchmark continues to use Chromium at device scale 1 and pagination
  scale 1.
- LibreOffice exports PDF from a fresh per-document user profile. The generated and reference PDFs
  are rasterized by the same Poppler executable at exactly 96 DPI and under one recorded color and
  antialiasing contract.
- Both processes use `C.UTF-8`, UTC, and the **font-substitution contract** below instead of the
  host's default fontconfig, so line wraps cannot drift with whatever fonts a host happens to
  carry. The summary records the Chromium, LibreOffice, and Poppler versions plus every contract
  resolution (family, file, font version, contract-file SHA-256).
- The comparison uses final-revision view: insertions are included and deletions/move markup are not
  rendered. LibreOffice's headless PDF filter follows the file's saved redline-display state and
  provides no final-view switch, so manifest cases marked `revisionMode: 'accepted'` are accepted
  once into a temporary DOCX outside the checkout; both engines then render those identical bytes.
  Each generated-PDF case records its explicit revision/comment profile. The older browser-page
  diagnostic keeps comments and Docxodus annotations disabled. Headers, footers, footnotes, and
  endnotes remain enabled.
- Generated PDF printing crosses the production readiness barriers for fonts, images, charts/SVG,
  pagination, running stories, and stable page geometry, then repeats them after reopening the exact
  serialized standalone document. The browser-page diagnostic separately waits for
  `document.fonts.ready`, every image load, and two animation frames. Animations, transitions,
  carets, page shadows, page labels, and page gaps are disabled in diagnostic capture.
- Pages are paired by one-based page index. Page count and page dimensions remain independent hard
  signals. A bounded translation search of only ±2 pixels normalizes raster-origin rounding; the
  chosen offset is always reported.
- No masks are applied. A future mask must be a bounded rectangle tied to one manifest case/page and
  carry a specific justification; a document-wide or text-wide mask is not acceptable.

## Font-substitution contract

Documents declare proprietary Office families no CI host may install. Which substitute each
family resolves to changes wrapping and line metrics, and a renderer-only fallback was tried and
rejected — it moved one engine without the other. Font policy is therefore a **shared contract**
(issue #379): `fonts.conf` pins each declared family to a license-safe metric-compatible
substitute, and both renderers load it via `FONTCONFIG_FILE` — LibreOffice through the runner's
subprocess environment, Chromium at browser launch through `playwright.config.ts` (scoped to the
benchmark opt-in so ordinary specs keep host fonts).

| Declared family | Substitute | Package | Metric-compatible |
|---|---|---|---|
| Calibri | Carlito | fonts-crosextra-carlito | yes |
| Calibri Light | Carlito | fonts-crosextra-carlito | no — documented approximation, no open metric clone exists |
| Cambria | Caladea | fonts-crosextra-caladea | yes |
| Times New Roman | Liberation Serif | fonts-liberation2 | yes |
| Arial | Liberation Sans | fonts-liberation2 | yes |
| Courier New | Liberation Mono | fonts-liberation2 | yes |

Enforcement is layered, and each layer fails with a message naming what to fix:

- `assertFontContract()` fails the run when `fc-match` under the contract does not resolve every
  family to its substitute, naming the package to install.
- An in-browser check fails when Chromium was launched without the contract (canvas advance
  widths of each family must equal its substitute's).
- A **wrapping probe** (`declared font families wrap identically…` in `visual-parity.spec.ts`)
  renders one generated paragraph per family through both engines and requires identical
  line counts — fc-match proves what fontconfig would resolve; the probe proves what the two
  renderers actually did. It is negatively validated: without the contract, Calibri Light wraps
  differently on a stock Ubuntu host.

With the contract pinned and recorded, a baseline delta traces to either the renderer or a
declared contract change — never to silent host-font drift. `environment` dispositions now mean
"the two engines lay out the SAME substitute differently" (rasterization, justification, line
breaking), not "the two engines picked different fonts".

## Reference-version contract

With fonts pinned, the LibreOffice version was the last uncontracted variable in the comparison
(issue #403): different LibreOffice releases render the same document differently, so a
runner-image bump could shift weekly numbers and masquerade as a renderer change. The benchmark
is therefore contracted to **LibreOffice 25.8** (exact major.minor; known-good build 25.8.7.3),
declared once in `environment-contract.ts` and enforced twice:

- `assertLibreOfficeContract()` fails the run **at start** when the host's LibreOffice is not
  the contract minor, naming the version found, the TDF archive to install, the bundled-font
  removal step, and the known cross-version differences — mirroring the font contract's failure
  mode. Fail fast beats discovering the mismatch twenty minutes later at the ratchet.
- The ratchet's environment fingerprint (below) refuses to compare numbers across LibreOffice,
  Chromium, Poppler, or font-contract changes, so even a deliberately out-of-contract run
  (`DOCXODUS_VISUAL_PARITY_ALLOW_VERSION_DRIFT=1`, for exploratory cross-version reproduction)
  can never be misread as a renderer regression, and can never refresh the record.

CI installs the contract build from the TDF archive rather than `ubuntu-latest` apt (which
carries whatever the runner image's Ubuntu shipped — 24.2 on 24.04). Before extraction, CI
verifies the archive's detached TDF signature and requires the exact pinned issuer fingerprint
`C2839ECAD9408FBE9531C3E9F434A1EFAFEEAEA3`. The pure spec
`visual-parity-ratchet.spec.ts` proves the failure message and asserts the declared version,
the committed record's fingerprint, and the CI pin cannot drift apart, on every pull request,
without LibreOffice installed.

The 25.8 line is archived and no longer receives security updates. It is retained only because a
visual ratchet must hold its reference renderer constant; CI uses it in a dedicated, read-only job
against repository-pinned fixtures, with no untrusted document input. A future reference-version
bump must update the archive URL, detached signature, key fingerprint, and measured records in one
reviewed change.

## Required tools and version identity

The measurement surface intentionally names every external component:

| Component | Contract used by CI |
|---|---|
| Node.js | 22; the export package itself requires at least 22.13 |
| .NET / WASM tools | .NET 10.0.x plus the `wasm-tools` workload |
| Docxodus Chromium | `@playwright/browser-chromium` 1.57.0 from `npm-export/package-lock.json` |
| PDF parsing | `pdf-lib` 1.17.1 and `pdfjs-dist` 6.2.108 |
| LibreOffice | exact known-good build 25.8.7.3; accepted fingerprint 25.8 |
| Poppler | `pdftoppm` and `pdftotext` from the same `poppler-utils` install |
| Font discovery | Fontconfig `fc-match` plus Carlito, Caladea, and Liberation contract fonts |

`npm ci` makes JavaScript package versions exact through the committed lockfiles. Poppler comes from
the CI runner's apt repository, so its major/minor is part of the ratchet environment fingerprint;
a runner-image change fails as `environment-changed` instead of being attributed to Docxodus.
Reproducing an existing recorded number requires the Poppler fingerprint in `ratchet.json`, not
merely any `pdftoppm` binary. Reports retain the full version banners and measurement arguments so
an artifact can be recreated without guessing which tools were used.

Known cross-version rendering differences the corpus is sensitive to (each new finding joins
`KNOWN_LIBREOFFICE_VERSION_DIFFERENCES` so the failure message teaches it):

| Versions | Difference | Sensitive cases |
|---|---|---|
| 24.2 vs ≥ 25.8 | Footnote separator width: 24.2 draws its legacy 25%-of-column separator where 25.8 draws the two-inch Word default. | `footnote` |

## Metrics and thresholds

Every paired page reports:

- exact RGB diff ratio (diagnostic only; antialiasing makes it unsuitable as a gate);
- CIE Lab ΔE76 mean and the ratio above ΔE 2.3, a conventional just-noticeable threshold;
- block SSIM over 8×8 luminance windows;
- ink precision, recall, and F1 after a two-pixel dilation, which tolerate glyph antialiasing
  without tolerating missing, extra, or displaced layout;
- raster width/height deltas and the bounded alignment offset; and
- for the generated-PDF path, inherited `UserUnit`/`Rotate` plus scaled, orientation-aware
  MediaBox/CropBox origins and dimensions in physical PDF points before either artifact is rasterized.

Severity is assigned by the worst signal:

| Severity | SSIM | Ink F1 | ΔE>2.3 ratio | Geometry |
|---|---:|---:|---:|---:|
| close | ≥ 0.98 | ≥ 0.95 | ≤ 0.02 | ≤ 1 px |
| minor | ≥ 0.95 | ≥ 0.90 | ≤ 0.05 | ≤ 1 px |
| major | ≥ 0.85 | ≥ 0.75 | ≤ 0.15 | ≤ 1 px |
| severe | below major | below major | above major | > 1 px |

A page-count mismatch is always severe. The generated-PDF path unconditionally fails conversion,
API/report/PageMap/PDF binding, page-count, MediaBox/CropBox, selectable-text, exact logical-link,
and chart-vector errors; raster similarity or a reviewed disposition cannot excuse a broken hard
signal. `DOCXODUS_VISUAL_PARITY_STRICT=1`
additionally turns renderer-attributable severe raster cases into a failing gate (see the disposition
contract below). This keeps the baseline useful while known environment deltas and reference
deviations are tracked without waiving PDF correctness. Independently of strict mode, every run
compares itself against the committed regression record described next.

## Regression ratchet

Full-strict mode stays unreachable while renderer-attributable severe cases remain, but *"no case
may get worse than recorded"* is enforceable today (issue #395). `ratchet.json` is a committed,
numbers-only record — one row per case: page counts, severity, mean SSIM, worst ink F1, and the
disposition. No images, no paths, no artifact hashes, so it stays reviewable as a diff.

The ratchet is deliberately **broader than strict mode**: strict gates only severe cases the
renderer owns, while the ratchet covers every case at every severity. A `close` case sliding to
`minor` is exactly the drift the weekly run existed to catch and previously could not — its
artifact expired in 14 days and nothing compared one run to the next.

| Signal | Fails when |
|---|---|
| page count | either engine's page count changes at all |
| severity | the case moves down the severity ladder |
| mean SSIM | falls more than `tolerance.ssim` (0.0005) below the record |
| worst ink F1 | falls more than `tolerance.inkF1` (0.001) below the record |
| conversion | the case errors where the record has none |
| physical geometry | a generated PDF fails the absolute MediaBox/CropBox contract |
| semantics | generated-PDF selectable text or supported hyperlink targets fail |
| coverage | a recorded case is missing, or a measured case is unrecorded |

The tolerances are tight because within one environment the benchmark is *deterministic* — two
clean passes produced identical metrics and identical SHA-256s for all 60 images. They are not
zero because the environment fingerprint is coarser than the environment itself. Every renderer
movement BASELINE.md records for a real fix is at least an order of magnitude larger; the smallest
is the two-inch footnote separator at +0.000128 SSIM and +0.003537 ink F1.

**Environment fingerprint.** Across environments the numbers move materially — LibreOffice 24.2
draws a different footnote separator than 25.8. The record therefore carries the LibreOffice
major.minor (contract-pinned since issue #403), the Chromium major, the Poppler major.minor
(pdftoppm's rasterizer sits between the reference PDF and every number), and `fonts.conf`'s
SHA-256. When they do not match, the run
reports **`environment-changed`** and demands a refresh instead of claiming a regression, so a
reference-renderer release can never be blamed on Docxodus. That outcome still fails the run: a
stale record silently comparing across environments is worse than an explicit demand to refresh it.

**Updating the record** is a deliberate act in the PR that changes rendering, so improvements and
accepted regressions are reviewed in the diff:

```bash
DOCXODUS_VISUAL_PARITY_OUTPUT=/tmp/docxodus-visual-parity \
DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1 \
npm run test:visual-parity
```

The generated-PDF record is distinct and uses its own complete run:

```bash
DOCXODUS_GENERATED_PDF_PARITY_OUTPUT="$(mktemp -d)" \
DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD=1 \
  npm --prefix npm run test:generated-pdf-parity
```

The update path refuses a filtered run and any dirty or unverified worktree. Commit the
implementation first, then generate the record from that exact clean commit so `sourceCommit`
identifies reproducible source rather than merely the base of local edits. A passing run still lists
every improvement it measured, so a stale record announces itself. Set
`DOCXODUS_VISUAL_PARITY_RATCHET=0` to observe without gating; the comparison is reported either
way. A filtered run compares only the cases it measured.

`sourceCommit` identifies the historical implementation that actually produced the stored numbers;
it is not rewritten to a rebased commit by analogy. Each current run separately records `gitCommit`
and `gitTree`, builds both packages itself, and compares the resulting PDFs to that historical
ratchet. Refresh `sourceCommit` only through the complete clean update command above, after every
hard/strict/evidence gate has passed.

The comparison layer (`ratchet.ts`) is pure, and `visual-parity-ratchet.spec.ts` exercises it on
**every** pull request — no LibreOffice, no renderer. That is what keeps "a deliberately introduced
regression fails, naming the case and the signal" a continuously proven property instead of a
claim demonstrated once by hand.

## Attribution dispositions

Severity measures *how different* two renderings are; it cannot say *whose difference* it is. Every
corpus entry therefore carries a reviewed `disposition` — an attribution claim with a mandatory
rationale and, where one exists, a tracking issue:

| Kind | Meaning | Gates strict? |
|---|---|---|
| `renderer-bug` | An established Docxodus rendering defect. | yes |
| `unattributed` | Not yet triaged; the safe default for new corpus entries. | yes |
| `environment` | Dominated by the comparison environment (Chromium vs LibreOffice font substitution, wrapping, line metrics), not OOXML geometry. | no |
| `reference-deviation` | Docxodus follows the OOXML evidence; LibreOffice deviates. | no |
| `unsupported-feature` | A known unimplemented feature tracked as a feature gap. | no |

A conversion error always gates regardless of disposition. Dispositions live in `corpus.ts` so they
are code-reviewed alongside the corpus, flow into `metrics.json`/`summary.json`
(`aggregate.severeByDisposition`, `aggregate.strictGatingCases`), and must be updated when a fix or
new evidence changes the triage. A disposition is a claim about the *dominant residual*
discrepancy — it never justifies masking, and changing one to make a gate pass requires the same
evidence bar as a renderer fix: OOXML semantics, Word behavior where available, and a reduced case.

## Word-reference evidence

"Word behavior where available" used to be an inference. `word-reference.json` (issue #402) makes
it recorded data: a committed, numbers-only store of what Microsoft Word renders for each corpus
fixture — page counts, page geometry, ink extents, and named per-case measurements — plus the
Word/OS versions they were taken under. The only manual step is exporting each fixture to PDF
with a licensed Word; `npm run capture:word-reference` automates everything downstream with the
benchmark's own contract (Poppler at 96 DPI, the shared ink model). See
[WORD_REFERENCE.md](WORD_REFERENCE.md) for the procedure, and for the honesty boundary: Word
renders with genuine Office fonts, not the contract substitutes, so Word evidence decides
STRUCTURAL questions (spacing suppressed or painted, page counts, block positions) while pixel
scores against Word stay advisory.

Every corpus case has a row; new cases enter `pending`, and the pure spec
`visual-parity-word-reference.spec.ts` keeps the store consistent with the corpus on every pull
request. A disposition citing Word data does it in `disposition.wordEvidence`, which the spec
refuses unless the cited case is actually `measured` — a rationale can never claim Word evidence
that was never captured. `summary.json` records each run's evidence coverage under
`wordReference`.

## Run locally

Prerequisites: a Writer-capable LibreOffice of the **contract minor** (see the reference-version
contract above; a bare `libreoffice-core` install fails every case with "source file could not
be loaded", and an out-of-contract version fails at run start with install guidance),
`poppler-utils` (`pdftoppm`/`pdftotext`), the contract fonts (`fonts-crosextra-carlito`,
`fonts-crosextra-caladea`, `fonts-liberation2` — the run fails with install instructions when
missing), Fontconfig (`fc-match`), .NET 10 with the `wasm-tools` workload, and the repository's two
npm package dependency sets.

When your distro does not package the contract minor, use the TDF archive build named in the
failure message — but TDF-packaged builds bundle their own Caladea/Carlito/Liberation copies
under `share/fonts/truetype/`, which silently override the font contract inside LibreOffice
only. Remove the bundled duplicates so both engines resolve the same system fonts (the wrapping
probe fails naming the family if you forget — see the issue-#400 baseline entry).
Verify the adjacent `.asc` signature with the exact key fingerprint printed by the failure message
before extracting or installing the archive.

Install the exact locked packages from the repository root. The companion's normal install fetches
its pinned Chromium revision; do not use `--ignore-scripts` for this benchmark. The public PDF gate
rebuilds both packages itself, in npm-then-exporter order, so a clean source commit cannot be paired
with stale ignored `dist` bytes.

```bash
npm --prefix npm ci
npm --prefix npm-export ci
npm --prefix npm exec -- playwright install --with-deps chromium
```

Run the PDF release gate into a fresh directory:

```bash
DOCXODUS_PARITY_OUTPUT="$(mktemp -d)"
DOCXODUS_GENERATED_PDF_PARITY_OUTPUT="$DOCXODUS_PARITY_OUTPUT" \
  npm --prefix npm run test:generated-pdf-parity
echo "Artifact viewer: $DOCXODUS_PARITY_OUTPUT/index.html"
```

The older browser-page diagnostic remains independently reproducible:

```bash
DOCXODUS_BROWSER_PARITY_OUTPUT="$(mktemp -d)"
DOCXODUS_VISUAL_PARITY_OUTPUT="$DOCXODUS_BROWSER_PARITY_OUTPUT" \
  npm --prefix npm run test:visual-parity
```

Run selected manifest IDs:

```bash
DOCXODUS_GENERATED_PDF_PARITY_FILTER=pdf-footnote,pdf-chart \
DOCXODUS_GENERATED_PDF_PARITY_OUTPUT="$(mktemp -d)" \
  npm --prefix npm run test:generated-pdf-parity
```

## Artifacts and failure retention

Open `index.html` first. It links the run context and summary, both original PDFs, side-by-side page
rasters, red perceptual-difference overlays, physical geometry, selectable-text/link evidence, and
per-case metrics. Artifact paths inside JSON are relative to the output root so the downloaded
directory remains portable. Source, PDF, raster, and overlay evidence carries SHA-256 digests, and
the summary records the source commit and tree, whether the worktree was dirty, the bounded built
exporter graph, asset manifest, module entries, exact executable hashes, and package lock hashes.
Those executable and build identities are rechecked after the run before a record can be updated.

The generated-PDF runner writes progress incrementally. A failed case retains all PDFs and page
evidence produced before the failure and receives structured failure evidence rather than being
deleted. Pull-request, manual, and scheduled CI create `ci-context.json` and a pending viewer before installing external tools, so even
an apt, LibreOffice, npm, build, or run-start failure leaves a viewable artifact that names the
commit and workflow run. `.github/workflows/visual-parity.yml` uploads this root with `if: always()`,
retains it for 14 days, and treats a missing artifact root as an error. The browser-page report is
uploaded separately, also on failure; one benchmark cannot suppress the other's evidence.

The generated-PDF test runs even when the preceding browser-page benchmark fails, unless the job was
explicitly cancelled. Both failures still contribute to the job result. A timeout may interrupt the
currently active case, but the initialized viewer and every previously finalized case remain in the
artifact.

Playwright retries write to `retry-N/` subdirectories beneath the configured artifact root. The
first attempt remains available through the root viewer, and a retry cannot replace the original
failure with a stale-output error.

## Triage rules

LibreOffice is a comparison implementation, not an oracle. Start with severe cases, inspect the two
page images and overlay, then reduce the discrepancy to an independently generated minimal OOXML
document. A fix belongs in the renderer only when the OOXML semantics, Word metadata/behavior where
available, and the reduced case support it. Confirmed fixes require generated regression tests and a
before/after benchmark rerun.
