# LibreOffice pixel-parity benchmark

This opt-in benchmark renders a small, stratified set of repository-tracked DOCX fixtures through
Docxodus and LibreOffice, rasterizes every page at 96 DPI, and writes page images, heatmaps, and
machine-readable metrics outside the repository.

It extends the single-frame arcade interoperability test into a document-wide diagnostic. It is not
part of the normal browser suite because LibreOffice and Poppler are host tools.

The initial measured findings and prioritized renderer gaps are recorded in [BASELINE.md](BASELINE.md).

## Licensing and corpus boundary

The corpus is declared in `corpus.ts`. At runtime, every path must pass
`git ls-files --error-unmatch`, resolve to a regular (non-symlink) file, and have the same Git blob
hash as `HEAD`. Absolute paths, paths outside the checkout, missing files, modified tracked files,
and untracked files are rejected before either renderer starts. The benchmark references existing
fixtures in place. It does not copy a third-party corpus or any ignored/untracked harness into the
repository.

Generated PNGs, overlays, and JSON must be written outside the checkout. The runner rejects an
output directory inside the repository, even if that directory is ignored, and rejects a non-empty
output directory so stale pages cannot contaminate a rerun.

## Deterministic rendering contract

- Docxodus uses Chromium at device scale 1 and pagination scale 1: one CSS pixel equals one 96-DPI
  raster pixel.
- LibreOffice exports PDF from a fresh per-document user profile. Poppler rasterizes the PDF at
  exactly 96 DPI.
- Both processes use `C.UTF-8`, UTC, and the **font-substitution contract** below instead of the
  host's default fontconfig, so line wraps cannot drift with whatever fonts a host happens to
  carry. The summary records the Chromium, LibreOffice, and Poppler versions plus every contract
  resolution (family, file, font version, contract-file SHA-256).
- The comparison uses final-revision view: insertions are included and deletions/move markup are not
  rendered. LibreOffice's headless PDF filter follows the file's saved redline-display state and
  provides no final-view switch, so manifest cases marked `revisionMode: 'accepted'` are accepted
  once into a temporary DOCX outside the checkout; both engines then render those identical bytes.
  Comments and Docxodus annotations are disabled; headers, footers, footnotes, and endnotes are
  enabled.
- Chromium waits for `document.fonts.ready`, every image load, and two animation frames. Animations,
  transitions, carets, page shadows, page labels, and page gaps are disabled.
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

## Metrics and thresholds

Every paired page reports:

- exact RGB diff ratio (diagnostic only; antialiasing makes it unsuitable as a gate);
- CIE Lab ΔE76 mean and the ratio above ΔE 2.3, a conventional just-noticeable threshold;
- block SSIM over 8×8 luminance windows;
- ink F1 after a two-pixel dilation, which tolerates glyph antialiasing without tolerating displaced
  layout;
- page width/height deltas and the bounded alignment offset.

Severity is assigned by the worst signal:

| Severity | SSIM | Ink F1 | ΔE>2.3 ratio | Geometry |
|---|---:|---:|---:|---:|
| close | ≥ 0.98 | ≥ 0.95 | ≤ 0.02 | ≤ 1 px |
| minor | ≥ 0.95 | ≥ 0.90 | ≤ 0.05 | ≤ 1 px |
| major | ≥ 0.85 | ≥ 0.75 | ≤ 0.15 | ≤ 1 px |
| severe | below major | below major | above major | > 1 px |

A page-count mismatch is always severe. Default runs are observational and succeed after producing
the complete report; `DOCXODUS_VISUAL_PARITY_STRICT=1` turns renderer-attributable severe cases into
a failing gate (see the disposition contract below). This keeps the baseline useful while known
environment deltas and reference deviations are tracked without blocking.

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

## Run locally

Prerequisites: `libreoffice-writer` (a bare `libreoffice-core` install fails every case with
"source file could not be loaded"), `poppler-utils` (`pdftoppm`/`pdftotext`), the contract fonts
(`fonts-crosextra-carlito`, `fonts-crosextra-caladea`, `fonts-liberation2` — the run fails with
install instructions when missing), Chromium installed for Playwright, and the repository's npm
dependencies.

```bash
cd npm
npm run build
DOCXODUS_VISUAL_PARITY_OUTPUT=/tmp/docxodus-visual-parity npm run test:visual-parity
```

Run selected manifest IDs:

```bash
DOCXODUS_VISUAL_PARITY_FILTER=text-formatting,merged-table \
DOCXODUS_VISUAL_PARITY_OUTPUT=/tmp/docxodus-visual-parity \
npm run test:visual-parity
```

The output contains `summary.json` plus one directory per case. Each case contains the two engine
PNGs, a red perceptual-difference heatmap, and `metrics.json`. Artifact paths inside JSON are relative
to the output root so an uploaded report remains portable. Every artifact also carries a SHA-256
digest, and the summary records whether the source worktree was dirty, so repeated runs can be
compared without trusting filenames or silently mixing source states.

## Triage rules

LibreOffice is a comparison implementation, not an oracle. Start with severe cases, inspect the two
page images and overlay, then reduce the discrepancy to an independently generated minimal OOXML
document. A fix belongs in the renderer only when the OOXML semantics, Word metadata/behavior where
available, and the reduced case support it. Confirmed fixes require generated regression tests and a
before/after benchmark rerun.
