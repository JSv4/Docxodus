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
- Both processes use `C.UTF-8`, UTC, and the **shared font-substitution contract** below. The
  summary records the Chromium, LibreOffice, and Poppler versions, and the exact font file each
  declared family resolved to in this run.
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

Neither engine ships Microsoft's Office fonts, so every Office family a fixture names is
substituted. When the two engines substitute *differently*, their line breaking and baselines
differ for reasons that have nothing to do with Docxodus — and a benchmark that cannot tell that
apart from a renderer regression is not measuring the renderer.

The policy is therefore **shared by both engines**, not implemented in one of them. A
Chromium-only `Calibri Light → Carlito` fallback was tried and rejected: it made one engine change
its mind, so the two disagreed *more* than before (accepted-revision SSIM 0.93177 → 0.92905, ink F1
0.46817 → 0.42477). Chromium and LibreOffice both resolve families through fontconfig on Linux, so
one fontconfig fragment governs both by construction.

| Family | Substitute | Package | Metric clone |
|---|---|---|---|
| Calibri | Carlito | `fonts-crosextra-carlito` | yes |
| Cambria | Caladea | `fonts-crosextra-caladea` | yes |
| Times New Roman | Liberation Serif | `fonts-liberation2` | yes |
| Arial | Liberation Sans | `fonts-liberation2` | yes |
| Courier New | Liberation Mono | `fonts-liberation2` | yes |
| Calibri Light | Carlito | `fonts-crosextra-carlito` | **no** |

Calibri Light has no metric-compatible free clone — Carlito ships no Light weight. Left unbound it
falls through to whichever generic sans each engine happens to prefer, which is exactly the
non-determinism the contract removes. Binding it is a choice about *agreement*, not fidelity: text
set in Calibri Light will not wrap where Word wraps it, and that must not be "corrected" in one
renderer.

The contract lives in exactly two places, and `font-contract.spec.ts` fails if they drift apart:

- `fontconfig/60-docxodus-office-substitutes.conf` — what the engines read.
- `fonts.ts` — what the run reports and verifies (`FONT_SUBSTITUTION_CONTRACT`).

No font is bundled with this repository, and nothing here affects the library at run time.

### Applying it

The benchmark applies the fragment itself: it writes a fontconfig root outside the checkout that
layers the fragment over the host's own configuration, and points both Chromium (through the
browser launch environment) and every LibreOffice subprocess at it. Nothing is written into your
home directory or `/etc`. You only need the packages:

```bash
sudo apt-get install -y fonts-liberation2 fonts-crosextra-carlito fonts-crosextra-caladea
```

To install the contract permanently instead — which is what CI does, and what makes other tools on
the machine agree too:

```bash
sudo cp npm/tests/visual-parity/fontconfig/60-docxodus-office-substitutes.conf /etc/fonts/conf.d/
fc-cache -f
fc-match "Calibri Light"   # => Carlito
```

Then run with `DOCXODUS_VISUAL_PARITY_HOST_FONTS=1` to measure the host configuration as it is,
rather than layering the fragment again.

### When the contract is unavailable

The run resolves every declared family through `fc-match` before either engine starts. If any
family resolves to the wrong substitute the run **skips** with a message naming the family and the
package to install — and **fails** instead under `DOCXODUS_VISUAL_PARITY_STRICT=1`. It never
proceeds and reports numbers produced by an unknown font environment.

### The drift probe

A generated probe document — one short, non-wrapping line and one long, wrapping paragraph per
declared family — is rendered by both engines before the corpus. The comparison is between the two
**engines**, not against a stored expectation:

- the short line's **advance** must agree within 3 px. This is pure font resolution, so the
  tolerance can sit just above hinting noise; a different face moves it by tens of pixels.
- the page's total **line count** must match, which is what a different face does to a paragraph
  long enough to wrap.

Break *positions* are deliberately not compared. With the contract satisfied and both engines
confirmed on Caladea, the Cambria paragraph's widest line still ends 34 px apart: the engines break
identically-measured text differently. That is a real difference and worth knowing, but it is not
font drift, and a probe that conflates the two explains nothing.

A probe failure is reported as *font environment drift, not a renderer regression*, and the
per-family advances are recorded in `summary.json` so two reports can be compared directly.

`font-contract.spec.ts` runs in the ordinary browser suite (no LibreOffice needed) and pins the
fragment against `fonts.ts`, the drift detector's sensitivity, and its noise tolerance.

Committed-screenshot specs such as `tabs-visual.spec.ts` are **not** governed by this contract:
they compare against images baked on one machine, so they remain environment-sensitive by design.

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
the complete report; `DOCXODUS_VISUAL_PARITY_STRICT=1` turns severe cases into a failing gate. This
keeps the initial baseline useful while known unsupported content is being triaged.

## Run locally

Prerequisites: `libreoffice`, `pdftoppm`, `fontconfig`, the contract's font packages (see
[Font-substitution contract](#font-substitution-contract)), Chromium installed for Playwright, and
the repository's npm dependencies.

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
