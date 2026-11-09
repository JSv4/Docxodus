# Separate rendering references

The renderer currently has one layout mode, recorded as `default`. Word and
LibreOffice are independent reference providers. They have separate immutable
reference bundles, main/candidate measurements and reports; their scores are
never combined into a shared series. This adds no public layout profile option.

The original LibreOffice oracle, pinned historical baseline, reference capture,
100-point calibration and zero-tolerance comparison stay unchanged. The original
`check-regressions.py` is retained. `check-provider-regressions.py` adds the Word
authority policy in `word-reference-policy.json`:

- Every candidate case still has to pass all mandatory checks, including the
  eight cases with Word evidence.
- Unmapped mandatory failures on current main are reported in
  `baselineMandatoryFailures` and may be repaired by the candidate. Mapped
  original inputs must pass on main as well as the candidate to establish Word
  authority. Historical baseline protection is unchanged.
- All five Word page metrics must hold or improve against current main, using
  the same registered Word bundle. Every Word case must score at least 95.
- Word authority covers only the exact eight original/captured input hashes in
  the policy. CI checks the immutable pack hash, source/program font bindings,
  original/captured source identities and renderer commit identities.
- Both main and candidate must produce byte-identical full-page PPM rasters from
  each original CI input and its Word capture input in the current run. Changing
  a font-package/schema-order probe cannot silently change the evidence scope.
- Only after those checks pass may those eight cases classify a LibreOffice
  similarity decline as a reviewed provider difference. All LibreOffice findings
  remain in `provider-comparison.json` and the unchanged oracle report. Every
  other passing historical baseline case retains its zero-tolerance protection.

The Word references use original authored inputs with separately licensed
Liberation Serif embedding.
Word's internal font aliases are qualified by exact reviewed source/native font
program hashes. The glyph audit is an operator attestation; the harness repeats
the cryptographic bindings, not the outline audit. No fonts are renamed and no
PDF coordinates or pixels are normalized. The immutable capture pack retains its
OFL notice and provenance outside the source tree. [Benchmark PR107](https://github.com/JSv4/docx-fidelity-bench/pull/107)
contains full-page screenshots, raw diagnostic results and limitations for all
eight cases. Strict Word CI measurement is required before merge readiness.

The Word reference pack is published in `JSv4/docx-fidelity-bench` and pinned by
URL, size and SHA-256 in the policy. CI registers it in each measurement environment.
Both provider reports, comparisons and registered references are retained in the
rendering artifact. The two pinned harnesses have byte-identical build recipes
and metric source, verified before sharing source-bound renderer builds.

## Cached measurements

The pinned baseline never changes and main changes once per merge, so the
`warm` job measures them on each push to main and caches, in main's scope, the
LibreOffice reference, the registered Word bundle, the baseline result and
main's two results. A pull request restores them and builds and measures only
its candidate, which is always built fresh.

- Every key carries `measurement-key.py`'s environment key: the benches' own
  environment contract (tool versions and hashes, resolved font file hashes) for
  both suites, the LibreOffice reference build, the runner image and the
  measurement packages. Main's results are keyed under the reference they were
  measured against. A miss of either measures it live, exactly as before.
- The benches refuse to compare results whose measurement contracts differ, and
  LibreOffice is recalibrated live against the cached reference in every run.
- The `canary` job re-measures the cached baseline live beside every pull
  request; `check-cached-results.py` fails unless it reproduces exactly, apart
  from PDF byte hashes (PDFs embed per-run metadata).
- Bump `CACHE_VERSION` in the workflow to discard every cached entry, or run
  the workflow with `mode: warm` from main to re-measure main.

Run the gate contract tests with the pinned Word harness on `PYTHONPATH`:

```sh
PYTHONPATH=/path/to/docx-fidelity-bench python3 -m unittest discover \
  -s tools/rendering-fidelity -p 'test_*.py' -v
```

Future actual Word and LibreOffice renderer modes will each need their own
accepted measurements against the corresponding provider bundle. Current
results' explicit layout identity prevents those modes from reusing each
other's baseline series.
