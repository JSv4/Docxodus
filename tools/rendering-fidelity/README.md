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
rendering artifact. Main's exact build is cached separately; the candidate is
always built fresh. The two pinned harnesses have byte-identical build recipes
and metric source, verified before sharing source-bound renderer builds.

Run the gate contract tests with the pinned Word harness on `PYTHONPATH`:

```sh
PYTHONPATH=/path/to/docx-fidelity-bench python3 -m unittest discover \
  -s tools/rendering-fidelity -p test_provider_regressions.py -v
```

Future actual Word and LibreOffice renderer modes will each need their own
accepted measurements against the corresponding provider bundle. Current
results' explicit layout identity prevents those modes from reusing each
other's baseline series.
