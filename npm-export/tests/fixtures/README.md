# Font security fixtures

These tiny fonts are generated from original geometric outlines authored for
the Docxodus test suite. They contain only space and `A`–`Z`; no host font file
or third-party outline is copied. The repository MIT license applies.

The `synthetic-carlito*.ttf` fixtures intentionally report the family `Carlito` so the
frozen Calibri substitution contract can be exercised without redistributing
or modifying the Reserved Font Name Carlito software. It is not a copy or
derivative of the Carlito typeface. The narrow and wide variants deliberately
change only test-outline advance widths to prove PageMap reflow from one source.

`docxodus-policy-base.ttf` is copied to a private temporary directory and its
OS/2 `fsType` bits are changed in-memory by the tests for restricted-license
and bitmap-only cases. `docxodus-load-failure.ttf` keeps readable name, cmap,
OS/2, and metric tables but has an impossible `glyf` table length so Chromium's
OpenType sanitizer deterministically rejects it after Node metadata discovery.

Regenerate with FontTools 4.63.0:

```sh
python3 npm-export/tests/fixtures/generate-font-fixtures.py
```

Expected SHA-256 digests:

- `synthetic-carlito.ttf`: `3013ff0d837c3cd054eb8987896087087fe8c0023e5d375fd05c1fc6abe2db6d`
- `synthetic-carlito-narrow.ttf`: `45b4c534c7cdfcd66421d48a18cd9473663a850a5fd237e4d522d8926deed39d`
- `synthetic-carlito-wide.ttf`: `6f5efcfc8a950b9dd19f08c7cfb8ac2506c6f05998d757e5f6be1ef3622ad5dd`
- `docxodus-policy-base.ttf`: `97a5c18c38e9d155651cc52f98fbe5a3e010363ead3c2d782041a499bcb0a43e`
- `docxodus-metric-test.ttf`: `52e60f325816808bd1049c8fa43887bf601287c97b73bc7b6f32be9ed1b399c2`
- `docxodus-metric-test.woff`: `030d8630b82db745f20811d775cb89f7569e82a78660ee66a421dfd8386a129f`
- `docxodus-load-failure.ttf`: `0ba0bbfde0e4070f02c87b8185c97e0a5fb6e1378209e0550680cbe3352d815a`
