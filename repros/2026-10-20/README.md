# DOCX reproductions

These thirteen inputs reproduce tab conversion, drawing, footnote, revision
resolution, and field rendering failures. `create_inputs.py` generates the DOCX files.

| Inputs | Observed behavior |
| --- | --- |
| `tabs-0`, `tabs-12` | HTML conversion throws `Sequence contains no matching element`; the v14 PDF API reports `conversion_failure: NoMatch`. |
| `tabs-720` | The same paragraph converts and exports with ordinary automatic tab spacing. |
| `shape-plain`, `shape-deleted`, `group-plain`, `group-deleted` | The rectangle is absent from HTML; v14 PDF export fails because its paragraph has no measurable fragment. |
| `note-arial20` | A short Arial 10 pt note with a run-format revision fails v14 PDF export with the footnote clipping error. Word presentation exports it. |
| `note-arial20-control` | Removing only the run-format revision marker permits default-presentation export. |
| `revisions-strict` | Native accept/reject leave the insertion, deletion and deleted text unresolved. |
| `revisions-transitional` | Equivalent content resolves to `Indigo prism` on accept and `Amber prism` on reject. |
| `footer-deleted` | Deleted PAGE/NUMPAGES instructions lose dynamic-field markers; both exported pages retain `Folio 1 / 2`. |
| `footer-control` | Dynamic markers survive, and the exported footers are `Folio 1 / 2` and `Folio 2 / 2`. |

Native behavior was checked with v14 commit
`53080d054565cd85e83115880a7a800b73992842` and main commit
`2b480288390c0df89ca1b3893cfba8cb4f670847`. Export observations use
`docxodus@14.0.0`, `@docxodus/export@14.0.0`, and Chrome 155.0.8059.84
on Linux. Main's browser runtime was not rebuilt for these observations.
The current-main native checks use the compiled current-main assembly.

The inputs have zero Open XML SDK 3.5.1 Office 2021 validation errors.
LibreOffice renders the drawings, both revision syntaxes, tab paragraphs, and
two-page footer examples. Its deleted-field footer advances to `Folio 2 / 2`
on page two. This is a LibreOffice observation; no Microsoft Word execution is
claimed.

## Recreate the inputs

From this directory, with Python 3:

```sh
python3 create_inputs.py
```

Only Python's standard library is required. `input-sha256.json` identifies the
committed DOCX bytes. ZIP timestamps can differ on regeneration; XML member
contents are the reproducible contract.

## Run the native checks

With .NET 10 and a checkout of the desired Docxodus version:

```sh
python3 run_native.py /absolute/path/to/Docxodus /tmp/docx-repro-native
```

The runner builds against that checkout, prints the loaded assembly version,
and writes validation reports, converted HTML, and resolved revision documents.
Each operation records success or the full exception in `results.jsonl`.

## Run PDF exports

Install the two v14 packages in a separate working directory. Install a compatible
Chromium or set `DOCXODUS_CHROMIUM` to an existing Chrome/Chromium executable.
Run this from the reproduction directory, substituting the installed module path:

```sh
export DOCXODUS_CHROMIUM=/absolute/path/to/chrome
node export_inputs.mjs \
  /absolute/path/to/node_modules/@docxodus/export/dist/index.js \
  /tmp/docx-repro-pdf \
  tabs-0 tabs-12 tabs-720 \
  shape-plain shape-deleted group-plain group-deleted \
  note-arial20 note-arial20-control footer-control footer-deleted
```

Use an empty output directory for each run: the API refuses to overwrite outputs.
The note is tested with both presentation options. Other cases use
`reviewProfile: "markup"`, `commentProfile: "hidden"`, and the default Docxodus
presentation. Each case writes a status JSON; successful cases also produce
PDF, HTML, PageMap, and render-report files.

The drawing examples are valid simple rectangles, not text boxes or missing-image
packages. The tab examples differ only in the automatic interval. The revision
examples differ only in Strict versus Transitional namespace/relationship URIs.
The note control differs only by removal of `w:rPrChange`.
