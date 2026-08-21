# Corpus provenance

#466 requires that corpus provenance and redistribution permissions are recorded. For this
corpus the answer is short, by construction.

## Every input is generated from source

`eval/fixtures/*.json` are build scripts, not documents. Each is replayed over a blank document
produced by `DocxSessionOps.CreateBlankDocx()` at run time, so no `.docx` bytes are committed and
nothing in this corpus originates outside the repository.

| Fixture | Origin | Rights |
|---------|--------|--------|
| `master-services-agreement` | Written for this repository | Same MIT license as the rest of the repository |

## The content is fictional

`master-services-agreement` is a short commercial agreement written to exercise specific engine
behaviour, not a real contract and not derived from one. "Acme Holdings LLC", "Northwind
Consulting Group", "Dana Reyes", and "Sam Okafor" are invented; the fee figures are arbitrary. It
is not legal drafting and must not be reused as a template.

Its shape is chosen to make careless edits visible:

- A defined term (**Service Provider**) used at three separate occurrences, so a scenario can
  require changing exactly one of them.
- A notice period written as **thirty (30) days**, so an amendment has to update both the word
  and the digits.
- A fee table with three adjacent year rows, so a cell edit that catches a neighbour is detectable.
- A signature block and a running header/footer, so an edit that disturbs section or running-story
  state shows up as a package-part delta rather than passing unnoticed.

## Adding a fixture

A new fixture must either be generated the same way, or arrive with its source and license
recorded in the table above. If a real-world document is ever added, its redistribution permission
belongs in this file *before* the bytes land — not afterwards.
