Independent synthetic comparison reproductions

All input packages are generated from literal XML and newly invented text by generate.py.
The generator reads no document, template, corpus, image, or external asset.
No renderer or reference application is needed for the structural assertions.

From the repository root:

  python3 repros/synthetic-presentation/generate.py repros/synthetic-presentation/fixtures
  dotnet run --project repros/synthetic-presentation/Repro.csproj -c Release -- repros/synthetic-presentation/fixtures

To run one case, append its name after the fixture directory, for example:

  dotnet run --project repros/synthetic-presentation/Repro.csproj -c Release -- repros/synthetic-presentation/fixtures default-style-id

The project uses internal alignment diagnostics under the existing Docxodus.Tests friend assembly name.
It calls the public DocxCompare.Compare and RevisionProcessor APIs for comparison and accepted/rejected output.
Results are saved in fixtures/results.json (or NAME-results.json) and fixtures/generated/.
Exit code 1 means at least one desired assertion failed; exit code 0 means the selected cases passed.
A failing reproduction exits 1 intentionally so it can become a TDD regression test.

Verified on main 2d71690bad160d91c03c3a6e27c5dae935d97c34, .NET SDK10.0.301,
DocumentFormat.OpenXml3.5.1, Office2019 validation.
All 55 packages (22 inputs and33 compared/accepted/rejected outputs) validate cleanly.
All accepted/rejected paragraph-text checks pass.

Failing cases:
  default-style-id: accepted font sizes/spacing remain original when default paragraph style IDs differ.
  named-style-id: same failure with explicit paragraph style references.
  defaults-with-table: accepted document defaults remain original when the document contains a table.
  inserted-heading-gap: the inserted heading pairs with the old relay body paragraph; the edited body is inserted separately.
  table-style-cell-margins: accepted table-style margins remain original.

Passing controls:
  same-style-id-control
  defaults-plain-control
  body-edits-control
  table-direct-margins-control
  section-margin-control
  paragraph-spacing-control

Assertion scope:
The format inspector follows the explicit style chain and document defaults and compares the declared
literal font slots, font sizes, and paragraph spacing used by these fixtures. It is not a general
font/theme resolver; unrelated synthesized metadata is deliberately excluded from those assertions.
The table fixture has one direct table style and no conditional style or table-style parent.
The heading case specifies a desired matching policy; accepted/rejected text is already correct.

The source fixtures, observations, controls, and driver are scaffolding for focused production tests.
Implementation work should preserve native revision semantics, schema validity, and both endpoints.
