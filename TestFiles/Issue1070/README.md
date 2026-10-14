Synthetic inputs linked from issue #1070, copied unchanged from
[dcd139d](https://github.com/JSv4/Docxodus/tree/dcd139d079889054fb3b44193836432068efb74b/repros/synthetic-formatting).
The linked `Program.cs` generates these packages from literal XML and invented station text.

`DocxDiffInheritedPresentationTests` covers the plain control and all four failures:
unused numbering, numbered body paragraphs, changed tab settings, and numbered table paragraphs.
It checks accepted/rejected text, effective fonts and spacing per character (including list labels),
and Office 2019 schema validation of both inputs and all three generated packages.
