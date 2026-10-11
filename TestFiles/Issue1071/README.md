Synthetic inputs linked from issue #1071, copied unchanged from
[dcd139d](https://github.com/JSv4/Docxodus/tree/dcd139d079889054fb3b44193836432068efb74b/repros/synthetic-tables).
The linked `Program.cs` generates these packages from literal XML and invented canal register text.

`DocxDiffDeletedCellGeometryTests` covers all fourteen cases: four horizontal merge/cell removal
failures and ten controls for width changes, row insertion/deletion, horizontal splits,
vertical merges/splits, and grid offsets. It compares both endpoint grids, row offsets and offset
widths, cell widths/types, spans, merge state, and text. Both inputs and all three generated
packages must validate against the Office 2019 schema.
